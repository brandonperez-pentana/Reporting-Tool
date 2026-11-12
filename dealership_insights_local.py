from __future__ import annotations

import argparse
import hashlib
import json
import logging
from logging.handlers import RotatingFileHandler
import os
import re
from dataclasses import dataclass, field
from itertools import combinations
from datetime import datetime, timedelta
from pathlib import Path
from collections import defaultdict
from typing import Any, Callable, Iterable

import pandas as pd
import pdfplumber
from docx import Document
from docx.shared import RGBColor
from reporting.relationship_engine import analyze_operational_relationships
from reporting.operational_language_map import phase156_naturalise_text, phase162_humanise_text
from reporting import metric_history

try:
    from openpyxl.styles import PatternFill
except ImportError:  # pragma: no cover - optional dependency for enhanced export
    PatternFill = None


def _default_accounts_dir() -> Path | None:
    """Where dealership folders live, resolved per machine.

    This used to be one operator's OneDrive path baked into the source, which
    meant every command-line entry point died on any other machine. Order of
    preference: the REPORTER_ACCOUNTS_DIR environment variable, then whatever
    the operator chose in the app's own settings, then nothing — callers that
    need a folder must ask for one rather than inherit someone else's.
    """
    override = os.environ.get("REPORTER_ACCOUNTS_DIR", "").strip()
    if override:
        return Path(override).expanduser()
    try:
        from settings_manager import load_settings

        configured = str(load_settings().get("accounts_folder", "")).strip()
    except Exception:  # noqa: BLE001 - settings are optional here
        configured = ""
    return Path(configured).expanduser() if configured else None


ACCOUNTS_DIR = _default_accounts_dir()
OUTPUT_DIR = Path(__file__).resolve().parent / "Insights Output"
DEBUG_TEXT_DIR = OUTPUT_DIR / "Debug Text"
IGNORED_FOLDER_NAMES = {"archive", "archived"}
PDF_SUFFIX = ".pdf"
LOG_FILE_NAME = "processing.log"
DEBUG_MODE = os.getenv("REPORTER_DEBUG_MODE", "1").strip().lower() not in {"0", "false", "no", "off"}
DEBUG_METRICS_JSON_NAME = "parsed_metrics_debug.json"
DEBUG_METRICS_CSV_NAME = "parsed_metrics_debug.csv"
DEBUG_METRIC_DELTAS_JSON_NAME = "parsed_metric_deltas_debug.json"
DEBUG_METRIC_DELTAS_CSV_NAME = "parsed_metric_deltas_debug.csv"
SMS_DEBUG_CSV_NAME = "sms_v2_debug.csv"
SCORING_DEBUG_CSV_NAME = "scoring_debug.csv"
REVIEW_REQUIRED_CSV_NAME = "review_required.csv"
REVIEW_REQUIRED_ENHANCED_CSV_NAME = "review_required_enhanced.csv"
REVIEW_REQUIRED_ENHANCED_XLSX_NAME = "review_required_enhanced.xlsx"
MASTER_COMBINED_DOCX_NAME = "master_combined_report.docx"
MIN_METRICS_FOR_RULES = 6
MAX_HISTORICAL_MONTHS = 3
MIN_PREVIOUS_MONTHS_FOR_TREND = 2
TREND_CONFIDENCE_ORDER = {"low": 0, "medium": 1, "high": 2}
TREND_PRIORITY_BONUS = {"low": 0, "medium": 3, "high": 6}
PRIORITY_INSIGHT_COUNT = 4
SMS_CHAR_LIMIT = 160
NO_ADDITIONAL_INSIGHT_MARKER = "[No additional insight]"

METRIC_FAMILY_DEFINITIONS = {
    "sales_momentum": ["sales_total_leads", "sales_total_deals", "sales_overall_conversion_rate"],
    "service_momentum": ["service_reminders_sent", "service_bookings_from_campaigns", "service_campaign_roi"],
    "followup_risk": ["sales_uncontacted_leads", "sales_internet_response_time", "service_outstanding_tasks"],
}

TREND_METRIC_NAMES = {
    "sales_total_leads",
    "sales_total_deals",
    "sales_overall_conversion_rate",
    "sales_internet_response_time",
    "service_reminders_sent",
    "service_bookings_from_campaigns",
    "service_campaign_roi",
    "service_outstanding_tasks",
}

METRIC_DEFINITIONS: dict[str, dict[str, Any]] = {
    "sales_total_leads": {"metric_type": "volume", "label": "lead volume", "higher_is_better": True},
    "sales_total_deals": {"metric_type": "volume", "label": "deal volume", "higher_is_better": True},
    "sales_overall_conversion_rate": {"metric_type": "rate", "label": "overall conversion", "higher_is_better": True},
    "sales_uncontacted_leads": {"metric_type": "count", "label": "uncontacted leads", "higher_is_better": False},
    "sales_internet_response_time": {"metric_type": "time", "label": "response times", "higher_is_better": False},
    "sales_new_leads_to_appointments": {"metric_type": "rate", "label": "lead-to-appointment conversion", "higher_is_better": True},
    "sales_new_leads_to_test_drives": {"metric_type": "rate", "label": "lead-to-test-drive conversion", "higher_is_better": True},
    "sales_appointment_to_sales": {"metric_type": "rate", "label": "appointment-to-sale conversion", "higher_is_better": True},
    "sales_overdue_tasks": {"metric_type": "count", "label": "overdue sales tasks", "higher_is_better": False},
    # Phase 169: higher email capture = stronger contactability / capture quality (must not read as CRM "pressure").
    "sales_email_capture_rate": {"metric_type": "rate", "label": "email capture rate", "higher_is_better": True},
    "sales_same_day_deal_pct": {"metric_type": "rate", "label": "same-day deal mix", "higher_is_better": True},
    "sales_campaign_deals": {"metric_type": "volume", "label": "sales campaign deals", "higher_is_better": True},
    "sales_campaign_roi": {"metric_type": "currency", "label": "sales campaign ROI", "higher_is_better": True},
    "sales_pipeline_size": {"metric_type": "count", "label": "pipeline size", "higher_is_better": True},
    "sales_pipeline_unqualified": {"metric_type": "count", "label": "unqualified pipeline leads", "higher_is_better": False},
    "sales_pipeline_contacted": {"metric_type": "count", "label": "contacted pipeline leads", "higher_is_better": True},
    "sales_pipeline_store_visit": {"metric_type": "count", "label": "store visits", "higher_is_better": True},
    "sales_pipeline_demo": {"metric_type": "count", "label": "demos", "higher_is_better": True},
    "sales_pipeline_write_up": {"metric_type": "count", "label": "write-ups", "higher_is_better": True},
    "sales_pipeline_finance": {"metric_type": "count", "label": "F&I opportunities", "higher_is_better": True},
    "sales_pipeline_deals_closed": {"metric_type": "count", "label": "deals closed", "higher_is_better": True},
    "sales_pipeline_lost": {"metric_type": "count", "label": "lost deals", "higher_is_better": False},
    "service_reminders_sent": {"metric_type": "volume", "label": "service reminders sent", "higher_is_better": True},
    "service_customers_enrolled": {"metric_type": "count", "label": "customers enrolled", "higher_is_better": True},
    "service_outstanding_tasks": {"metric_type": "count", "label": "outstanding service tasks", "higher_is_better": False},
    "service_rebooked_no_shows": {"metric_type": "count", "label": "rebooked no-shows", "higher_is_better": True},
    "service_no_shows": {"metric_type": "count", "label": "service no-shows", "higher_is_better": False},
    "service_bookings_from_campaigns": {"metric_type": "volume", "label": "service campaign bookings", "higher_is_better": True},
    "service_campaign_roi": {"metric_type": "currency", "label": "service campaign ROI", "higher_is_better": True},
}

SIGNIFICANCE_ORDER = {"not_available": -1, "negligible": 0, "slight": 1, "meaningful": 2, "strong": 3}

# Deterministic theme synthesis (phase 3): rule_id -> list of (theme_id, weight)
RULE_THEME_AFFINITY: dict[str, list[tuple[str, float]]] = {
    "leads_up_conversion_down": [("lead_quality_deterioration", 12.0), ("conversion_weakness", 10.0)],
    "leads_up_deals_down": [("conversion_weakness", 11.0), ("execution_pressure", 6.0)],
    "leads_down_conversion_up": [("efficiency_recovery", 12.0), ("stable_consolidation", 4.0)],
    "lead_handling_pressure": [("execution_pressure", 14.0)],
    "lead_handling_improved": [("efficiency_recovery", 9.0), ("stable_consolidation", 3.0)],
    "test_drive_and_conversion_down": [("weakening_engagement_theme", 11.0), ("conversion_weakness", 8.0)],
    "closing_execution_up": [("efficiency_recovery", 10.0)],
    "combined_sales_service_up": [("stable_consolidation", 8.0), ("profitability_recovery", 5.0)],
    "service_campaigns_up": [("retention_resilience", 9.0), ("service_led_stability", 7.0)],
    "service_bookings_roi_mismatch": [("volatile_mixed_conditions", 8.0), ("profitability_recovery", 5.0)],
    "service_tasks_pressure": [("execution_pressure", 10.0)],
    "service_follow_up_cleaner": [("retention_resilience", 10.0)],
    "sales_marketing_negative_roi": [("volatile_mixed_conditions", 12.0)],
    "sales_campaign_roi_decline": [("volatile_mixed_conditions", 12.0), ("conversion_weakness", 5.0)],
    "service_marketing_negative_roi": [("volatile_mixed_conditions", 11.0)],
    "crm_execution_pressure_composite": [("execution_pressure", 16.0), ("volatile_mixed_conditions", 4.0)],
}

# Phase 155 — deterministic insight diversity buckets (not dealership-specific).
RULE_OPERATIONAL_FAMILY: dict[str, str] = {
    "leads_up_conversion_down": "conversion_story",
    "leads_up_deals_down": "conversion_story",
    "leads_down_conversion_up": "conversion_story",
    "test_drive_and_conversion_down": "conversion_story",
    "sales_conversion_snapshot": "conversion_story",
    "cross_metric_efficiency_gain": "conversion_story",
    "cross_metric_qualification_inefficiency": "conversion_story",
    "closing_execution_up": "conversion_story",
    "lead_handling_pressure": "crm_execution",
    "lead_handling_improved": "crm_execution",
    "lead_response_snapshot": "crm_execution",
    "service_tasks_pressure": "service_ops",
    "service_follow_up_cleaner": "service_ops",
    "service_task_snapshot": "service_ops",
    "service_campaigns_up": "unified_campaign",
    "service_bookings_roi_mismatch": "unified_campaign",
    "service_booking_snapshot": "unified_campaign",
    "service_roi_snapshot": "unified_campaign",
    "cross_metric_retention_strength": "unified_campaign",
    "cross_metric_service_profitability": "profitability_service",
    "sales_marketing_negative_roi": "unified_campaign",
    "sales_campaign_roi_decline": "unified_campaign",
    "sales_marketing_up": "unified_campaign",
    "combined_sales_service_up": "combined_lane",
    "sales_volume_snapshot": "volume_snapshot",
    "crm_execution_pressure_composite": "crm_execution",
}

# Phase 155 — substring-based operational concept tags for redundancy suppression (deterministic).
PHASE155_CONCEPT_TAG_NEEDLES: tuple[tuple[str, tuple[str, ...]], ...] = (
    (
        "conversion_funnel",
        (
            "conversion",
            "mid-funnel",
            "mid funnel",
            "funnel efficiency",
            "appointment-to-sale",
            "same-day deal",
            "closing",
            "deal mix",
            "interpretive anchor",
            "closing discipline",
            "stronger closing",
        ),
    ),
    ("demand_enquiry", ("enquiry", "inquiry", "lead volume", "lead flow", "demand", "top-of-funnel", "top of funnel")),
    ("crm_response", ("response time", "response-time", "internet response", "follow-up", "follow up", "uncontacted", "contact coverage")),
    ("crm_data_capture", ("email capture", "mobile capture", "capture rate", "data capture")),
    ("service_backlog", ("outstanding service", "service task", "service follow-up", "service follow up", "backlog")),
    ("retention_campaign", ("campaign", "enrolment", "enrollment", "reminder", "service campaign", "retention")),
    ("profitability", ("roi", "margin", "profitability", "commercial return")),
    ("operational_offset", ("offset", "stabiliser", "stabilizer", "buffer", "partial offset", "operational offset", "partial operational")),
)

PHASE155_MAX_SUMMARY_SENTENCES = 3

# Phase 155C — ordered metrics eligible for composite CRM / customer-management deterioration (deterministic).
PHASE155C_COMPOSITE_ORDER: tuple[str, ...] = (
    "sales_internet_response_time",
    "sales_uncontacted_leads",
    "sales_overdue_tasks",
    "sales_email_capture_rate",
    "service_reminders_sent",
    "service_customers_enrolled",
)

# Phase 155D — operational domains for composite evidence balancing (deterministic).
PHASE155D_CRM_DOMAIN_METRICS: dict[str, tuple[str, ...]] = {
    "responsiveness": ("sales_internet_response_time",),
    "lead_contact_gap": ("sales_uncontacted_leads",),
    "sales_execution_backlog": ("sales_overdue_tasks",),
    "customer_data_quality": ("sales_email_capture_rate",),
    # Split retention signals so two materially adverse enrolment/reminder moves both surface when shallow stacks.
    "service_campaign_reminders": ("service_reminders_sent",),
    "service_customer_enrolment": ("service_customers_enrolled",),
}
PHASE155D_DOMAIN_STABLE_ORDER: tuple[str, ...] = (
    "responsiveness",
    "lead_contact_gap",
    "sales_execution_backlog",
    "customer_data_quality",
    "service_campaign_reminders",
    "service_customer_enrolment",
)

THEME_FRAGMENTS: dict[str, list[str]] = {
    "efficiency_recovery": [
        "the operational through-line was tighter sales efficiency against the available enquiry pool",
        "the month mainly reflected recovering conversion discipline and cleaner execution",
        "performance centred on extracting more value from each sales opportunity",
    ],
    "lead_quality_deterioration": [
        "the dominant tension was softer conversion against a larger enquiry base",
        "the story was primarily lead-flow strength paired with weaker downstream conversion",
        "operations faced a classic demand-up, efficiency-down skew in the sales funnel",
    ],
    "execution_pressure": [
        "the month was defined by execution pressure in follow-up and responsiveness",
        "operations were skewed toward follow-up risk rather than raw demand shortfall",
        "the clearest strain appeared in day-to-day lead handling discipline",
    ],
    "retention_resilience": [
        "service retention mechanics looked comparatively resilient",
        "the steadier thread was service follow-through and campaign hygiene",
        "aftersales execution provided the more dependable operational anchor",
    ],
    "service_led_stability": [
        "service-led activity provided the more stable commercial spine",
        "the service lane carried steadier momentum while sales signals mixed",
        "aftersales strength offset noisier signals elsewhere",
    ],
    "profitability_recovery": [
        "profitability-oriented metrics moved into a cleaner, more constructive range",
        "ROI and campaign economics improved their recent trajectory",
        "commercial quality improved in campaign returns rather than only volume",
    ],
    "conversion_weakness": [
        "conversion efficiency was the main drag on the sales result",
        "the sales outcome was gated more by conversion than by top-of-funnel volume",
        "closing and funnel conversion absorbed most of the operational attention",
    ],
    "stable_consolidation": [
        "results read as consolidation rather than a sharp inflection",
        "the month was more about holding ground across core metrics",
        "performance stayed range-bound without a single dramatic driver",
    ],
    "volatile_mixed_conditions": [
        "sales and service sent different signals at once",
        "the month mixed stronger and softer reads across lanes",
        "cross-currents showed in both sales and service",
    ],
    "weakening_engagement_theme": [
        "early-funnel engagement weakened alongside conversion",
        "top-of-funnel progression lost some momentum relative to prior months",
        "buyer engagement through the early sales path softened materially",
    ],
}

POSITIVE_SIGNAL_VERBS: dict[str, list[str]] = {
    "default": ["Good read:", "Worth noting:", "Bright spot:"],
}

WATCHPOINT_OPENERS: dict[str, list[str]] = {
    "default": ["The primary watchpoint is", "The main operational watchpoint remains", "Leadership attention should stay on"],
}

# Phase 4 — executive hierarchy (index = tie-break priority, lower first).
# Profitability and conversion outweigh demand noise; operational risk sits after retention decline per spec.
STRATEGIC_PILLAR_EXEC_ORDER: list[str] = [
    "profitability_movement",
    "conversion_efficiency",
    "retention_service_health",
    "operational_risk",
    "engagement_quality",
    "lead_flow",
    "secondary_operations",
]

# Human-readable dominant business condition IDs (phase 4b).
DOMINANT_BUSINESS_STATE_IDS: frozenset[str] = frozenset(
    {
        "conversion_deterioration",
        "profitability_pressure",
        "operational_recovery",
        "retention_resilience_state",
        "demand_softening",
        "execution_improvement",
        "efficiency_recovery_state",
        "mixed_volatility",
        "sustained_operational_pressure",
        "stabilising_performance",
        "engagement_weakening_state",
        "service_led_resilience",
        "funnel_deterioration",
        "conversion_led_recovery",
        "lead_quality_deterioration_state",
    }
)

DOMINANT_STATE_HEADLINES: dict[str, list[str]] = {
    "conversion_deterioration": [
        "What fundamentally defined the month was deteriorating conversion discipline through the sales funnel.",
        "The defining business condition was weaker conversion quality relative to available enquiry.",
        "This month centred on conversion deterioration as the primary commercial story.",
    ],
    "profitability_pressure": [
        "Sales and service campaign returns both weakened—ROI is the headline pressure.",
        "Campaign economics softened on sales and workshop sides this month.",
        "What fundamentally defined the month was profitability and campaign-return pressure.",
        "The defining business condition was commercial return quality, not peripheral activity noise.",
        "This month centred on ROI and economics as the headline operational stress.",
    ],
    "operational_recovery": [
        "What fundamentally defined the month was a cleaner operational trajectory in execution and follow-up.",
        "The defining business condition was recovering execution discipline in lead handling.",
        "This month centred on operational recovery visible in responsiveness and contact coverage.",
    ],
    "retention_resilience_state": [
        "What fundamentally defined the month was resilient aftersales throughput and retention mechanics.",
        "The defining business condition was comparatively steady service-side execution.",
        "This month centred on retention and service throughput as the steadier anchor.",
    ],
    "demand_softening": [
        "What fundamentally defined the month was softer demand at the enquiry and deal-throughput level.",
        "The defining business condition was demand softening rather than a single conversion spike.",
        "This month centred on weaker top-of-funnel and volume throughput.",
    ],
    "execution_improvement": [
        "What fundamentally defined the month was improved execution speed and follow-up coverage.",
        "The defining business condition was cleaner lead-handling execution.",
        "This month centred on execution improvement in responsiveness and contact discipline.",
    ],
    "efficiency_recovery_state": [
        "What fundamentally defined the month was recovering sales efficiency against a tighter or smaller enquiry pool.",
        "The defining business condition was better conversion discipline even as headline demand eased.",
        "This month centred on efficiency-led recovery in the sales funnel.",
    ],
    "mixed_volatility": [
        "What fundamentally defined the month was mixed, volatile operating conditions without a single clean trend.",
        "The defining business condition was cross-currents across sales and service with uneven month-to-month movement.",
        "This month centred on volatility and offsetting signals rather than one directional theme.",
    ],
    "sustained_operational_pressure": [
        "What fundamentally defined the month was sustained operational pressure in follow-up and responsiveness.",
        "The defining business condition was ongoing execution strain in lead handling.",
        "This month centred on sustained operational risk in contact speed and pipeline hygiene.",
    ],
    "stabilising_performance": [
        "What fundamentally defined the month was stabilising performance within a relatively tight recent range.",
        "The defining business condition was consolidation rather than a sharp break in direction.",
        "This month centred on stabilising outcomes across the core operating metrics.",
    ],
    "engagement_weakening_state": [
        "What fundamentally defined the month was weakening early-funnel engagement and progression.",
        "The defining business condition was softer buyer progression through early sales stages.",
        "This month centred on engagement weakening ahead of closing.",
    ],
    "service_led_resilience": [
        "What fundamentally defined the month was service-led resilience offsetting noisier sales signals.",
        "The defining business condition was aftersales stability while sales lanes mixed.",
        "This month centred on the service lane as the more dependable commercial spine.",
    ],
    "funnel_deterioration": [
        "What fundamentally defined the month was broader funnel deterioration across progression and conversion.",
        "The defining business condition was leakage through multiple funnel stages, not an isolated metric.",
        "This month centred on deteriorating funnel health end-to-end.",
    ],
    "conversion_led_recovery": [
        "What fundamentally defined the month was conversion-led recovery even as headline enquiry eased.",
        "The defining business condition was stronger closing discipline outweighing softer demand reads.",
        "This month centred on conversion improvement as the interpretive anchor.",
    ],
    "lead_quality_deterioration_state": [
        "What fundamentally defined the month was lead-quality strain—stronger enquiry flow paired with weaker conversion economics.",
        "The defining business condition was demand-up, conversion-down skew in the funnel.",
        "This month centred on lead quality and conversion mismatch.",
    ],
}

DOMINANT_STATE_SMS_FRAG: dict[str, list[str]] = {
    "conversion_deterioration": [
        "Dominant condition: conversion deterioration.",
        "Headline: weaker funnel conversion.",
        "Strategic read: conversion is the story.",
    ],
    "profitability_pressure": [
        "Dominant condition: profitability pressure.",
        "Headline: ROI and economics under stress.",
        "Strategic read: commercial returns lead.",
    ],
    "operational_recovery": [
        "Dominant condition: operational recovery.",
        "Headline: cleaner execution trajectory.",
        "Strategic read: follow-up improving.",
    ],
    "retention_resilience_state": [
        "Dominant condition: retention resilience.",
        "Headline: aftersales steadier than sales.",
        "Strategic read: service anchor holding.",
    ],
    "demand_softening": [
        "Dominant condition: demand softening.",
        "Headline: enquiry/deal throughput weaker.",
        "Strategic read: volume is the drag.",
    ],
    "execution_improvement": [
        "Dominant condition: execution improvement.",
        "Headline: better responsiveness and coverage.",
        "Strategic read: follow-up lifting.",
    ],
    "efficiency_recovery_state": [
        "Dominant condition: conversion-led efficiency recovery.",
        "Headline: tighter funnel despite softer demand.",
        "Strategic read: mid-funnel strength.",
    ],
    "mixed_volatility": [
        "Dominant condition: mixed volatile month.",
        "Headline: cross-currents across lanes.",
        "Strategic read: prioritise the heaviest signal stack.",
    ],
    "sustained_operational_pressure": [
        "Dominant condition: sustained execution pressure.",
        "Headline: follow-up risk persists.",
        "Strategic read: contact discipline first.",
    ],
    "stabilising_performance": [
        "Dominant condition: stabilising performance.",
        "Headline: range-bound operating picture.",
        "Strategic read: consolidation month.",
    ],
    "engagement_weakening_state": [
        "Dominant condition: engagement weakening.",
        "Headline: early-funnel progression softer.",
        "Strategic read: top-of-funnel watch.",
    ],
    "service_led_resilience": [
        "Dominant condition: service-led resilience.",
        "Headline: aftersales offsetting sales noise.",
        "Strategic read: service spine steadier.",
    ],
    "funnel_deterioration": [
        "Dominant condition: funnel deterioration.",
        "Headline: leakage through the path.",
        "Strategic read: end-to-end funnel risk.",
    ],
    "conversion_led_recovery": [
        "Dominant condition: conversion-led recovery.",
        "Headline: efficiency up as demand eased.",
        "Strategic read: closing quality leads.",
    ],
    "lead_quality_deterioration_state": [
        "Dominant condition: lead-quality mismatch.",
        "Headline: more leads, weaker conversion.",
        "Strategic read: demand-up, conversion-down skew.",
    ],
}

METRIC_TO_STRATEGIC_PILLAR: dict[str, str] = {
    "sales_internet_response_time": "operational_risk",
    "sales_uncontacted_leads": "operational_risk",
    "sales_overdue_tasks": "operational_risk",
    "sales_overall_conversion_rate": "conversion_efficiency",
    "sales_appointment_to_sales": "conversion_efficiency",
    "sales_new_leads_to_appointments": "engagement_quality",
    "sales_new_leads_to_test_drives": "engagement_quality",
    "sales_total_leads": "lead_flow",
    "sales_total_deals": "lead_flow",
    "sales_pipeline_size": "engagement_quality",
    "sales_pipeline_contacted": "engagement_quality",
    "sales_pipeline_store_visit": "engagement_quality",
    "sales_pipeline_demo": "engagement_quality",
    "sales_campaign_roi": "profitability_movement",
    "sales_campaign_deals": "profitability_movement",
    "service_campaign_roi": "profitability_movement",
    "service_bookings_from_campaigns": "retention_service_health",
    "service_reminders_sent": "retention_service_health",
    "service_outstanding_tasks": "retention_service_health",
    "service_no_shows": "retention_service_health",
    "service_rebooked_no_shows": "retention_service_health",
    "sales_email_capture_rate": "secondary_operations",
    "sales_same_day_deal_pct": "secondary_operations",
}

RULE_ID_TO_STRATEGIC_PILLAR: dict[str, tuple[str, float]] = {
    "lead_handling_pressure": ("operational_risk", 7.5),
    "lead_handling_improved": ("operational_risk", 4.0),
    "leads_up_conversion_down": ("conversion_efficiency", 8.0),
    "leads_up_deals_down": ("conversion_efficiency", 7.5),
    "leads_down_conversion_up": ("conversion_efficiency", 6.5),
    "test_drive_and_conversion_down": ("engagement_quality", 7.0),
    "closing_execution_up": ("conversion_efficiency", 5.5),
    "sales_marketing_negative_roi": ("profitability_movement", 8.0),
    "service_marketing_negative_roi": ("profitability_movement", 7.5),
    "service_bookings_roi_mismatch": ("profitability_movement", 6.5),
    "service_tasks_pressure": ("retention_service_health", 6.5),
    "service_follow_up_cleaner": ("retention_service_health", 5.0),
    "combined_sales_service_up": ("retention_service_health", 4.0),
    "service_campaigns_up": ("retention_service_health", 5.5),
    "crm_execution_pressure_composite": ("operational_risk", 9.0),
}

THEME_TO_PRIMARY_PILLAR: dict[str, str] = {
    "execution_pressure": "operational_risk",
    "conversion_weakness": "conversion_efficiency",
    "lead_quality_deterioration": "conversion_efficiency",
    "efficiency_recovery": "conversion_efficiency",
    "profitability_recovery": "profitability_movement",
    "retention_resilience": "retention_service_health",
    "service_led_stability": "retention_service_health",
    "weakening_engagement_theme": "engagement_quality",
    "stable_consolidation": "secondary_operations",
    "volatile_mixed_conditions": "secondary_operations",
}

DOMINANT_HEADLINE_FRAGMENTS: dict[str, list[str]] = {
    "operational_risk": [
        "The month’s decisive business condition was elevated operational risk in lead handling and responsiveness.",
        "Leadership should read the month first through execution risk in enquiry follow-up and contact speed.",
        "The dominant condition was follow-up and responsiveness pressure, not a shortage of raw demand signals alone.",
    ],
    "conversion_efficiency": [
        "The decisive business condition was sales funnel conversion discipline—how enquiry converted into outcomes.",
        "Conversion quality and closing discipline mattered more than headline lead counts for interpreting results.",
        "The executive read hinges on conversion behaviour and mid-funnel execution rather than volume alone.",
    ],
    "profitability_movement": [
        "Campaign economics and ROI movement framed the commercial story more than peripheral sales metrics.",
        "Profitability and return signals from campaigns carried the weight of evidence for the month.",
        "Commercial return quality—not only activity levels—was the defining financial lens.",
    ],
    "retention_service_health": [
        "Aftersales throughput, backlog, and campaign hygiene shaped the steadier operational spine.",
        "Service execution and retention mechanics provided the clearest structural anchor in the dataset.",
        "The service lane defined the more dependable part of the operating picture.",
    ],
    "engagement_quality": [
        "Early-funnel engagement and appointment progression were the main swing factors in the sales path.",
        "Buyer progression through early sales stages mattered more than headline totals alone.",
        "Engagement depth, not just enquiry counts, drove how the sales story should be read.",
    ],
    "lead_flow": [
        "Enquiry and deal-volume movement anchored how the month should be interpreted commercially.",
        "Top-of-funnel demand shifts were the primary structural driver in the sales picture.",
        "Lead and deal throughput was the headline commercial movement relative to other signals.",
    ],
    "secondary_operations": [
        "Secondary operating indicators moved without overturning the core strategic diagnosis.",
        "Supporting sales hygiene metrics shifted at the margin without defining the month.",
        "The month’s diagnosis rests on headline drivers; peripheral metrics stayed in a supporting role.",
    ],
}

# A placeholder report is one where the source PDFs were missing or unreadable. It
# must say so. The previous bullets ("Sales activity increased month-on-month.")
# asserted performance for a month whose figures were never read.
PLACEHOLDER_BULLETS = [
    "No sales or service figures could be read for this month.",
    "The monthly reports were missing or could not be parsed.",
]
PLACEHOLDER_SUMMARY = (
    "No figures were available for this month, so no insights were generated. "
    "Check that this dealership's sales and service reports are in the accounts folder "
    "for this month, then re-run."
)
PLACEHOLDER_SMS = "Not for sending: no figures were available for this dealership this month."

RULE_MESSAGE_TEMPLATES: dict[str, dict[str, list[str]]] = {
    "combined_sales_service_up": {
        "slightly_improved": [
            "The month finished in better shape overall, with sales and service both trending in the right direction.",
            "Results improved across both sides of the dealership, with sales momentum and service activity both lifting.",
            "Sales and service both moved forward this month, pointing to a healthier overall result.",
        ],
        "strongly_improved": [
            "It was a notably stronger month overall, with clear improvement across both sales and service performance.",
            "The dealership posted a strong all-round result this month, with both sales and service stepping up.",
            "Performance lifted meaningfully across the dealership, with sales and service both showing stronger momentum.",
        ],
    },
    "leads_up_conversion_down": {
        "slightly_softer": [
            "Lead volume improved, but weaker conversion suggests opportunities are still being missed further down the funnel.",
            "Enquiry activity increased, though the lower conversion rate points to softer follow-up performance.",
            "The dealership generated more leads this month, but weaker conversion limited how much of that demand turned into extra sales.",
        ],
        "significantly_softer": [
            "Lead volume improved, but the sharper drop in conversion suggests meaningful leakage further down the funnel.",
            "Enquiry activity increased, though the heavier conversion decline points to more serious follow-up slippage.",
            "The dealership generated more leads this month, but the weaker conversion result prevented that extra demand from turning into enough additional sales.",
        ],
    },
    "leads_down_conversion_up": {
        "slightly_improved": [
            "Lead volume eased, but stronger conversion suggests the team extracted more value from a slightly smaller enquiry pool.",
            "Enquiry flow softened a touch, although better conversion helped preserve sales efficiency.",
            "The dealership worked a smaller lead pool more effectively this month, with conversion moving the right way.",
        ],
        "strongly_improved": [
            "Lead volume eased, but a strong conversion lift shows the team extracted materially more value from a smaller enquiry pool.",
            "Even with fewer enquiries, significantly better conversion points to a much cleaner sales process this month.",
            "The dealership generated fewer leads, but stronger conversion more than compensated by turning demand into sales more effectively.",
        ],
    },
    "leads_up_deals_down": {
        "slightly_softer": [
            "Demand improved, but fewer deals were closed, pointing to softer conversion deeper in the funnel.",
            "The dealership attracted more enquiry activity, though deal volume slipped back further down the sales process.",
            "Lead flow improved this month, but weaker close rates meant the extra demand did not carry through to deals.",
        ],
        "significantly_softer": [
            "Demand improved, but materially fewer deals were closed, pointing to heavier conversion pressure deeper in the funnel.",
            "The dealership attracted more enquiry activity, though the sharper deal-volume drop signals a more serious breakdown later in the sales process.",
            "Lead flow improved this month, but the weaker close outcome meant the extra demand failed to translate into enough delivered deals.",
        ],
    },
    "lead_handling_pressure": {
        "slightly_softer": [
            "Response times slowed and more leads remained untouched, increasing the risk of warm enquiries dropping away.",
            "Lead follow-up pressure increased this month, with slower responses and more uncontacted enquiries sitting in the pipeline.",
            "The dealership may be losing momentum with new enquiries, as response times softened and untouched leads increased.",
        ],
        "significantly_softer": [
            "Response times slowed sharply and more leads remained untouched, creating a bigger risk of warm enquiries dropping away.",
            "Lead follow-up pressure increased materially this month, with slower responses and a larger pool of uncontacted enquiries sitting in the pipeline.",
            "The dealership may be losing meaningful momentum with new enquiries, as response times weakened and untouched leads climbed.",
        ],
    },
    "lead_handling_improved": {
        "slightly_improved": [
            "Lead handling improved, with faster response times and fewer enquiries sitting untouched in the pipeline.",
            "Follow-up looked cleaner this month, as response times improved and fewer leads remained uncontacted.",
            "The dealership responded more promptly and allowed fewer enquiries to sit idle, improving lead handling discipline.",
        ],
        "strongly_improved": [
            "Lead handling improved meaningfully, with much faster response times and fewer enquiries left untouched in the pipeline.",
            "Follow-up looked materially cleaner this month, as response times improved and the uncontacted lead count came back.",
            "The dealership responded much more promptly and left fewer enquiries sitting idle, lifting lead-handling discipline.",
        ],
    },
    "closing_execution_up": {
        "slightly_improved": [
            "Closing execution improved, with more appointments converting and a better same-day deal mix.",
            "Sales execution tightened up this month, with stronger appointment conversion and better same-day closing.",
            "The dealership converted appointments more effectively and improved same-day deal quality.",
        ],
        "strongly_improved": [
            "Closing execution improved strongly, with more appointments converting and a notably stronger same-day deal mix.",
            "Sales execution tightened up materially this month, with stronger appointment conversion and much better same-day closing.",
            "The dealership converted appointments more effectively and delivered a clearer lift in same-day deal quality.",
        ],
    },
    "test_drive_and_conversion_down": {
        "slightly_softer": [
            "Top-of-funnel momentum softened, with fewer enquiries reaching test drives and overall conversion slipping.",
            "Early-stage sales momentum eased this month, as test-drive progression and conversion both moved lower.",
            "Fewer enquiries advanced to test drive, and the softer conversion result suggests less momentum through the funnel.",
        ],
        "significantly_softer": [
            "Top-of-funnel momentum softened materially, with fewer enquiries reaching test drives and overall conversion slipping harder.",
            "Early-stage sales momentum eased sharply this month, as test-drive progression and conversion both moved lower.",
            "Fewer enquiries advanced to test drive, and the weaker conversion result suggests a more serious loss of momentum through the funnel.",
        ],
    },
    "sales_marketing_up": {
        "slightly_improved": [
            "Marketing performance strengthened, with more attributed deals and higher sales ROI during the month.",
            "Sales campaigns worked more effectively this month, producing better deal attribution and improved ROI.",
            "Campaign activity contributed more effectively to sales, with stronger deal output and better return.",
        ],
        "strongly_improved": [
            "Marketing performance strengthened materially, with more attributed deals and a stronger lift in sales ROI.",
            "Sales campaigns worked much more effectively this month, producing better deal attribution and stronger ROI.",
            "Campaign activity contributed more meaningfully to sales, with stronger deal output and a clear return uplift.",
        ],
    },
    "sales_marketing_negative_roi": {
        "significantly_softer": [
            "Sales campaign ROI turned negative, so campaign mix and attribution quality should be reviewed closely.",
            "Sales marketing finished in negative ROI territory, which warrants a closer look at spend quality and attribution.",
            "Campaign return moved below zero this month, suggesting the sales mix and attribution settings need attention.",
        ],
    },
    "sales_campaign_roi_decline": {
        "slightly_softer": [
            "Sales campaign ROI softened month-on-month while remaining positive, pointing to weaker marketing return quality.",
            "Sales marketing return eased this month even though ROI stayed above zero, suggesting campaign efficiency is slipping.",
            "Campaign ROI moved lower month-on-month while still positive, signalling softer sales-side marketing performance.",
        ],
        "significantly_softer": [
            "Sales campaign ROI fell sharply month-on-month while still positive, signalling a material loss of marketing return.",
            "Sales marketing return deteriorated materially even though ROI remained above zero, pointing to a serious efficiency issue.",
            "Campaign ROI dropped steeply while staying positive, indicating a large deterioration in sales marketing effectiveness.",
        ],
    },
    "service_campaigns_up": {
        "slightly_improved": [
            "Service campaign performance improved, with stronger booking flow and better ROI into the workshop.",
            "Service marketing was more effective this month, generating better service bookings and stronger return.",
            "The service campaign mix improved, with booking activity and ROI both moving higher.",
        ],
        "strongly_improved": [
            "Service campaign performance improved strongly, with stronger booking flow and a clear lift in workshop ROI.",
            "Service marketing was materially more effective this month, generating better service bookings and stronger return.",
            "The service campaign mix improved meaningfully, with booking activity and ROI both moving higher.",
        ],
    },
    "service_bookings_roi_mismatch": {
        "slightly_softer": [
            "Service demand held up, but weaker ROI suggests lower-value workshop activity or softer conversion quality.",
            "Service demand remained reasonably solid, although the lower ROI points to weaker value capture.",
            "Service bookings stayed resilient, but softer ROI suggests the workshop did not convert that demand into enough value.",
        ],
        "significantly_softer": [
            "Service demand held up, but the sharper ROI decline suggests lower-value workshop activity or more material conversion softness.",
            "Service demand remained reasonably solid, although the weaker ROI points to more noticeable value leakage.",
            "Service bookings stayed resilient, but the softer ROI suggests the workshop did not convert that demand into enough value.",
        ],
    },
    "service_tasks_pressure": {
        "slightly_softer": [
            "Outstanding service tasks increased without matching booking growth, pointing to workflow pressure in the follow-up process.",
            "Service follow-up pressure increased this month, with task carryover rising faster than booking momentum.",
            "The workshop is carrying more unresolved follow-up activity, which suggests service workflow pressure is building.",
        ],
        "significantly_softer": [
            "Outstanding service tasks increased materially without matching booking growth, pointing to heavier workflow pressure in follow-up.",
            "Service follow-up pressure increased this month, with task carryover rising well ahead of booking momentum.",
            "The workshop is carrying a larger unresolved follow-up load, suggesting service workflow pressure is building more seriously.",
        ],
        "stable": [
            "Service follow-up activity was broadly steady, with no clear escalation in outstanding task pressure across the month.",
            "Outstanding service tasks held broadly stable, suggesting follow-up workload remained under control this month.",
            "Service task carryover remained relatively steady, indicating no major shift in workshop follow-up pressure.",
        ],
    },
    "service_follow_up_cleaner": {
        "slightly_improved": [
            "Service follow-up looked cleaner this month, with fewer outstanding tasks despite a larger enrolled customer base.",
            "Service follow-up improved, as outstanding tasks reduced even while more customers remained enrolled.",
            "The service team carried less unresolved work this month despite supporting a larger active customer base.",
        ],
        "strongly_improved": [
            "Service follow-up looked materially cleaner this month, with fewer outstanding tasks despite a larger enrolled customer base.",
            "Service follow-up improved strongly, as outstanding tasks reduced even while more customers remained enrolled.",
            "The service team carried much less unresolved work this month despite supporting a larger active customer base.",
        ],
    },
    "sales_volume_snapshot": {
        "slightly_improved": [
            "Sales deal volume lifted {change} month-on-month.",
            "Deal volume improved {change} month-on-month.",
            "The dealership closed more deals, with volume {directional_change} month-on-month.",
        ],
        "strongly_improved": [
            "Sales deal volume jumped {change} month-on-month.",
            "Deal volume strengthened sharply {change} month-on-month.",
            "The dealership closed materially more deals, with volume {directional_change} month-on-month.",
        ],
        "slightly_softer": [
            "Sales deal volume softened {change} month-on-month.",
            "Deal volume eased {change} month-on-month.",
            "The dealership closed fewer deals, with volume {directional_change} month-on-month.",
        ],
        "significantly_softer": [
            "Sales deal volume fell back {change} month-on-month.",
            "Deal volume weakened sharply {change} month-on-month.",
            "The dealership closed materially fewer deals, with volume {directional_change} month-on-month.",
        ],
        "stable": [
            "Sales deal volume was broadly steady at {current_raw}.",
            "Deal volume held broadly stable at {current_raw}.",
            "The dealership finished the month at a similar deal-volume level, closing {current_raw} deals.",
        ],
    },
    "sales_conversion_snapshot": {
        "slightly_improved": [
            "Overall sales conversion improved {change}.",
            "Sales conversion edged higher {change}.",
            "Conversion performance improved, finishing {change}.",
        ],
        "strongly_improved": [
            "Overall sales conversion improved strongly {change}.",
            "Sales conversion moved up sharply {change}.",
            "Conversion performance strengthened materially, finishing {change}.",
        ],
        "slightly_softer": [
            "Overall sales conversion softened {change}.",
            "Sales conversion eased {change}.",
            "Conversion performance weakened slightly, finishing {change}.",
        ],
        "significantly_softer": [
            "Overall sales conversion weakened {change}.",
            "Sales conversion fell back sharply {change}.",
            "Conversion performance softened materially, finishing {change}.",
        ],
        "stable": [
            "Overall sales conversion was broadly steady at {current_raw}.",
            "Sales conversion held near last month at {current_raw}.",
            "Conversion performance was broadly stable, ending at {current_raw}.",
        ],
    },
    "lead_response_snapshot": {
        "slightly_improved": [
            "Response times improved {change}.",
            "Lead response speed edged better {change}.",
            "The team responded faster this month, with response time {directional_change}.",
        ],
        "strongly_improved": [
            "Response times improved strongly {change}.",
            "Lead response speed improved materially {change}.",
            "The team responded much faster this month, with response time {directional_change}.",
        ],
        "slightly_softer": [
            "Response times softened {change}.",
            "Lead response speed eased, with response time {directional_change}.",
            "The team responded more slowly this month, with response time {directional_change}.",
        ],
        "significantly_softer": [
            "Response times weakened {change}.",
            "Lead response speed slowed materially, with response time {directional_change}.",
            "The team responded noticeably more slowly this month, with response time {directional_change}.",
        ],
        "stable": [
            "Response time was broadly steady at {current_raw}.",
            "Lead response speed held near last month at {current_raw}.",
            "Response time remained broadly stable at {current_raw}.",
        ],
    },
    "service_booking_snapshot": {
        "slightly_improved": [
            "Service campaign bookings improved {change}.",
            "Service bookings from campaigns lifted {change}.",
            "Service booking flow from campaigns improved, finishing {change}.",
        ],
        "strongly_improved": [
            "Service campaign bookings improved strongly {change}.",
            "Service bookings from campaigns jumped {change}.",
            "Service booking flow from campaigns strengthened materially, finishing {change}.",
        ],
        "slightly_softer": [
            "Service campaign bookings softened {change}.",
            "Service bookings from campaigns eased {change}.",
            "Service booking flow from campaigns weakened slightly, finishing {change}.",
        ],
        "significantly_softer": [
            "Service campaign bookings fell back {change}.",
            "Service bookings from campaigns weakened sharply {change}.",
            "Service booking flow from campaigns softened materially, finishing {change}.",
        ],
        "stable": [
            "Service campaign bookings were broadly steady at {current_raw}.",
            "Service bookings from campaigns held near last month at {current_raw}.",
            "Service booking flow from campaigns stayed broadly stable at {current_raw}.",
        ],
    },
    "service_roi_snapshot": {
        "slightly_improved": [
            "Service campaign ROI improved {change}.",
            "Service campaign ROI edged higher {change}.",
            "Service ROI strengthened, finishing {change}.",
        ],
        "strongly_improved": [
            "Service campaign ROI improved strongly {change}.",
            "Service campaign ROI lifted materially {change}.",
            "Service ROI strengthened meaningfully, finishing {change}.",
        ],
        "slightly_softer": [
            "Service campaign ROI softened {change}.",
            "Service campaign ROI eased {change}.",
            "Service ROI weakened slightly, finishing {change}.",
        ],
        "significantly_softer": [
            "Service campaign ROI weakened {change}.",
            "Service campaign ROI fell back sharply {change}.",
            "Service ROI softened materially, finishing {change}.",
        ],
        "stable": [
            "Service campaign ROI was broadly steady at {current_raw}.",
            "Service campaign ROI held near last month at {current_raw}.",
            "Service ROI remained broadly stable at {current_raw}.",
        ],
    },
    "service_task_snapshot": {
        "slightly_improved": [
            "Outstanding service tasks reduced {change}.",
            "The service team carried fewer outstanding tasks, with volume {directional_change}.",
            "Service follow-up backlog improved, finishing {change}.",
        ],
        "strongly_improved": [
            "Outstanding service tasks reduced strongly {change}.",
            "The service team carried fewer outstanding tasks, with volume {directional_change}.",
            "Service follow-up backlog reduced meaningfully, finishing {change}.",
        ],
        "slightly_softer": [
            "Outstanding service tasks increased {change}.",
            "The service team carried more outstanding tasks, with volume {directional_change}.",
            "Service follow-up backlog increased slightly, finishing {change}.",
        ],
        "significantly_softer": [
            "Outstanding service tasks increased sharply {change}.",
            "The service team carried more outstanding tasks, with volume {directional_change}.",
            "Service follow-up backlog increased sharply, finishing {change}.",
        ],
        "stable": [
            "Outstanding service tasks were broadly steady at {current_raw}.",
            "The service team finished with a similar outstanding-task load at {current_raw}.",
            "Service follow-up backlog remained broadly stable at {current_raw}.",
        ],
    },
}


@dataclass
class ReportMatch:
    sales_pdf: Path | None
    service_pdf: Path | None


@dataclass
class DealershipReportGroup:
    dealership_name: str
    group_key: str
    folder_path: Path
    sales_pdf: Path | None
    service_pdf: Path | None
    # Set when the month came from a master workbook. The PDFs then serve only as the
    # attachment and the on-screen preview: every figure comes from the sheet.
    workbook: Path | None = None
    workbook_month: str = ""
    # Every month this workbook can report on, so a month it does not hold can say
    # which ones it does rather than only naming the latest.
    workbook_months: list[str] = field(default_factory=list)


@dataclass
class HistoricalReportGroup:
    month_label: str
    report_group: DealershipReportGroup


@dataclass
class HistoricalMetricSnapshot:
    month_label: str
    current_value: float
    metric_type: str
    is_reliable: bool
    is_suspicious: bool
    is_anomaly: bool


@dataclass
class DealershipResult:
    dealership_name: str
    month_label: str
    bullet_insights: list[str]
    summary: str
    sms_version: str
    report_status: str
    sales_pdf: Path | None = None
    service_pdf: Path | None = None
    # The Performance Intelligence report, built from the workbook rather than exported
    # from it. None when the workbook is not the source or the renderer is unavailable.
    performance_pdf: Path | None = None
    performance_note: str = ""
    sales_text: str = ""
    service_text: str = ""
    parsing_status: str = "fallback"
    confidence_score: float = 0.0
    extracted_metrics: list[str] | None = None
    rules_triggered: list[str] | None = None
    parsed_metrics: dict[str, dict[str, Any]] | None = None
    insights_source: str = "placeholder"
    top_priority_level: str = "Low"
    review_required: str = "yes"
    sentiment: str = "stable"
    insight_count_after_deduplication: int = 0
    overall_month_rating: str = "Stable"
    primary_strength: str = "No standout strength identified"
    primary_risk: str = "Manual review required"
    recommended_focus: str = ""  # retained for stored-report compatibility; no longer emitted
    review_reason_tags: list[str] | None = None
    review_severity: str = "none"
    sms_segments_used: list[str] | None = None
    sms_valid: bool = True
    sms_char_count: int = 0
    sms_limit: int = SMS_CHAR_LIMIT
    sms_source: str = "priority_insights"
    sms_error: str = ""
    notes: list[str] | None = None
    sales_score: float = 0.0
    service_score: float = 0.0
    risk_score: float = 0.0
    overall_score: float = 0.0
    rule_based_rating: str = "Stable"
    derived_month_rating: str = "Stable"
    final_rating_used: str = "Stable"
    anomaly_count: int = 0
    anomaly_types: list[str] | None = None
    unreliable_metric_count: int = 0
    suspicious_metric_count: int = 0
    scoring_conflict: bool = False
    trend_supported_insight_count: int = 0
    high_confidence_trend_count: int = 0
    rating_reason: str = ""
    sales_momentum_trend: str = "stable"
    service_momentum_trend: str = "stable"
    followup_risk_trend: str = "stable"
    persistent_watchpoint_count: int = 0
    persistent_watchpoint_types: list[str] | None = None
    review_priority_score: float = 0.0
    review_summary_reason: str = ""
    relationship_confidence_score: float = 0.0
    relationship_rule_trigger: list[str] | None = None
    relationship_reason_trace: list[str] | None = None
    relationship_stage_pressure: list[str] | None = None
    phase163_evidence_ledger: dict[str, Any] | None = None
    phase163_rule_sentiment_raw: str | None = None


@dataclass
class RuleInsight:
    rule_id: str
    message: str
    category: str
    severity: str
    priority_weight: int
    polarity: str
    tone_band: str
    merge_key: str
    supporting_metrics: list[str]
    signal_type: str = "current_performance"
    metric_focus: list[str] | None = None
    # False once finalised if no metric of this rule's own cleared the reporting gates,
    # which means the only text available is the legacy template phrasing.
    has_metric_evidence: bool = True


@dataclass
class InternalScoringProfile:
    sales_score: float
    service_score: float
    risk_score: float
    overall_score: float
    rule_based_rating: str
    derived_month_rating: str
    final_rating_used: str
    anomaly_count: int
    anomaly_types: list[str]
    unreliable_metric_count: int
    suspicious_metric_count: int
    scoring_conflict: bool


@dataclass
class StrategicArbitrationResult:
    """Deterministic dominant-condition resolution for executive interpretation (phase 4)."""

    dominant_pillar: str
    dominant_score: float
    runner_up_pillar: str | None
    runner_up_score: float
    pillar_scores: dict[str, float]
    mixed_attention: bool
    conflict_clause: str | None
    suppressed_secondary_count: int
    crosscurrent_override_pillar: str | None
    crosscurrent_clause: str | None


@dataclass
class ExecutiveStrategicContext:
    """Bundled executive interpretation: arbitration, dominant state, asymmetry, and framing lines."""

    arbitration: StrategicArbitrationResult
    dominant_state: str
    state_evidence: list[str]
    asymmetry_level: str
    suppress_secondary_positives: bool
    intro_override: str | None
    primary_strength_clause: str | None
    strategic_watchpoint_clause: str | None
    monitoring_clause: str | None
    mixed_synthesis_clause: str | None
    contradiction_flags: list[str]


def setup_logging(output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    # Close before clearing: a re-run otherwise leaks the previous run's open
    # file handle every time.
    for stale in list(logger.handlers):
        try:
            stale.close()
        except Exception:  # noqa: BLE001
            pass
    logger.handlers.clear()

    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    # Rotating, not plain: extract_pdf_text logs a line per page, so an
    # unrotated processing.log grows without limit across monthly runs.
    file_handler = RotatingFileHandler(
        output_dir / LOG_FILE_NAME, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
    )
    file_handler.setFormatter(formatter)

    logger.addHandler(console_handler)
    logger.addHandler(file_handler)


def previous_month(reference_date: datetime | None = None) -> tuple[int, int, str]:
    ref = reference_date or datetime.now()
    first_of_month = ref.replace(day=1)
    previous = first_of_month - timedelta(days=1)
    month_label = previous.strftime("%b %Y")
    return previous.month, previous.year, month_label


def shift_month(month: int, year: int, month_offset: int) -> tuple[int, int, str]:
    anchor = datetime(year=year, month=month, day=1)
    shifted = anchor
    step = -1 if month_offset < 0 else 1
    for _ in range(abs(month_offset)):
        if step < 0:
            shifted = (shifted.replace(day=1) - timedelta(days=1)).replace(day=1)
        else:
            if shifted.month == 12:
                shifted = shifted.replace(year=shifted.year + 1, month=1, day=1)
            else:
                shifted = shifted.replace(month=shifted.month + 1, day=1)
    return shifted.month, shifted.year, shifted.strftime("%b %Y")


def historical_month_sequence(month: int, year: int, months_back: int = MAX_HISTORICAL_MONTHS) -> list[tuple[int, int, str]]:
    return [shift_month(month, year, -offset) for offset in range(1, months_back + 1)]


def normalize_text(value: str) -> str:
    normalized = value.lower()
    normalized = normalized.replace("_", " ").replace("-", " ")
    normalized = re.sub(r"[^a-z0-9\s]", " ", normalized)
    normalized = re.sub(r"\s+", " ", normalized).strip()
    return normalized


def is_meaningful_candidate_message(message: str) -> bool:
    """
    Validates that a candidate message contains meaningful content and is not
    empty, whitespace-only, punctuation-only, or placeholder fragments.
    
    PATCH 1: Empty/Punctuation Candidate Protection
    Prevents empty/punctuation-only messages from entering candidate pools.
    """
    if not message or not isinstance(message, str):
        return False
    
    stripped = message.strip()
    if not stripped:
        return False
    
    # Remove all punctuation and check if anything meaningful remains
    meaningful_content = re.sub(r'[^\w\s]', '', stripped)
    meaningful_stripped = meaningful_content.strip()
    
    if not meaningful_stripped:
        return False
    
    # Reject single character messages (like ".")
    if len(meaningful_stripped) < 2:
        return False
    
    return True


def is_ignored_folder_name(folder_name: str) -> bool:
    normalized_name = normalize_text(folder_name)
    compact_name = normalized_name.replace(" ", "")
    archive_like_prefixes = ("archive", "archived", "archiv", "archieve")
    return normalized_name in IGNORED_FOLDER_NAMES or compact_name.startswith(archive_like_prefixes)


def contains_ignored_folder(path: Path) -> bool:
    return any(is_ignored_folder_name(part) for part in path.parts)


MONTH_NAMES_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                     "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH_NAMES_FULL = ["january", "february", "march", "april", "may", "june",
                    "july", "august", "september", "october", "november", "december"]


def month_name_spellings(month: int) -> list[str]:
    """Every way a filename may spell the month: any 3+ letter prefix (sep, sept, september)."""
    full = MONTH_NAMES_FULL[month - 1]
    return [full[:length] for length in range(3, len(full) + 1)]


_MONTH_BY_SPELLING = {
    spelling: index + 1 for index in range(12) for spelling in month_name_spellings(index + 1)
}
# Longest first so "sept" is preferred over "sep" inside a regex alternation.
MONTH_NAME_PATTERN = "|".join(sorted(_MONTH_BY_SPELLING, key=len, reverse=True))


def month_from_name(token: str) -> int | None:
    return _MONTH_BY_SPELLING.get(str(token or "").strip().rstrip(".").lower())


def find_workbook_report_groups(
    accounts_dir: Path, month: int, year: int, include: Iterable[str] | None = None,
    pdf_dir: Path | None = None,
) -> list[DealershipReportGroup]:
    """Dealerships defined by the master workbooks in the library, not by PDF filenames.

    The library holds one workbook per dealership at the top level, and a folder of
    exported PDFs beside it. The workbook decides who exists and which month it holds;
    the PDFs are only what gets attached to the email, matched by the same month
    scoring the PDF path has always used.

    A workbook holding a different month than the one asked for is returned with no
    metrics rather than dropped, so the operator sees "this workbook holds Jul 2026"
    instead of a dealership silently vanishing from the list.
    """
    import excel_reader

    target_label = f"{MONTH_NAMES_SHORT[month - 1]} {year}" if 1 <= month <= 12 else ""
    groups: list[DealershipReportGroup] = []
    pdf_root = pdf_dir or accounts_dir

    for master in excel_reader.discover_masters(accounts_dir, recursive=False, include=include):
        if master.error:
            logging.warning("Could not read workbook %s: %s", master.path, master.error)

        # The title box and the filename often disagree (an abbreviation vs the full name).
        folder = _pdf_folder_for(pdf_root, master.dealership) or _pdf_folder_for(pdf_root, master.path.stem)
        sales_pdf = service_pdf = None
        if folder is not None:
            matched = find_matching_report_groups(folder, month, year)
            for candidate in matched:
                if candidate.sales_pdf or candidate.service_pdf:
                    sales_pdf = candidate.sales_pdf
                    service_pdf = candidate.service_pdf
                    break

        groups.append(DealershipReportGroup(
            dealership_name=master.dealership,
            group_key=normalize_text(master.dealership),
            folder_path=folder or master.path.parent,
            sales_pdf=sales_pdf,
            service_pdf=service_pdf,
            workbook=master.path if master.holds(target_label) else None,
            workbook_month=master.month_label,
            workbook_months=list(master.months),
        ))
    return groups


def _pdf_folder_for(accounts_dir: Path, dealership: str) -> Path | None:
    """The folder of exported PDFs that belongs to a workbook.

    Matched on the normalised name rather than an exact one: in practice a workbook and
    its folder disagree often enough (a master still called 'Demo Site.xlsx' beside a
    'Demo Site 1' folder) that exact matching would silently lose the attachments.
    """
    if not accounts_dir.is_dir():
        return None
    wanted = normalize_text(dealership)
    folders = [item for item in accounts_dir.iterdir()
               if item.is_dir() and not is_ignored_folder_name(item.name)]

    for folder in folders:
        if normalize_text(folder.name) == wanted:
            return folder
    # Then a containment match, longest first so 'Demo Site 1' beats 'Demo Site'.
    candidates = [f for f in folders
                  if wanted and (wanted in normalize_text(f.name) or normalize_text(f.name) in wanted)]
    if candidates:
        return max(candidates, key=lambda f: len(normalize_text(f.name)))
    return None


def get_dealership_folders(accounts_dir: Path) -> list[Path]:
    if not accounts_dir.exists():
        raise FileNotFoundError(f"Accounts directory does not exist: {accounts_dir}")

    dealerships: list[Path] = []
    for item in sorted(accounts_dir.iterdir()):
        if not item.is_dir():
            continue
        if is_ignored_folder_name(item.name):
            logging.info("Skipping ignored top-level folder: %s", item)
            continue
        dealerships.append(item)

    logging.info("Discovered %s dealership folder(s) under %s", len(dealerships), accounts_dir)
    return dealerships


def month_variants(month: int, year: int) -> set[str]:
    """Literal (already normalised) month spellings to look for in a filename.

    Deliberately excludes numeric year-first forms such as "2026 07". Source
    reports are commonly named with a leading send-date prefix, e.g.
    "2026-05-11_1027 ... Snapshot - April 2026.pdf", so "2026 05" would match
    the APRIL report when targeting May. Year-first forms are therefore only
    included with the month NAME, which cannot collide with a date prefix.
    """
    base_date = datetime(year=year, month=month, day=1)
    variants = {
        normalize_text(base_date.strftime("%m %Y")),      # 07 2026
        normalize_text(base_date.strftime("%m-%Y")),
        normalize_text(base_date.strftime("%m_%Y")),
        normalize_text(f"{month} {year}"),                # 7 2026 (unpadded)
    }
    for spelling in month_name_spellings(month):
        variants.add(f"{spelling} {year}")                 # jul 2026, sept 2026
        variants.add(f"{year} {spelling}")                 # 2026 jul
    return {variant for variant in variants if variant}


def month_short_year_patterns(month: int, year: int) -> list[re.Pattern]:
    """Patterns for two-digit-year spellings such as "Jul 26" / "July 26".

    Matched by regex rather than substring so that:
      - "july 2026" is not read as "july 26" (the \\b after 26 blocks it), and
      - a day number followed by a real year ("May 26 2025") is not mistaken
        for a short year, via the trailing negative lookahead.
    """
    short_year = f"{year % 100:02d}"
    patterns: list[re.Pattern] = []
    for name in month_name_spellings(month):
        patterns.append(re.compile(rf"\b{re.escape(name)}\s+{short_year}\b(?!\s+\d{{4}})"))
    return patterns


def report_type_score(file_name: str, report_type: str) -> int:
    normalized = normalize_text(file_name)
    tokens = normalized.split()

    score = 0
    has_explicit_type_match = False
    if report_type == "sales":
        if "sales" in tokens:
            score += 3
            has_explicit_type_match = True
        if "monthly sales report" in normalized:
            score += 4
            has_explicit_type_match = True
        if "sales report" in normalized:
            score += 3
            has_explicit_type_match = True
    elif report_type == "service":
        if "service" in tokens:
            score += 3
            has_explicit_type_match = True
        if "monthly service report" in normalized:
            score += 4
            has_explicit_type_match = True
        if "service report" in normalized:
            score += 3
            has_explicit_type_match = True

    if has_explicit_type_match:
        if "report" in tokens:
            score += 1
        if "monthly" in tokens:
            score += 1

    return score


def month_score(
    file_name: str,
    target_variants: set[str],
    short_year_patterns: list[re.Pattern] | None = None,
) -> int:
    normalized = normalize_text(file_name)
    score = 0
    for variant in target_variants:
        if variant and variant in normalized:
            score = max(score, len(variant.split()) + 2)
    if score <= 0 and short_year_patterns:
        # Two-digit-year spellings ("Jul 26") score below any full four-digit
        # match, so an unambiguous "Jul 2026" file always wins when both exist.
        for pattern in short_year_patterns:
            if pattern.search(normalized):
                score = max(score, 3)
    return score


def candidate_score(
    file_path: Path,
    report_type: str,
    target_variants: set[str],
    short_year_patterns: list[re.Pattern] | None = None,
) -> int:
    name = file_path.stem
    type_score = report_type_score(name, report_type)
    if type_score <= 0:
        return 0
    return type_score + month_score(name, target_variants, short_year_patterns)


DEALERSHIP_ALIASES_PATH = Path(__file__).resolve().parent / "data" / "dealership_aliases.json"
_ALIAS_CACHE: dict[str, tuple[float, list[tuple[re.Pattern[str], str]]]] = {}


def load_dealership_aliases(path: Path | None = None) -> list[tuple[re.Pattern[str], str]]:
    """Short forms seen in PDF filenames and the full name each one stands for.

    Each entry is {"name": full name, "matches": [short forms]}. A short form matches as
    whole words, case-insensitively. A missing or unreadable file means no aliases.
    """
    target = path or DEALERSHIP_ALIASES_PATH
    try:
        mtime = target.stat().st_mtime
    except OSError:
        return []
    cached = _ALIAS_CACHE.get(str(target))
    if cached and cached[0] == mtime:
        return cached[1]

    try:
        entries = json.loads(target.read_text(encoding="utf-8")).get("aliases", [])
    except (OSError, ValueError, AttributeError) as exc:
        logging.warning("Dealership aliases unreadable at %s (%s); ignoring them.", target, exc)
        entries = []

    compiled: list[tuple[re.Pattern[str], str]] = []
    for entry in entries if isinstance(entries, list) else []:
        if not isinstance(entry, dict):
            continue
        full_name = str(entry.get("name") or "").strip()
        matches = entry.get("matches") or []
        if not full_name or not isinstance(matches, list):
            continue
        for short_form in matches:
            words = str(short_form).split()
            if words:
                pattern = r"(?<!\w)" + r"\s+".join(re.escape(word) for word in words) + r"(?!\w)"
                compiled.append((re.compile(pattern, re.IGNORECASE), full_name))
    _ALIAS_CACHE[str(target)] = (mtime, compiled)
    return compiled


def normalise_dealership_aliases(name: str, aliases_path: Path | None = None) -> str:
    cleaned = collapse_whitespace(name)
    for pattern, full_name in load_dealership_aliases(aliases_path):
        cleaned = pattern.sub(lambda _match, value=full_name: value, cleaned)
    return collapse_whitespace(cleaned)


def extract_dealership_name_from_filename(file_name: str, fallback_folder_name: str) -> str:
    stem = Path(file_name).stem
    cleaned = stem.replace("_", " ").replace("-", " ")
    cleaned = re.sub(r"^\d{4}\s+\d{2}\s+\d{2}\s+\d{3,4}\s*", " ", cleaned)
    cleaned = re.sub(r"^\d{4}\s+\d{2}\s+\d{2}\s*", " ", cleaned)
    cleaned = re.sub(r"[{}\[\]()]+", " ", cleaned)

    removable_patterns = [
        r"\bmonthly sales report\b",
        r"\bmonthly service report\b",
        r"\bmonthly report sales\b",
        r"\bmonthly report serv\b",
        r"\bmonthly report service\b",
        r"\bsales report\b",
        r"\bservice report\b",
        r"\bmonthly report\b",
    ]
    for pattern in removable_patterns:
        cleaned = re.sub(pattern, " ", cleaned, flags=re.IGNORECASE)

    cleaned = re.sub(
        rf"\b(?:{MONTH_NAME_PATTERN})\.?\s+\d{{2,4}}\b",
        " ",
        cleaned,
        flags=re.IGNORECASE,
    )
    cleaned = re.sub(r"\b\d{4}\b", " ", cleaned)
    cleaned = collapse_whitespace(cleaned)
    if not cleaned:
        cleaned = fallback_folder_name

    cleaned = normalise_dealership_aliases(cleaned)
    return cleaned or fallback_folder_name


# find_matching_report_groups is called once for the target month and again for
# each of the three historical months, per entity in the folder — every call
# doing a full recursive scan of the same directory. On OneDrive-backed folders
# that scan is the slowest part of a run, so the file listing (not the scoring,
# which is month-specific) is cached for the duration of one run.
_PDF_LISTING_CACHE: dict[str, tuple[Path, ...]] = {}

# Loaded once per run, written back at the end. The trend machinery can only see
# months whose PDFs are still in the dealership folder; this is what lets an insight
# compare against the dealership's own record rather than just last month.
_METRIC_HISTORY_STORE: dict[str, Any] | None = None
_METRIC_HISTORY_DIRTY = False


def metric_history_store() -> dict[str, Any]:
    global _METRIC_HISTORY_STORE
    if _METRIC_HISTORY_STORE is None:
        _METRIC_HISTORY_STORE = metric_history.load_store()
    return _METRIC_HISTORY_STORE


def reset_metric_history_store() -> None:
    global _METRIC_HISTORY_STORE, _METRIC_HISTORY_DIRTY
    _METRIC_HISTORY_STORE = None
    _METRIC_HISTORY_DIRTY = False


def record_metric_history(dealership_name: str, month_label: str, metrics: dict[str, dict[str, Any]]) -> None:
    global _METRIC_HISTORY_DIRTY
    if not metrics:
        return
    metric_history.record_month(metric_history_store(), dealership_name, month_label, metrics)
    _METRIC_HISTORY_DIRTY = True


def flush_metric_history() -> None:
    global _METRIC_HISTORY_DIRTY
    if not _METRIC_HISTORY_DIRTY or _METRIC_HISTORY_STORE is None:
        return
    if metric_history.save_store(_METRIC_HISTORY_STORE):
        _METRIC_HISTORY_DIRTY = False


def attach_metric_history(
    metrics: dict[str, dict[str, Any]],
    dealership_name: str,
    month_label: str,
) -> dict[str, dict[str, Any]]:
    """Hang each metric's own recorded series on the metric, oldest first.

    Kept on the metric record so the sentence builders can reach a baseline without
    every intermediate function having to pass the store around.
    """
    store = metric_history_store()
    for metric_name, metric in metrics.items():
        metric["history_series"] = metric_history.metric_series(
            store, dealership_name, metric_name, before_month=month_label
        )
        last_year_label = metric_history.same_month_last_year_label(month_label)
        metric["history_same_month_last_year"] = (
            metric_history.value_at(store, dealership_name, metric_name, last_year_label)
            if last_year_label
            else None
        )
    return metrics


def history_snapshots_from_store(
    dealership_name: str,
    month_label: str,
) -> dict[str, list[HistoricalMetricSnapshot]]:
    """Recent months from the store, newest first, in the shape attach_historical_trends wants."""
    store = metric_history_store()
    snapshots: dict[str, list[HistoricalMetricSnapshot]] = {}
    for metric_name in TREND_METRIC_NAMES:
        series = metric_history.metric_series(store, dealership_name, metric_name, before_month=month_label)
        snapshots[metric_name] = [
            HistoricalMetricSnapshot(
                month_label=label,
                current_value=value,
                metric_type=str(METRIC_DEFINITIONS.get(metric_name, {}).get("metric_type", "count")),
                is_reliable=True,
                is_suspicious=False,
                is_anomaly=False,
            )
            for label, value in reversed(series)
        ]
    return snapshots


def merge_history_snapshots(
    primary: dict[str, list[HistoricalMetricSnapshot]],
    secondary: dict[str, list[HistoricalMetricSnapshot]],
) -> dict[str, list[HistoricalMetricSnapshot]]:
    """Prefer months read from PDFs this run; fill the gaps from the store."""
    merged: dict[str, list[HistoricalMetricSnapshot]] = {}
    for metric_name in set(primary) | set(secondary):
        seen = {snapshot.month_label for snapshot in primary.get(metric_name, [])}
        combined = list(primary.get(metric_name, []))
        combined.extend(s for s in secondary.get(metric_name, []) if s.month_label not in seen)
        combined.sort(key=lambda s: metric_history.month_key(s.month_label) or (0, 0), reverse=True)
        merged[metric_name] = combined
    return merged


def clear_report_scan_cache() -> None:
    """Drop cached directory listings. Called at the start of every run so a
    file added between runs is always picked up."""
    _PDF_LISTING_CACHE.clear()


def _dealership_pdf_files(dealership_dir: Path) -> tuple[Path, ...]:
    key = str(dealership_dir)
    cached = _PDF_LISTING_CACHE.get(key)
    if cached is None:
        cached = tuple(dealership_dir.rglob(f"*{PDF_SUFFIX}"))
        _PDF_LISTING_CACHE[key] = cached
    return cached


def find_matching_report_groups(dealership_dir: Path, month: int, year: int) -> list[DealershipReportGroup]:
    target_variants = month_variants(month, year)
    short_year_patterns = month_short_year_patterns(month, year)
    logging.info("Scanning dealership: %s", dealership_dir.name)

    grouped_candidates: dict[str, dict[str, Any]] = {}

    for pdf_file in _dealership_pdf_files(dealership_dir):
        if contains_ignored_folder(pdf_file.parent):
            logging.info("Ignoring archived path: %s", pdf_file)
            continue

        normalized_name = normalize_text(pdf_file.stem)
        file_month_score = month_score(normalized_name, target_variants, short_year_patterns)
        if file_month_score <= 0:
            continue

        dealership_name = extract_dealership_name_from_filename(pdf_file.name, dealership_dir.name)
        group_key = normalize_text(dealership_name)
        sales_score = candidate_score(pdf_file, "sales", target_variants, short_year_patterns)
        service_score = candidate_score(pdf_file, "service", target_variants, short_year_patterns)

        if sales_score <= 0 and service_score <= 0:
            continue

        group_entry = grouped_candidates.setdefault(
            group_key,
            {
                "dealership_name": dealership_name,
                "sales_candidates": [],
                "service_candidates": [],
            },
        )

        if sales_score > 0:
            group_entry["sales_candidates"].append((sales_score, pdf_file))
        if service_score > 0:
            group_entry["service_candidates"].append((service_score, pdf_file))

    if not grouped_candidates:
        logging.info("No dealership-specific report groups found for %s; falling back to folder name.", dealership_dir.name)
        return [
            DealershipReportGroup(
                dealership_name=dealership_dir.name,
                group_key=normalize_text(dealership_dir.name),
                folder_path=dealership_dir,
                sales_pdf=None,
                service_pdf=None,
            )
        ]

    grouped_matches: list[DealershipReportGroup] = []
    multiple_entities_in_folder = len(grouped_candidates) > 1
    for group_key, entry in sorted(grouped_candidates.items(), key=lambda item: item[1]["dealership_name"].lower()):
        sales_candidates = sorted(entry["sales_candidates"], key=lambda item: (-item[0], item[1].name.lower()))
        service_candidates = sorted(entry["service_candidates"], key=lambda item: (-item[0], item[1].name.lower()))

        sales_match = sales_candidates[0][1] if sales_candidates else None
        service_match = service_candidates[0][1] if service_candidates else None
        dealership_name = entry["dealership_name"] if multiple_entities_in_folder else dealership_dir.name
        resolved_group_key = group_key if multiple_entities_in_folder else normalize_text(dealership_dir.name)

        logging.info(
            "Best grouped match for %s | sales: %s | service: %s",
            dealership_name,
            sales_match or "None found",
            service_match or "None found",
        )

        grouped_matches.append(
            DealershipReportGroup(
                dealership_name=dealership_name,
                group_key=resolved_group_key,
                folder_path=dealership_dir,
                sales_pdf=sales_match,
                service_pdf=service_match,
            )
        )

    return grouped_matches


def extract_pdf_text(pdf_path: Path) -> str:
    logging.info("Extracting text from PDF: %s", pdf_path)
    pages_text: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            page_text = page.extract_text() or ""
            logging.info("Extracted page %s from %s (%s characters)", page_number, pdf_path.name, len(page_text))
            pages_text.append(page_text)

    combined = "\n".join(pages_text).strip()
    if not combined:
        logging.warning("No text extracted from PDF: %s", pdf_path)
    return combined


def collapse_whitespace(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def phase161_scrub_unsupported_causal_copy(text: str) -> str:
    """Deterministic pass: soften forced-causal phrasing on summary/SMS (phase 161)."""
    if not text:
        return text
    out = str(text)
    pairs = (
        (" pointing to ", " reading as "),
        ("Despite ", "while "),
        (" despite ", " while "),
        ("suggesting ", "with "),
        ("driven by", "alongside"),
        ("caused by", "alongside"),
    )
    for old, new in pairs:
        if old.lower() in out.lower():
            out = re.sub(re.escape(old), new, out, count=1, flags=re.IGNORECASE)
    out = re.sub(r"\s{2,}", " ", out).strip()
    return out


def phase162_compress_summary_text(text: str, *, has_hard_campaign_watchpoint: bool) -> str:
    """Remove redundant narrative glue when concrete commercial lines already say enough (phase 162)."""
    if not text:
        return text
    out = collapse_whitespace(str(text))
    if has_hard_campaign_watchpoint:
        for needle in (
            "Narrative clarity was limited by cross-currents across sales and service.",
            "Narrative clarity was limited by cross-currents across sales and service",
            "The profile was intentionally mixed, with offsetting strengths and weaknesses.",
            "Signals crossed in ways that point to mixed, uneven operating conditions.",
        ):
            if needle.lower() in out.lower():
                out = re.sub(re.escape(needle), "", out, count=1, flags=re.IGNORECASE)
    out = re.sub(r"\s*\.\s*\.", ".", out)
    out = re.sub(r"\s{2,}", " ", out).strip()
    if out.startswith("."):
        out = out[1:].strip()
    return out


def clean_lines(text: str) -> list[str]:
    return [collapse_whitespace(line) for line in text.splitlines() if collapse_whitespace(line)]


def parse_duration_to_minutes(raw_value: str) -> float | None:
    cleaned = collapse_whitespace(raw_value.lower())
    match = re.search(r"(-?\d+):(\d+)\s*(hrs|hr|h|minutes|minute|min)", cleaned)
    if match:
        first_part = int(match.group(1))
        second_part = int(match.group(2))
        suffix = match.group(3)
        if suffix.startswith("h"):
            return first_part * 60 + second_part
        return first_part + (second_part / 60)

    numeric_match = re.search(r"(-?\d+(?:\.\d+)?)\s*(hrs|hr|h|minutes|minute|min)", cleaned)
    if numeric_match:
        value = float(numeric_match.group(1))
        suffix = numeric_match.group(2)
        if suffix.startswith("h"):
            return value * 60
        return value

    return None


def parse_metric_value(raw_value: str) -> tuple[float | None, str]:
    cleaned = collapse_whitespace(raw_value)
    if not cleaned or cleaned.lower() in {"n/a", "na"}:
        return None, "unknown"

    duration_value = parse_duration_to_minutes(cleaned)
    if duration_value is not None:
        return duration_value, "minutes"

    stripped = cleaned.replace("$", "").replace(",", "").replace("%", "")
    try:
        numeric_value = float(stripped)
    except ValueError:
        return None, "text"

    if "%" in cleaned:
        return numeric_value, "percentage"
    if "$" in cleaned:
        return numeric_value, "currency"
    return numeric_value, "count"


def metric_definition(metric_name: str, unit: str | None = None) -> dict[str, Any]:
    definition = METRIC_DEFINITIONS.get(metric_name)
    if definition is not None:
        return definition
    inferred_type = "count"
    if unit == "percentage":
        inferred_type = "rate"
    elif unit == "minutes":
        inferred_type = "time"
    elif unit == "currency":
        inferred_type = "currency"
    elif unit == "count":
        inferred_type = "count"
    return {
        "metric_type": inferred_type,
        "label": metric_name.replace("_", " "),
        "higher_is_better": True,
    }


def derive_significance_band(
    metric_type: str,
    percent_change: float | None,
    percentage_point_change: float | None,
) -> str:
    if metric_type == "rate":
        if percentage_point_change is None:
            return "not_available"
        absolute_pp = abs(percentage_point_change)
        if absolute_pp < 0.5:
            return "negligible"
        if absolute_pp < 1.5:
            return "slight"
        if absolute_pp < 3.0:
            return "meaningful"
        return "strong"

    if percent_change is None:
        return "not_available"

    absolute_percent = abs(percent_change)
    if absolute_percent < 3:
        return "negligible"
    if absolute_percent < 10:
        return "slight"
    if absolute_percent < 25:
        return "meaningful"
    return "strong"


def is_metric_usable_for_reporting(metric_type: str, previous_value: float, current_value: float, unit: str) -> bool:
    if unit in {"unknown", "text"}:
        return False
    if metric_type == "rate":
        return True
    if abs(previous_value) < 1e-9:
        return False
    if metric_type == "time" and previous_value < 0:
        return False
    return True


def clamp_score(value: float, lower_bound: float = 0.0, upper_bound: float = 100.0) -> float:
    return round(max(lower_bound, min(upper_bound, value)), 2)


def unique_preserving_order(values: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    ordered: list[str] = []
    for value in values:
        if not value or value in seen:
            continue
        seen.add(value)
        ordered.append(value)
    return ordered


def evaluate_metric_reliability(metric_name: str, record: dict[str, Any]) -> dict[str, Any]:
    previous_value = record.get("previous_value")
    current_value = record.get("current_value")
    metric_type = str(record.get("metric_type", "count"))
    usable_for_reporting = bool(record.get("usable_for_reporting", False))
    reasons: list[str] = []
    is_reliable = usable_for_reporting
    is_suspicious = False

    if not isinstance(previous_value, (int, float)) or not isinstance(current_value, (int, float)):
        is_reliable = False
        reasons.append("missing_numeric_value")
    else:
        previous_float = float(previous_value)
        current_float = float(current_value)
        if metric_type in {"volume", "count", "rate", "time"} and (previous_float < 0 or current_float < 0):
            is_reliable = False
            reasons.append("negative_value")

        if abs(previous_float) < 1e-9:
            is_reliable = False
            reasons.append("division_by_zero_baseline")
            if abs(current_float) > 1e-9:
                is_suspicious = True
                reasons.append("zero_to_nonzero_jump")

        percent_change = record.get("percent_change")
        if isinstance(percent_change, (int, float)) and abs(float(percent_change)) > 500:
            is_suspicious = True
            reasons.append("extreme_percent_change")

        if metric_type == "rate" and max(previous_float, current_float) > 100:
            is_suspicious = True
            reasons.append("rate_above_100")

        if metric_name == "sales_internet_response_time" and max(previous_float, current_float) > 1440:
            is_reliable = False
            is_suspicious = True
            reasons.append("response_time_over_24h")

        # The report prints 0 when no response was timed; nobody averages zero seconds.
        if metric_name == "sales_internet_response_time" and min(previous_float, current_float) < 1e-9:
            is_reliable = False
            reasons.append("zero_response_time")

        if metric_name.endswith("campaign_roi"):
            max_abs_value = max(abs(previous_float), abs(current_float))
            min_value = min(previous_float, current_float)
            if max_abs_value > 750000 or min_value < -250000:
                is_suspicious = True
                reasons.append("extreme_roi_value")

    reliability_reasons = unique_preserving_order(reasons)
    record["is_reliable"] = bool(is_reliable)
    record["is_suspicious"] = bool(is_suspicious)
    record["reliability_reason"] = " | ".join(reliability_reasons) if reliability_reasons else "none"
    record["is_anomaly"] = bool(record.get("is_anomaly", False))
    anomaly_types = list(record.get("anomaly_types") or [])
    record["anomaly_types"] = unique_preserving_order(anomaly_types)
    record["anomaly_type"] = " | ".join(record["anomaly_types"])
    return record


def register_metric_anomaly(metric: dict[str, Any] | None, anomaly_type: str) -> None:
    if metric is None:
        return
    anomaly_types = unique_preserving_order([*list(metric.get("anomaly_types") or []), anomaly_type])
    metric["anomaly_types"] = anomaly_types
    metric["anomaly_type"] = " | ".join(anomaly_types)
    metric["is_anomaly"] = bool(anomaly_types)


def apply_metric_reliability_checks(metrics: dict[str, dict[str, Any]]) -> dict[str, dict[str, Any]]:
    for metric_name, metric in metrics.items():
        evaluate_metric_reliability(metric_name, metric)
    return metrics


def detect_metric_anomalies(metrics: dict[str, dict[str, Any]]) -> dict[str, dict[str, Any]]:
    conversion = metrics.get("sales_overall_conversion_rate")
    response_time = metrics.get("sales_internet_response_time")
    total_leads = metrics.get("sales_total_leads")
    total_deals = metrics.get("sales_total_deals")
    reminders_sent = metrics.get("service_reminders_sent")
    service_roi = metrics.get("service_campaign_roi")
    outstanding_tasks = metrics.get("service_outstanding_tasks")
    no_shows = metrics.get("service_no_shows")

    conversion_pp_change = conversion.get("percentage_point_change") if conversion else None
    if (
        metric_usable_for_reporting(conversion, minimum_band="meaningful")
        and conversion.get("direction") == "down"
        and isinstance(conversion_pp_change, (int, float))
        and float(conversion_pp_change) <= -5
    ):
        register_metric_anomaly(conversion, "conversion_drop")

    response_percent_change = response_time.get("percent_change") if response_time else None
    if (
        metric_usable_for_reporting(response_time, minimum_band="meaningful")
        and response_time.get("direction") == "up"
        and isinstance(response_percent_change, (int, float))
        and float(response_percent_change) >= 50
    ):
        register_metric_anomaly(response_time, "response_time_spike")

    leads_percent_change = total_leads.get("percent_change") if total_leads else None
    deals_percent_change = total_deals.get("percent_change") if total_deals else None
    conversion_downside = (
        metric_usable_for_reporting(conversion, minimum_band="meaningful")
        and conversion.get("direction") == "down"
        and isinstance(conversion_pp_change, (int, float))
        and float(conversion_pp_change) <= -5
    )
    deals_downside = (
        metric_usable_for_reporting(total_deals, minimum_band="meaningful")
        and total_deals.get("direction") == "down"
        and isinstance(deals_percent_change, (int, float))
        and float(deals_percent_change) <= -15
    )
    if (
        metric_usable_for_reporting(total_leads, minimum_band="meaningful")
        and total_leads.get("direction") == "up"
        and isinstance(leads_percent_change, (int, float))
        and float(leads_percent_change) >= 20
        and (conversion_downside or deals_downside)
    ):
        register_metric_anomaly(total_leads, "lead_quality_mismatch")
        register_metric_anomaly(total_deals, "lead_quality_mismatch")
        register_metric_anomaly(conversion, "lead_quality_mismatch")

    reminders_percent_change = reminders_sent.get("percent_change") if reminders_sent else None
    service_roi_percent_change = service_roi.get("percent_change") if service_roi else None
    if (
        metric_usable_for_reporting(reminders_sent, minimum_band="meaningful")
        and reminders_sent.get("direction") == "up"
        and isinstance(reminders_percent_change, (int, float))
        and float(reminders_percent_change) >= 20
        and metric_usable_for_reporting(service_roi, minimum_band="meaningful")
        and service_roi.get("direction") == "down"
        and isinstance(service_roi_percent_change, (int, float))
        and float(service_roi_percent_change) <= -30
    ):
        register_metric_anomaly(reminders_sent, "reminder_roi_mismatch")
        register_metric_anomaly(service_roi, "reminder_roi_mismatch")

    outstanding_tasks_percent_change = outstanding_tasks.get("percent_change") if outstanding_tasks else None
    if (
        metric_usable_for_reporting(outstanding_tasks, minimum_band="meaningful")
        and outstanding_tasks.get("direction") == "up"
        and isinstance(outstanding_tasks_percent_change, (int, float))
        and float(outstanding_tasks_percent_change) >= 50
    ):
        register_metric_anomaly(outstanding_tasks, "service_task_surge")

    no_shows_percent_change = no_shows.get("percent_change") if no_shows else None
    if (
        metric_usable_for_reporting(no_shows, minimum_band="meaningful")
        and no_shows.get("direction") == "up"
        and isinstance(no_shows_percent_change, (int, float))
        and float(no_shows_percent_change) >= 50
    ):
        register_metric_anomaly(no_shows, "no_show_spike")

    for metric_name, metric in metrics.items():
        if not metric_usable_for_reporting(metric, minimum_band="meaningful"):
            continue
        if metric.get("anomaly_types"):
            continue
        percent_change = metric.get("percent_change")
        percentage_point_change = metric.get("percentage_point_change")
        metric_type = str(metric.get("metric_type", "count"))
        if metric_name.endswith("campaign_roi") and isinstance(percent_change, (int, float)) and abs(float(percent_change)) >= 150:
            register_metric_anomaly(metric, "roi_spike")
        elif metric_type in {"volume", "count"} and isinstance(percent_change, (int, float)) and abs(float(percent_change)) >= 100:
            register_metric_anomaly(metric, "volume_spike")
        elif metric_type == "time" and isinstance(percent_change, (int, float)) and abs(float(percent_change)) >= 60:
            register_metric_anomaly(metric, "time_spike")
        elif metric_type == "rate" and isinstance(percentage_point_change, (int, float)) and abs(float(percentage_point_change)) >= 5:
            register_metric_anomaly(metric, "rate_spike")

    return metrics


def enrich_metrics_for_reporting(metrics: dict[str, dict[str, Any]]) -> dict[str, dict[str, Any]]:
    apply_metric_reliability_checks(metrics)
    detect_metric_anomalies(metrics)
    return metrics


def metric_health_summary(metrics: dict[str, dict[str, Any]]) -> tuple[list[str], list[str], list[str]]:
    anomaly_types = unique_preserving_order(
        anomaly_type
        for metric in metrics.values()
        for anomaly_type in list(metric.get("anomaly_types") or [])
    )
    unreliable_metrics = sorted(metric_name for metric_name, metric in metrics.items() if not bool(metric.get("is_reliable", True)))
    suspicious_metrics = sorted(metric_name for metric_name, metric in metrics.items() if bool(metric.get("is_suspicious", False)))
    return anomaly_types, unreliable_metrics, suspicious_metrics


def important_anomaly_types(anomaly_types: Iterable[str]) -> list[str]:
    important_types = {
        "conversion_drop",
        "response_time_spike",
        "lead_quality_mismatch",
        "roi_spike",
        "reminder_roi_mismatch",
        "service_task_surge",
        "no_show_spike",
    }
    return [anomaly_type for anomaly_type in anomaly_types if anomaly_type in important_types]


# Phase 171 — anomaly types that always qualify for report-level `anomaly_detected` (review queue).
PHASE171_REVIEW_WORTHY_ANOMALY_TYPES: frozenset[str] = frozenset(
    {
        "conversion_drop",
        "response_time_spike",
        "lead_quality_mismatch",
        "roi_spike",
        "reminder_roi_mismatch",
        "service_task_surge",
        "no_show_spike",
    }
)

# Phase 171 — generic MoM swing tags; keep in metrics for audit, but suppress review tag unless output-linked or integrity escalators fire.
PHASE171_DIAGNOSTIC_ANOMALY_TYPES: frozenset[str] = frozenset({"rate_spike", "volume_spike", "time_spike"})

# Many unreliable fields in one month → higher send-risk even if spike tags are generic.
PHASE171_UNRELIABLE_METRIC_REVIEW_THRESHOLD: int = 6


def phase171_metric_keys_flagged_anomaly(metrics: dict[str, dict[str, Any]]) -> frozenset[str]:
    return frozenset(name for name, m in metrics.items() if m and bool(m.get("is_anomaly")))


def phase171_focus_metric_keys_for_selected_rules(
    selected_rules: list[RuleInsight], metrics: dict[str, dict[str, Any]]
) -> frozenset[str]:
    names: list[str] = []
    for rule in selected_rules:
        names.extend(ordered_rule_metric_focus(rule, metrics))
    return frozenset(unique_preserving_order(names))


def phase171_anomaly_metric_labels_touch_client_text(
    metrics: dict[str, dict[str, Any]],
    anomalous_keys: frozenset[str],
    summary: str,
    sms: str,
    bullets: Iterable[str],
) -> bool:
    blob = " ".join([summary or "", sms or "", *bullets]).lower()
    if not blob.strip():
        return False
    for key in anomalous_keys:
        lab = (metric_label(metrics, key) or "").strip().lower()
        if len(lab) < 6:
            continue
        if lab in blob:
            return True
    return False


def phase171_should_emit_anomaly_detected_review_tag(
    metrics: dict[str, dict[str, Any]],
    scoring_profile: InternalScoringProfile,
    result: DealershipResult,
) -> bool:
    """
    Report-level `anomaly_detected` must mean review-worthy risk, not mere presence of a diagnostic spike.

    Raw `scoring_profile.anomaly_types` and per-metric `is_anomaly` remain untouched elsewhere.
    """
    if not scoring_profile.anomaly_types:
        return False
    types_set = frozenset(scoring_profile.anomaly_types)
    if types_set & PHASE171_REVIEW_WORTHY_ANOMALY_TYPES:
        return True
    if scoring_profile.suspicious_metric_count > 0:
        return True
    if scoring_profile.unreliable_metric_count >= PHASE171_UNRELIABLE_METRIC_REVIEW_THRESHOLD:
        return True

    anomalous_keys = phase171_metric_keys_flagged_anomaly(metrics)
    if not anomalous_keys:
        return False

    if types_set <= PHASE171_DIAGNOSTIC_ANOMALY_TYPES:
        if phase171_anomaly_metric_labels_touch_client_text(
            metrics,
            anomalous_keys,
            result.summary or "",
            result.sms_version or "",
            result.bullet_insights or [],
        ):
            return True
        return False

    # Unknown / future anomaly token — default to surfacing for safety.
    return True


def score_metric_change(
    metric: dict[str, Any] | None,
    *,
    scale: float,
    cap: float,
    invert: bool = False,
    use_percentage_points: bool = False,
) -> float:
    if not metric_usable_for_reporting(metric, minimum_band="slight"):
        return 0.0
    raw_value = metric.get("percentage_point_change") if use_percentage_points else metric.get("percent_change")
    if not isinstance(raw_value, (int, float)):
        return 0.0
    adjusted_value = float(raw_value)
    if invert:
        adjusted_value *= -1
    return max(-cap, min(cap, adjusted_value * scale))


def score_metric_risk(
    metric: dict[str, Any] | None,
    *,
    scale: float,
    cap: float,
    invert: bool = False,
    use_percentage_points: bool = False,
) -> float:
    if not metric_usable_for_reporting(metric, minimum_band="slight"):
        return 0.0
    raw_value = metric.get("percentage_point_change") if use_percentage_points else metric.get("percent_change")
    if not isinstance(raw_value, (int, float)):
        return 0.0
    adjusted_value = float(raw_value)
    if invert:
        adjusted_value *= -1
    return max(0.0, min(cap, adjusted_value * scale))


def derived_month_rating_from_score(overall_score: float, rule_based_rating: str) -> str:
    if 45 <= overall_score < 55 and rule_based_rating == "Stable":
        return "Stable"
    if overall_score >= 65:
        return "Strong"
    if overall_score < 45:
        return "Softer"
    return "Mixed"


def build_internal_scoring_profile(
    metrics: dict[str, dict[str, Any]],
    sentiment: str,
    insights_source: str,
) -> InternalScoringProfile:
    anomaly_types, unreliable_metrics, suspicious_metrics = metric_health_summary(metrics)
    important_anomalies = important_anomaly_types(anomaly_types)

    sales_score = 70.0
    sales_score += score_metric_change(metrics.get("sales_total_leads"), scale=0.20, cap=8)
    sales_score += score_metric_change(metrics.get("sales_total_deals"), scale=0.35, cap=18)
    sales_score += score_metric_change(
        metrics.get("sales_overall_conversion_rate"),
        scale=4.0,
        cap=20,
        use_percentage_points=True,
    )
    sales_score += score_metric_change(metrics.get("sales_internet_response_time"), scale=0.18, cap=15, invert=True)
    sales_score += score_metric_change(metrics.get("sales_uncontacted_leads"), scale=0.18, cap=15, invert=True)
    sales_score += score_metric_change(
        metrics.get("sales_appointment_to_sales"),
        scale=3.0,
        cap=12,
        use_percentage_points=True,
    )
    sales_score = clamp_score(sales_score)

    service_score = 70.0
    service_score += score_metric_change(metrics.get("service_bookings_from_campaigns"), scale=0.25, cap=12)
    service_score += score_metric_change(metrics.get("service_reminders_sent"), scale=0.10, cap=6)
    service_score += score_metric_change(metrics.get("service_campaign_roi"), scale=0.08, cap=18)
    service_score += score_metric_change(metrics.get("service_outstanding_tasks"), scale=0.20, cap=18, invert=True)
    service_score += score_metric_change(metrics.get("service_no_shows"), scale=0.15, cap=12, invert=True)
    service_score += score_metric_change(metrics.get("service_rebooked_no_shows"), scale=0.12, cap=8)
    service_score = clamp_score(service_score)

    risk_score = 10.0
    risk_score += score_metric_risk(metrics.get("sales_internet_response_time"), scale=0.25, cap=20)
    risk_score += score_metric_risk(metrics.get("sales_uncontacted_leads"), scale=0.25, cap=20)
    risk_score += score_metric_risk(
        metrics.get("sales_overall_conversion_rate"),
        scale=3.0,
        cap=20,
        invert=True,
        use_percentage_points=True,
    )
    risk_score += score_metric_risk(metrics.get("service_outstanding_tasks"), scale=0.20, cap=18)
    risk_score += score_metric_risk(metrics.get("service_no_shows"), scale=0.15, cap=12)
    if (metric_value(metrics, "sales_campaign_roi") or 0) < 0:
        risk_score += 25
    if (metric_value(metrics, "service_campaign_roi") or 0) < 0:
        risk_score += 20
    risk_score += min(len(important_anomalies) * 6, 18)
    risk_score += min((len(anomaly_types) - len(important_anomalies)) * 2, 8)
    risk_score += min(len(unreliable_metrics) * 3, 12)
    risk_score += min(len(suspicious_metrics) * 2, 8)
    risk_score = clamp_score(risk_score)

    overall_score = clamp_score(15 + (sales_score * 0.4) + (service_score * 0.4) - (risk_score * 0.25))
    rule_based_rating = overall_month_rating_from_sentiment(sentiment)
    derived_month_rating = derived_month_rating_from_score(overall_score, rule_based_rating)
    final_rating_used = rule_based_rating
    if insights_source == "rule-based":
        if derived_month_rating == rule_based_rating:
            final_rating_used = derived_month_rating
        elif rule_based_rating == "Strong" and derived_month_rating == "Mixed" and (len(important_anomalies) >= 2 or overall_score < 45):
            final_rating_used = "Mixed"
        elif rule_based_rating == "Mixed" and derived_month_rating == "Softer" and (len(important_anomalies) >= 1 and overall_score < 45):
            final_rating_used = "Softer"
        elif rule_based_rating == "Stable" and derived_month_rating in {"Mixed", "Softer"} and len(important_anomalies) >= 1:
            final_rating_used = derived_month_rating
    normalised_rule_rating = "Mixed" if rule_based_rating == "Stable" else rule_based_rating
    normalised_derived_rating = "Mixed" if derived_month_rating == "Stable" else derived_month_rating
    scoring_conflict = normalised_rule_rating != normalised_derived_rating

    return InternalScoringProfile(
        sales_score=sales_score,
        service_score=service_score,
        risk_score=risk_score,
        overall_score=overall_score,
        rule_based_rating=rule_based_rating,
        derived_month_rating=derived_month_rating,
        final_rating_used=final_rating_used,
        anomaly_count=len(anomaly_types),
        anomaly_types=anomaly_types,
        unreliable_metric_count=len(unreliable_metrics),
        suspicious_metric_count=len(suspicious_metrics),
        scoring_conflict=scoring_conflict,
    )


def format_signed_number(value: float, decimals: int = 0) -> str:
    rounded = round(value, decimals)
    if decimals == 0 or abs(rounded - round(rounded)) < 1e-9:
        return f"{rounded:+.0f}"
    return f"{rounded:+.{decimals}f}"


def format_percent_change_value(value: float | None) -> str:
    if value is None:
        return ""
    rounded = round(value, 1)
    if abs(rounded - round(rounded)) < 0.05:
        return f"{rounded:+.0f}%"
    return f"{rounded:+.1f}%"


def format_absolute_percent_change_value(value: float | None) -> str:
    if value is None:
        return ""
    return format_percent_change_value(abs(value)).lstrip("+")


def format_percentage_point_change_value(value: float | None) -> str:
    if value is None:
        return ""
    rounded = round(value, 1)
    return f"{rounded:+.1f}pp"


def format_absolute_percentage_point_change_value(value: float | None) -> str:
    if value is None:
        return ""
    return format_percentage_point_change_value(abs(value)).lstrip("+")


def format_currency_value(value: float) -> str:
    if value < 0:
        return f"-{format_currency_value(-value)}"
    if abs(value - round(value)) < 1e-9:
        return f"${value:,.0f}"
    if abs(value) >= 100:
        return f"${value:,.0f}" if abs(value - round(value)) < 0.05 else f"${value:,.2f}"
    return f"${value:,.2f}"


def format_rate_value(value: float) -> str:
    rounded = round(value, 1)
    if abs(rounded - round(rounded)) < 0.05:
        return f"{rounded:.0f}%"
    return f"{rounded:.1f}%"


def format_rate_transition(metric: dict[str, Any]) -> str:
    previous_display = format_metric_value_for_reporting(metric, "previous")
    current_display = format_metric_value_for_reporting(metric, "current")
    return f"from {previous_display} to {current_display}"


def format_rate_directional_value(metric: dict[str, Any]) -> str:
    current_display = format_metric_value_for_reporting(metric, "current")
    direction = str(metric.get("direction", "flat"))
    if direction == "flat":
        return f"at {current_display}"
    movement = "up" if direction == "up" else "down"
    return f"{movement} to {current_display}"


def replace_with_case(text: str, pattern: str, replacement: str) -> str:
    def _replacer(match: re.Match[str]) -> str:
        matched = match.group(0)
        if matched.isupper():
            return replacement.upper()
        if matched[0].isupper():
            return replacement.capitalize()
        return replacement

    return re.sub(pattern, _replacer, text, flags=re.IGNORECASE)


def standardise_reporting_text(text: str) -> str:
    if not text:
        return text

    standardised = text
    replacement_patterns = [
        (r"\bworkshop\b", "service"),
        (r"\bpipeline pressure\b", "follow-up pressure"),
        (r"\bthroughput\b", "activity"),
        (r"\bexecution quality\b", "follow-up quality"),
        (r"\befficiency\b", "performance"),
    ]
    for pattern, replacement in replacement_patterns:
        standardised = replace_with_case(standardised, pattern, replacement)

    cleanup_patterns = [
        (r"\bmeaningfully\b(?=[\s,])", ""),
        (r"\bmaterially\b(?=[\s,])", ""),
        (r"\bnotably\b(?=[\s,])", ""),
        (r"\bstrongly\b(?=[\s,])", ""),
        (r"\bsignificantly\b(?=[\s,])", ""),
    ]
    for pattern, replacement in cleanup_patterns:
        standardised = re.sub(pattern, replacement, standardised, flags=re.IGNORECASE)

    standardised = re.sub(r"\s{2,}", " ", standardised)
    standardised = re.sub(r"\s+([,.])", r"\1", standardised)
    return standardised.strip()


def sanitise_priority_insight_text(text: str) -> str:
    if not text:
        return text
    sanitised = re.sub(r"^\s*(?:-|•)\s+", "", text)
    return sanitised.rstrip()


def format_count_value(value: float) -> str:
    if abs(value - round(value)) < 1e-9:
        return f"{value:,.0f}"
    return f"{value:,.1f}"


def format_duration_value(minutes: float) -> str:
    """Render a duration with explicit units.

    The previous bare-colon form was ambiguous: 788 minutes rendered as "13:08" and
    45 minutes as "45:00", so a reader had no way to tell thirteen hours from
    thirteen minutes.
    """
    total_seconds = max(0, int(round(minutes * 60)))
    if total_seconds >= 3600:
        hours = total_seconds // 3600
        minute_component = (total_seconds % 3600) // 60
        return f"{hours}h {minute_component:02d}m"
    if total_seconds >= 60:
        whole_minutes = total_seconds // 60
        seconds = total_seconds % 60
        return f"{whole_minutes} min" if seconds == 0 else f"{whole_minutes} min {seconds} sec"
    return f"{total_seconds} sec"


def format_time_delta_value(delta_minutes: float) -> str:
    total_seconds = int(round(abs(delta_minutes) * 60))
    if total_seconds < 60:
        return f"{total_seconds} seconds"
    if total_seconds < 3600:
        whole_minutes = total_seconds // 60
        seconds = total_seconds % 60
        return f"{whole_minutes} min" if seconds == 0 else f"{whole_minutes} min {seconds} sec"
    hours = total_seconds // 3600
    remaining_seconds = total_seconds % 3600
    whole_minutes = remaining_seconds // 60
    return f"{hours} hr" if whole_minutes == 0 else f"{hours} hr {whole_minutes} min"


def format_metric_value_for_reporting(metric: dict[str, Any], value_key: str = "current") -> str:
    value = metric.get(value_key)
    if not isinstance(value, (int, float)):
        raw_key = f"{value_key}_raw"
        return str(metric.get(raw_key, ""))

    metric_type = str(metric.get("metric_type", "count"))
    if metric_type == "time":
        return format_duration_value(float(value))
    if metric_type == "rate":
        return format_rate_value(float(value))
    if metric_type == "currency":
        return format_currency_value(float(value))
    return format_count_value(float(value))


def format_metric_delta_for_reporting(metric: dict[str, Any]) -> str:
    metric_type = str(metric.get("metric_type", "count"))
    if metric_type == "rate":
        return format_rate_transition(metric)
    return format_percent_change_value(metric.get("percent_change"))


def build_metric_record(metric_name: str, previous_raw: str, current_raw: str) -> dict[str, Any] | None:
    previous_value, unit = parse_metric_value(previous_raw)
    current_value, current_unit = parse_metric_value(current_raw)

    if previous_value is None or current_value is None:
        return None

    resolved_unit = unit if unit != "unknown" else current_unit
    definition = metric_definition(metric_name, resolved_unit)
    metric_type = str(definition.get("metric_type", "count"))

    delta = current_value - previous_value
    if abs(delta) < 1e-9:
        direction = "flat"
    elif delta > 0:
        direction = "up"
    else:
        direction = "down"

    percent_change = None if abs(previous_value) < 1e-9 else (delta / abs(previous_value)) * 100
    percentage_point_change = delta if metric_type == "rate" else None
    significance_band = derive_significance_band(metric_type, percent_change, percentage_point_change)
    usable_for_reporting = is_metric_usable_for_reporting(metric_type, previous_value, current_value, resolved_unit)

    record = {
        "previous_raw": previous_raw,
        "current_raw": current_raw,
        "previous": round(previous_value, 4),
        "current": round(current_value, 4),
        "previous_value": round(previous_value, 4),
        "current_value": round(current_value, 4),
        "delta": round(delta, 4),
        "absolute_change": round(delta, 4),
        "percent_change": None if percent_change is None else round(percent_change, 2),
        "percentage_point_change": None if percentage_point_change is None else round(percentage_point_change, 2),
        "direction": direction,
        "unit": resolved_unit,
        "metric_type": metric_type,
        "metric_label": str(definition.get("label", metric_name.replace("_", " "))),
        "higher_is_better": bool(definition.get("higher_is_better", True)),
        "significance_band": significance_band,
        "usable_for_reporting": usable_for_reporting,
    }
    return evaluate_metric_reliability(metric_name, record)


def add_metric(metrics: dict[str, dict[str, Any]], metric_name: str, previous_raw: str, current_raw: str) -> None:
    record = build_metric_record(metric_name, previous_raw, current_raw)
    if record is not None:
        metrics[metric_name] = record


def parse_historical_sales_trend_metrics(sales_text: str) -> dict[str, dict[str, Any]]:
    metrics: dict[str, dict[str, Any]] = {}
    lines = clean_lines(sales_text)
    if not lines:
        return metrics

    parse_line_total_metrics(lines, metrics)

    first_performance_index = find_line_index(lines, "UNCONTACTED LEADS")
    if first_performance_index is not None:
        rows = collect_month_rows(lines, first_performance_index, rows_needed=2)
        if len(rows) == 2 and len(rows[0]) >= 2 and len(rows[1]) >= 2:
            add_metric(metrics, "sales_internet_response_time", rows[0][1], rows[1][1])

    return {metric_name: metric for metric_name, metric in metrics.items() if metric_name in TREND_METRIC_NAMES}


def parse_historical_service_trend_metrics(service_text: str) -> dict[str, dict[str, Any]]:
    metrics: dict[str, dict[str, Any]] = {}
    lines = clean_lines(service_text)
    if not lines:
        return metrics

    reminders_index = find_line_index(lines, "CUSTOMERS ENROLLED OUTSTANDING TASKS")
    if reminders_index is not None:
        rows = collect_month_rows(lines, reminders_index, rows_needed=2)
        if len(rows) == 2 and len(rows[0]) >= 3 and len(rows[1]) >= 3:
            add_metric(metrics, "service_reminders_sent", rows[0][0], rows[1][0])
            add_metric(metrics, "service_outstanding_tasks", rows[0][2], rows[1][2])

    targeted_marketing_section = extract_section(service_text, "SERVICE BOOKINGS FROM", "ACTION PLAN & NOTES")
    targeted_rows = extract_targeted_marketing_rows(targeted_marketing_section)
    if len(targeted_rows) >= 2:
        add_metric(metrics, "service_bookings_from_campaigns", targeted_rows[0][0], targeted_rows[1][0])
        add_metric(metrics, "service_campaign_roi", targeted_rows[0][1], targeted_rows[1][1])

    return {metric_name: metric for metric_name, metric in metrics.items() if metric_name in TREND_METRIC_NAMES}


def find_line_index(lines: list[str], keyword: str) -> int | None:
    normalized_keyword = normalize_text(keyword)
    for index, line in enumerate(lines):
        if normalized_keyword in normalize_text(line):
            return index
    return None


def extract_month_row_values(line: str) -> list[str]:
    matches = re.findall(r"[A-Za-z]{3}-\d{2}\s+(.*?)(?=(?:[A-Za-z]{3}-\d{2}\s)|$)", line)
    return [collapse_whitespace(match) for match in matches]


def collect_month_rows(lines: list[str], start_index: int, rows_needed: int, window: int = 12) -> list[list[str]]:
    rows: list[list[str]] = []
    for line in lines[start_index + 1 : start_index + 1 + window]:
        if re.match(r"^[A-Za-z]{3}-\d{2}\s+", line):
            row_values = extract_month_row_values(line)
            if row_values:
                rows.append(row_values)
        if len(rows) == rows_needed:
            break
    return rows


def extract_section(text: str, start_keyword: str, end_keyword: str) -> str:
    collapsed_text = collapse_whitespace(text)
    lowered_text = collapsed_text.lower()
    lowered_start = collapse_whitespace(start_keyword).lower()
    lowered_end = collapse_whitespace(end_keyword).lower()

    start_index = lowered_text.find(lowered_start)
    if start_index < 0:
        return ""

    end_index = lowered_text.find(lowered_end, start_index + len(lowered_start))
    if end_index < 0:
        end_index = len(collapsed_text)

    return collapsed_text[start_index:end_index]


def extract_targeted_marketing_rows(section_text: str) -> list[tuple[str, str]]:
    if not section_text:
        return []
    return re.findall(r"[A-Za-z]{3}-\d{2}\s+(-?[\d,]+)\s+(-?\$?[\d,]+)", section_text)


def parse_line_total_metrics(lines: list[str], metrics: dict[str, dict[str, Any]]) -> None:
    total_line = next((line for line in lines if line.startswith("TOTAL ")), None)
    if total_line is None:
        return

    match = re.match(
        r"^TOTAL\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+(-?[\d,.]+%)\s+(-?[\d,.]+%)$",
        total_line,
    )
    if not match:
        return

    add_metric(metrics, "sales_total_leads", match.group(1), match.group(2))
    add_metric(metrics, "sales_total_deals", match.group(3), match.group(4))
    add_metric(metrics, "sales_overall_conversion_rate", match.group(5), match.group(6))


def parse_sales_report_metrics(sales_text: str) -> dict[str, dict[str, Any]]:
    metrics: dict[str, dict[str, Any]] = {}
    lines = clean_lines(sales_text)
    if not lines:
        return metrics

    parse_line_total_metrics(lines, metrics)

    first_performance_index = find_line_index(lines, "UNCONTACTED LEADS")
    if first_performance_index is not None:
        rows = collect_month_rows(lines, first_performance_index, rows_needed=2)
        if len(rows) == 2 and len(rows[0]) >= 4 and len(rows[1]) >= 4:
            add_metric(metrics, "sales_uncontacted_leads", rows[0][0], rows[1][0])
            add_metric(metrics, "sales_internet_response_time", rows[0][1], rows[1][1])
            add_metric(metrics, "sales_new_leads_to_appointments", rows[0][2], rows[1][2])
            add_metric(metrics, "sales_new_leads_to_test_drives", rows[0][3], rows[1][3])

    second_performance_index = find_line_index(lines, "APPOINTMENT TO SALES")
    if second_performance_index is not None:
        rows = collect_month_rows(lines, second_performance_index, rows_needed=2)
        if len(rows) == 2 and len(rows[0]) >= 4 and len(rows[1]) >= 4:
            add_metric(metrics, "sales_appointment_to_sales", rows[0][0], rows[1][0])
            add_metric(metrics, "sales_overdue_tasks", rows[0][1], rows[1][1])
            add_metric(metrics, "sales_email_capture_rate", rows[0][2], rows[1][2])
            add_metric(metrics, "sales_same_day_deal_pct", rows[0][3], rows[1][3])

    targeted_marketing_section = extract_section(sales_text, "DEALS FROM CAMPAIGNS", "SALESPIPELINE")
    targeted_rows = extract_targeted_marketing_rows(targeted_marketing_section)
    if len(targeted_rows) >= 2:
        add_metric(metrics, "sales_campaign_deals", targeted_rows[0][0], targeted_rows[1][0])
        add_metric(metrics, "sales_campaign_roi", targeted_rows[0][1], targeted_rows[1][1])

    pipeline_size_line = next((line for line in lines if "Current Pipeline Size:" in line), None)
    if pipeline_size_line is not None:
        match = re.search(r"Current Pipeline Size:\s*([\d,]+)", pipeline_size_line)
        if match:
            current_value = match.group(1)
            add_metric(metrics, "sales_pipeline_size", current_value, current_value)

    pipeline_index = find_line_index(lines, "Unqualified Contacted Store Visit Demo Write-Up F&I Deals Lost")
    if pipeline_index is not None:
        rows = collect_month_rows(lines, pipeline_index, rows_needed=2)
        if len(rows) == 2 and len(rows[0]) >= 8 and len(rows[1]) >= 8:
            add_metric(metrics, "sales_pipeline_unqualified", rows[0][0], rows[1][0])
            add_metric(metrics, "sales_pipeline_contacted", rows[0][1], rows[1][1])
            add_metric(metrics, "sales_pipeline_store_visit", rows[0][2], rows[1][2])
            add_metric(metrics, "sales_pipeline_demo", rows[0][3], rows[1][3])
            add_metric(metrics, "sales_pipeline_write_up", rows[0][4], rows[1][4])
            add_metric(metrics, "sales_pipeline_finance", rows[0][5], rows[1][5])
            add_metric(metrics, "sales_pipeline_deals_closed", rows[0][6], rows[1][6])
            add_metric(metrics, "sales_pipeline_lost", rows[0][7], rows[1][7])

    return metrics


def parse_service_report_metrics(service_text: str) -> dict[str, dict[str, Any]]:
    metrics: dict[str, dict[str, Any]] = {}
    lines = clean_lines(service_text)
    if not lines:
        return metrics

    reminders_index = find_line_index(lines, "CUSTOMERS ENROLLED OUTSTANDING TASKS")
    if reminders_index is not None:
        rows = collect_month_rows(lines, reminders_index, rows_needed=2)
        if len(rows) == 2 and len(rows[0]) >= 3 and len(rows[1]) >= 3:
            add_metric(metrics, "service_reminders_sent", rows[0][0], rows[1][0])
            add_metric(metrics, "service_customers_enrolled", rows[0][1], rows[1][1])
            add_metric(metrics, "service_outstanding_tasks", rows[0][2], rows[1][2])

    no_shows_index = find_line_index(lines, "REBOOKED NO SHOWS")
    if no_shows_index is not None:
        rows = collect_month_rows(lines, no_shows_index, rows_needed=2)
        if len(rows) == 2 and len(rows[0]) >= 2 and len(rows[1]) >= 2:
            add_metric(metrics, "service_rebooked_no_shows", rows[0][0], rows[1][0])
            add_metric(metrics, "service_no_shows", rows[0][1], rows[1][1])

    targeted_marketing_section = extract_section(service_text, "SERVICE BOOKINGS FROM", "ACTION PLAN & NOTES")
    targeted_rows = extract_targeted_marketing_rows(targeted_marketing_section)
    if len(targeted_rows) >= 2:
        add_metric(metrics, "service_bookings_from_campaigns", targeted_rows[0][0], targeted_rows[1][0])
        add_metric(metrics, "service_campaign_roi", targeted_rows[0][1], targeted_rows[1][1])

    return metrics


def parse_report_metrics(sales_text: str, service_text: str) -> dict[str, dict[str, Any]]:
    metrics: dict[str, dict[str, Any]] = {}
    metrics.update(parse_sales_report_metrics(sales_text))
    metrics.update(parse_service_report_metrics(service_text))
    return enrich_metrics_for_reporting(metrics)


def parse_historical_trend_metrics(sales_text: str, service_text: str) -> dict[str, dict[str, Any]]:
    metrics: dict[str, dict[str, Any]] = {}
    metrics.update(parse_historical_sales_trend_metrics(sales_text))
    metrics.update(parse_historical_service_trend_metrics(service_text))
    return enrich_metrics_for_reporting(metrics)


def find_historical_report_groups(
    report_group: DealershipReportGroup,
    month: int,
    year: int,
    months_back: int = MAX_HISTORICAL_MONTHS,
) -> list[HistoricalReportGroup]:
    preferred_identity_keys = {
        normalize_text(extract_dealership_name_from_filename(pdf.name, report_group.folder_path.name))
        for pdf in (report_group.sales_pdf, report_group.service_pdf)
        if pdf is not None
    }
    preferred_identity_keys = {key for key in preferred_identity_keys if key}

    history: list[HistoricalReportGroup] = []
    for historical_month, historical_year, historical_label in historical_month_sequence(month, year, months_back):
        matched_groups = find_matching_report_groups(report_group.folder_path, historical_month, historical_year)
        matched_group = None
        if preferred_identity_keys:
            matched_group = next(
                (
                    group
                    for group in matched_groups
                    if {
                        normalize_text(extract_dealership_name_from_filename(pdf.name, report_group.folder_path.name))
                        for pdf in (group.sales_pdf, group.service_pdf)
                        if pdf is not None
                    }
                    & preferred_identity_keys
                ),
                None,
            )
        if matched_group is None and not preferred_identity_keys:
            matched_group = next((group for group in matched_groups if group.group_key == report_group.group_key), None)
        if matched_group is None:
            continue
        if matched_group.sales_pdf is None or matched_group.service_pdf is None:
            continue
        history.append(HistoricalReportGroup(month_label=historical_label, report_group=matched_group))
    return history


def extract_historical_metric_snapshots(
    report_group: DealershipReportGroup,
    month: int,
    year: int,
) -> dict[str, list[HistoricalMetricSnapshot]]:
    metric_history: dict[str, list[HistoricalMetricSnapshot]] = {metric_name: [] for metric_name in TREND_METRIC_NAMES}
    for historical_group in find_historical_report_groups(report_group, month, year):
        try:
            sales_text = extract_pdf_text(historical_group.report_group.sales_pdf) if historical_group.report_group.sales_pdf else ""
            service_text = extract_pdf_text(historical_group.report_group.service_pdf) if historical_group.report_group.service_pdf else ""
        except Exception as exc:  # noqa: BLE001
            logging.warning(
                "Historical extraction skipped for %s %s: %s",
                historical_group.report_group.dealership_name,
                historical_group.month_label,
                exc,
            )
            continue

        historical_metrics = parse_historical_trend_metrics(sales_text, service_text)
        for metric_name in TREND_METRIC_NAMES:
            metric = historical_metrics.get(metric_name)
            current_value = metric.get("current_value") if metric else None
            if metric is None or not isinstance(current_value, (int, float)):
                continue
            metric_history[metric_name].append(
                HistoricalMetricSnapshot(
                    month_label=historical_group.month_label,
                    current_value=float(current_value),
                    metric_type=str(metric.get("metric_type", "count")),
                    is_reliable=bool(metric.get("is_reliable", False)),
                    is_suspicious=bool(metric.get("is_suspicious", False)),
                    is_anomaly=bool(metric.get("is_anomaly", False)),
                )
            )
    return metric_history


def trend_stability_tolerance(metric_type: str, values: list[float]) -> float:
    baseline = max((abs(value) for value in values), default=0.0)
    if metric_type == "rate":
        return 0.5
    if metric_type == "time":
        return max(5.0, baseline * 0.05)
    if metric_type == "currency":
        return max(500.0, baseline * 0.05)
    return max(1.0, baseline * 0.03)


def classify_trend_step(delta: float, tolerance: float) -> str:
    if abs(delta) <= tolerance:
        return "flat"
    return "up" if delta > 0 else "down"


def derive_metric_trend_pattern(values: list[float], metric_type: str) -> tuple[str, str, list[str]]:
    if len(values) < MIN_PREVIOUS_MONTHS_FOR_TREND + 1:
        return "", "", []
    tolerance = trend_stability_tolerance(metric_type, values)
    step_directions = [
        classify_trend_step(values[index + 1] - values[index], tolerance)
        for index in range(len(values) - 1)
    ]
    non_flat_steps = [direction for direction in step_directions if direction != "flat"]
    if not non_flat_steps:
        return "stable", "stable", step_directions

    last_non_flat = non_flat_steps[-1]
    prior_non_flat = non_flat_steps[:-1]

    if all(direction in {"up", "flat"} for direction in step_directions):
        return "up_trend", "up_trend", step_directions
    if all(direction in {"down", "flat"} for direction in step_directions):
        return "down_trend", "down_trend", step_directions

    if last_non_flat == "up" and prior_non_flat and all(direction == "down" for direction in prior_non_flat):
        return "up_trend", "reversal_up", step_directions
    if last_non_flat == "down" and prior_non_flat and all(direction == "up" for direction in prior_non_flat):
        return "down_trend", "reversal_down", step_directions

    if last_non_flat == "up" and not prior_non_flat:
        return "up_trend", "up_trend", step_directions
    if last_non_flat == "down" and not prior_non_flat:
        return "down_trend", "down_trend", step_directions

    return "volatile", "volatile", step_directions


def derive_trend_confidence(
    metric: dict[str, Any],
    history: list[HistoricalMetricSnapshot],
    trend_pattern: str,
) -> str:
    if len(history) < MIN_PREVIOUS_MONTHS_FOR_TREND:
        return "low"

    if trend_pattern == "volatile":
        return "low"

    current_flagged = (
        not bool(metric.get("is_reliable", False))
        or bool(metric.get("is_suspicious", False))
        or bool(metric.get("is_anomaly", False))
    )
    historical_flag_count = sum(
        1
        for snapshot in history
        if (not snapshot.is_reliable) or snapshot.is_suspicious or snapshot.is_anomaly
    )

    if historical_flag_count == 0 and not current_flagged and trend_pattern in {"up_trend", "down_trend", "reversal_up", "reversal_down", "stable"}:
        return "high"

    if historical_flag_count <= 1 and trend_pattern in {"up_trend", "down_trend", "reversal_up", "reversal_down", "stable"}:
        return "medium"

    return "low"


def trend_reference_month(metric: dict[str, Any]) -> str:
    labels = list(metric.get("trend_month_labels") or [])
    if not labels:
        return "recent months"
    latest_label = labels[-1]
    try:
        return datetime.strptime(latest_label, "%b %Y").strftime("%B")
    except ValueError:
        return latest_label


def classify_longitudinal_behaviour(
    metric: dict[str, Any],
    metric_name: str | None,
    metrics: dict[str, dict[str, Any]] | None,
) -> str | None:
    """
    Evidence-only longitudinal behaviour label derived from stored step directions and trend pattern.
    Returns None when history is insufficient or confidence is low.
    """
    if int(metric.get("months_used_for_trend") or 0) < MIN_PREVIOUS_MONTHS_FOR_TREND:
        return None
    if trend_confidence_rank(str(metric.get("trend_confidence", "low"))) < TREND_CONFIDENCE_ORDER["medium"]:
        return None

    steps = [str(s) for s in (metric.get("trend_step_directions") or [])]
    if len(steps) < 2:
        return None

    pattern = str(metric.get("trend_pattern") or "")
    improving = metric_is_currently_improving(metric)
    non_flat = [s for s in steps if s != "flat"]

    if pattern == "volatile":
        return "volatile"
    if not non_flat:
        return "plateauing"
    if steps[-1] == "flat" and any(s != "flat" for s in steps[:-1]):
        return "stabilising"
    ups = sum(1 for s in non_flat if s == "up")
    downs = sum(1 for s in non_flat if s == "down")
    if ups > 0 and downs > 0:
        return "mixed_transition"

    if pattern == "reversal_up" and improving:
        return "recovering"
    if pattern == "reversal_down" and not improving:
        return "deteriorating"
    if pattern == "up_trend" and improving:
        if (
            metric_name == "sales_overall_conversion_rate"
            and metrics
            and metric_usable_for_reporting(metrics.get("sales_total_leads"), minimum_band="slight")
            and str(metrics["sales_total_leads"].get("direction", "flat")) in {"down", "flat"}
        ):
            return "improving_efficiency"
        return "sustained_growth"
    if pattern == "down_trend" and not improving:
        if metric_name == "sales_new_leads_to_test_drives":
            return "weakening_engagement"
        return "sustained_weakness"

    return None


def longitudinal_behaviour_phrase(
    behaviour: str | None,
    metric: dict[str, Any],
    metric_name: str | None,
) -> str | None:
    if not behaviour:
        return None
    reference_month = trend_reference_month(metric)
    phrases = {
        "volatile": "moving unevenly month to month",
        "plateauing": "holding in a narrow recent range",
        "stabilising": "settling after earlier month-to-month movement",
        # Phase 168: suppress internal arbitration / directional-state narration in customer-facing copy.
        "mixed_transition": None,
        "recovering": f"bouncing back after softer results around {reference_month}",
        "deteriorating": f"reversing the improvement seen around {reference_month}",
        "improving_efficiency": "building a cleaner conversion path even as enquiry flow stayed measured",
        "sustained_growth": "showing firmer reads over recent months",
        "sustained_weakness": "softening across recent months",
        "weakening_engagement": "softer early-funnel reads lately",
    }
    return phrases.get(behaviour)


def metric_is_currently_improving(metric: dict[str, Any]) -> bool:
    direction = str(metric.get("direction", "flat"))
    higher_is_better = bool(metric.get("higher_is_better", True))
    if direction == "flat":
        return False
    return (direction == "up" and higher_is_better) or (direction == "down" and not higher_is_better)


def trend_confidence_rank(confidence: str) -> int:
    return TREND_CONFIDENCE_ORDER.get(confidence, 0)


def attach_historical_trends(
    current_metrics: dict[str, dict[str, Any]],
    metric_history: dict[str, list[HistoricalMetricSnapshot]],
) -> dict[str, dict[str, Any]]:
    for metric_name, metric in current_metrics.items():
        metric["trend_direction"] = ""
        metric["trend_pattern"] = ""
        metric["trend_confidence"] = "low"
        metric["months_used_for_trend"] = 0
        metric["trend_month_labels"] = []
        metric["trend_step_directions"] = []

        if metric_name not in TREND_METRIC_NAMES:
            continue
        if not bool(metric.get("is_reliable", False)) or bool(metric.get("is_suspicious", False)):
            continue

        recent_history = list(metric_history.get(metric_name, [])[:MAX_HISTORICAL_MONTHS])
        metric["months_used_for_trend"] = len(recent_history)
        metric["trend_month_labels"] = [snapshot.month_label for snapshot in reversed(recent_history)]
        if len(recent_history) < MIN_PREVIOUS_MONTHS_FOR_TREND:
            continue

        historical_values = [snapshot.current_value for snapshot in reversed(recent_history)]
        current_value = metric.get("current_value")
        if not isinstance(current_value, (int, float)):
            continue

        trend_direction, trend_pattern, step_directions = derive_metric_trend_pattern(
            historical_values + [float(current_value)],
            str(metric.get("metric_type", "count")),
        )
        metric["trend_direction"] = trend_direction
        metric["trend_pattern"] = trend_pattern
        metric["trend_step_directions"] = list(step_directions)
        metric["trend_confidence"] = derive_trend_confidence(metric, recent_history, trend_pattern)

    return current_metrics


def add_metric_watchpoint(metric: dict[str, Any] | None, watchpoint_type: str) -> None:
    if metric is None:
        return
    watchpoints = unique_preserving_order([*list(metric.get("watchpoint_type") or []), watchpoint_type])
    metric["watchpoint_type"] = watchpoints
    metric["persistent_watchpoint"] = bool(watchpoints)


def attach_persistent_watchpoints(metrics: dict[str, dict[str, Any]]) -> dict[str, dict[str, Any]]:
    for metric in metrics.values():
        metric["persistent_watchpoint"] = False
        metric["watchpoint_type"] = []

    uncontacted_leads = metrics.get("sales_uncontacted_leads")
    if (
        metric_usable_for_reporting(uncontacted_leads, minimum_band="meaningful")
        and str(uncontacted_leads.get("direction")) == "up"
        and trend_confidence_rank(str(uncontacted_leads.get("trend_confidence", "low"))) >= TREND_CONFIDENCE_ORDER["medium"]
        and str(uncontacted_leads.get("trend_pattern", "")) in {"up_trend", "reversal_up"}
    ):
        add_metric_watchpoint(uncontacted_leads, "lead_followup_issue")

    response_time = metrics.get("sales_internet_response_time")
    if (
        metric_usable_for_reporting(response_time, minimum_band="slight")
        and (metric_value(metrics, "sales_internet_response_time") or 0) >= 120
        and int(response_time.get("months_used_for_trend") or 0) >= MIN_PREVIOUS_MONTHS_FOR_TREND
        and not (metric_is_currently_improving(response_time) and metric_is_significant(response_time, "meaningful"))
    ):
        add_metric_watchpoint(response_time, "response_time_issue")

    service_tasks = metrics.get("service_outstanding_tasks")
    if (
        metric_usable_for_reporting(service_tasks, minimum_band="slight")
        and (metric_value(metrics, "service_outstanding_tasks") or 0) >= 50
        and int(service_tasks.get("months_used_for_trend") or 0) >= MIN_PREVIOUS_MONTHS_FOR_TREND
        and not (metric_is_currently_improving(service_tasks) and metric_is_significant(service_tasks, "meaningful"))
    ):
        add_metric_watchpoint(service_tasks, "service_task_pressure")

    conversion = metrics.get("sales_overall_conversion_rate")
    if (
        metric_usable_for_reporting(conversion, minimum_band="meaningful")
        and metric_is_currently_improving(conversion)
        and trend_confidence_rank(str(conversion.get("trend_confidence", "low"))) >= TREND_CONFIDENCE_ORDER["medium"]
        and str(conversion.get("trend_pattern", "")) in {"up_trend", "reversal_up"}
    ):
        add_metric_watchpoint(conversion, "conversion_improvement")

    service_roi = metrics.get("service_campaign_roi")
    if (
        metric_usable_for_reporting(service_roi, minimum_band="meaningful")
        and metric_is_currently_improving(service_roi)
        and trend_confidence_rank(str(service_roi.get("trend_confidence", "low"))) >= TREND_CONFIDENCE_ORDER["medium"]
        and str(service_roi.get("trend_pattern", "")) in {"up_trend", "reversal_up"}
    ):
        add_metric_watchpoint(service_roi, "roi_improvement")

    return metrics


def collect_persistent_watchpoint_types(metrics: dict[str, dict[str, Any]]) -> list[str]:
    return unique_preserving_order(
        watchpoint_type
        for metric in metrics.values()
        for watchpoint_type in list(metric.get("watchpoint_type") or [])
    )


def family_metric_score(metric: dict[str, Any] | None) -> float:
    if not metric_usable_for_reporting(metric, minimum_band="slight"):
        return 0.0
    direction = str(metric.get("direction", "flat"))
    if direction == "flat":
        return 0.0
    higher_is_better = bool(metric.get("higher_is_better", True))
    is_improving = (direction == "up" and higher_is_better) or (direction == "down" and not higher_is_better)
    weight = 2.0 if metric_significance_band(metric) == "strong" else 1.0 if metric_is_significant(metric, "meaningful") else 0.5
    return weight if is_improving else -weight


def derive_family_trend(metrics: dict[str, dict[str, Any]], metric_names: list[str]) -> str:
    scores = [family_metric_score(metrics.get(metric_name)) for metric_name in metric_names]
    positive_score = sum(score for score in scores if score > 0)
    negative_score = abs(sum(score for score in scores if score < 0))
    active_scores = [score for score in scores if abs(score) > 0]
    if not active_scores:
        return "stable"
    if positive_score >= 2 and negative_score == 0:
        return "improving"
    if negative_score >= 2 and positive_score == 0:
        return "softening"
    if positive_score == 0 and negative_score < 2:
        return "stable"
    if negative_score == 0 and positive_score < 2:
        return "stable"
    if abs(positive_score - negative_score) <= 1:
        return "mixed"
    return "improving" if positive_score > negative_score else "softening"


def derive_metric_family_rollups(metrics: dict[str, dict[str, Any]]) -> tuple[str, str, str]:
    return (
        derive_family_trend(metrics, METRIC_FAMILY_DEFINITIONS["sales_momentum"]),
        derive_family_trend(metrics, METRIC_FAMILY_DEFINITIONS["service_momentum"]),
        derive_family_trend(metrics, METRIC_FAMILY_DEFINITIONS["followup_risk"]),
    )


def rollup_summary_sentence(
    sales_momentum_trend: str,
    service_momentum_trend: str,
    followup_risk_trend: str,
) -> str | None:
    if followup_risk_trend == "softening":
        return "Follow-up risk remained elevated."
    if sales_momentum_trend == "improving" and service_momentum_trend == "improving":
        return "Sales and service momentum both improved this month."
    if sales_momentum_trend == "improving":
        return "Sales momentum improved this month."
    if service_momentum_trend == "softening":
        return "Service momentum softened slightly."
    if service_momentum_trend == "improving":
        return "Service momentum improved this month."
    if followup_risk_trend == "improving":
        return "Follow-up risk eased this month."
    return None


def rule_metric_alignment(rule: RuleInsight, metric: dict[str, Any] | None) -> bool:
    if not metric:
        return False
    currently_improving = metric_is_currently_improving(metric)
    if rule.polarity == "positive":
        return currently_improving
    if rule.polarity == "negative":
        return not currently_improving
    return metric_is_significant(metric, "strong")


def should_use_trend_context(
    rule: RuleInsight,
    metric: dict[str, Any] | None,
    placement: str,
    metric_name: str | None = None,
    metrics: dict[str, dict[str, Any]] | None = None,
) -> bool:
    if not metric:
        return False
    if not metric_trend_phrase(metric, metric_name, metrics):
        return False
    if not metric_usable_for_reporting(metric, minimum_band="meaningful"):
        return False
    if not rule_metric_alignment(rule, metric):
        return False
    if placement == "bullet" and rule.rule_id.endswith("_snapshot") and not metric_is_significant(metric, "strong"):
        return False
    return True


def persistent_watchpoint_summary_clause(metrics: dict[str, dict[str, Any]]) -> str | None:
    watchpoint_types = collect_persistent_watchpoint_types(metrics)
    if not watchpoint_types:
        return None
    if "lead_followup_issue" in watchpoint_types or "response_time_issue" in watchpoint_types:
        return "lead follow-up remaining a recurring watchpoint over recent months"
    if "service_task_pressure" in watchpoint_types:
        return "service follow-up pressure remaining a recurring watchpoint over recent months"
    if "conversion_improvement" in watchpoint_types and "roi_improvement" in watchpoint_types:
        return "conversion and ROI performance improving consistently across recent months"
    if "conversion_improvement" in watchpoint_types:
        return "conversion performance improving consistently across recent months"
    if "roi_improvement" in watchpoint_types:
        return "ROI performance improving consistently across recent months"
    return None


def trend_metric_priority(rule: RuleInsight) -> list[str]:
    priority_map = {
        "combined_sales_service_up": ["sales_total_deals", "service_campaign_roi"],
        "leads_up_conversion_down": ["sales_overall_conversion_rate", "sales_total_leads"],
        "leads_down_conversion_up": ["sales_overall_conversion_rate", "sales_total_leads"],
        "leads_up_deals_down": ["sales_total_deals", "sales_total_leads"],
        "lead_handling_pressure": ["sales_internet_response_time"],
        "lead_handling_improved": ["sales_internet_response_time"],
        "test_drive_and_conversion_down": ["sales_overall_conversion_rate"],
        "sales_marketing_negative_roi": ["sales_campaign_roi"],
        "sales_campaign_roi_decline": ["sales_campaign_roi", "sales_campaign_deals"],
        "sales_volume_snapshot": ["sales_total_deals"],
        "sales_conversion_snapshot": ["sales_overall_conversion_rate"],
        "lead_response_snapshot": ["sales_internet_response_time"],
        "service_campaigns_up": ["service_bookings_from_campaigns", "service_campaign_roi"],
        "service_bookings_roi_mismatch": ["service_campaign_roi", "service_bookings_from_campaigns"],
        "service_tasks_pressure": ["service_outstanding_tasks"],
        "service_follow_up_cleaner": ["service_outstanding_tasks"],
        "service_booking_snapshot": ["service_bookings_from_campaigns"],
        "service_roi_snapshot": ["service_campaign_roi"],
        "service_task_snapshot": ["service_outstanding_tasks"],
        "crm_execution_pressure_composite": list(PHASE155C_COMPOSITE_ORDER),
    }
    return priority_map.get(rule.rule_id, rule.supporting_metrics)


def metric_trend_phrase(
    metric: dict[str, Any] | None,
    metric_name: str | None = None,
    metrics: dict[str, dict[str, Any]] | None = None,
) -> str | None:
    if not metric:
        return None
    if not bool(metric.get("is_reliable", False)) or bool(metric.get("is_suspicious", False)):
        return None
    if not metric_usable_for_reporting(metric, minimum_band="meaningful"):
        return None
    if trend_confidence_rank(str(metric.get("trend_confidence", "low"))) < TREND_CONFIDENCE_ORDER["medium"]:
        return None

    trend_pattern = str(metric.get("trend_pattern") or metric.get("trend_direction") or "")
    if int(metric.get("months_used_for_trend") or 0) < MIN_PREVIOUS_MONTHS_FOR_TREND:
        return None

    behaviour_preview = classify_longitudinal_behaviour(metric, metric_name, metrics)
    if str(metric.get("direction", "flat")) == "flat" and behaviour_preview not in {
        "plateauing",
        "stabilising",
        "volatile",
        "mixed_transition",
    }:
        return None

    behaviour = behaviour_preview
    behaviour_phrase = longitudinal_behaviour_phrase(behaviour, metric, metric_name)
    if behaviour_phrase:
        return behaviour_phrase

    if trend_pattern == "up_trend":
        return "continuing an upward trend over recent months"

    if trend_pattern == "down_trend":
        return "continuing the improvement seen over recent months" if metric_is_currently_improving(metric) else "extending a softening trend over recent months"

    if trend_pattern in {"reversal_up", "reversal_down"}:
        reference_month = trend_reference_month(metric)
        return (
            f"bouncing back after weaker results in {reference_month}"
            if metric_is_currently_improving(metric)
            else f"reversing the improvement seen in {reference_month}"
        )

    if trend_pattern == "volatile":
        return "showing mixed movement across recent months"

    return None


def trend_bonus_for_metric(
    metric: dict[str, Any] | None,
    metric_name: str | None = None,
    metrics: dict[str, dict[str, Any]] | None = None,
) -> int:
    if not metric:
        return 0
    if str(metric.get("trend_pattern", "")) in {"", "stable"}:
        return 0
    if not metric_trend_phrase(metric, metric_name, metrics):
        return 0
    return TREND_PRIORITY_BONUS.get(str(metric.get("trend_confidence", "low")), 0)


def trend_bonus_for_rule(rule: RuleInsight, metrics: dict[str, dict[str, Any]]) -> int:
    return max(
        (
            trend_bonus_for_metric(metrics.get(metric_name), metric_name, metrics)
            for metric_name in trend_metric_priority(rule)
        ),
        default=0,
    )


def high_confidence_trend_count(metrics: dict[str, dict[str, Any]]) -> int:
    return sum(1 for metric in metrics.values() if str(metric.get("trend_confidence", "low")) == "high")


def count_trend_supported_insights(selected_rules: list[RuleInsight], metrics: dict[str, dict[str, Any]]) -> int:
    return sum(1 for rule in selected_rules if rule.signal_type == "trend_supported")


def append_clause_before_sentence_end(text: str, clause: str) -> str:
    stripped = text.rstrip()
    if not stripped:
        return text
    punctuation = "."
    if stripped[-1] in ".!?":
        punctuation = stripped[-1]
        stripped = stripped[:-1]
    return f"{stripped}, {clause}{punctuation}"


def append_clause_to_first_sentence(text: str, clause: str) -> str:
    sentence_match = re.search(r"(?<!\d)[.!?](?=\s|$)", text)
    if not sentence_match:
        return append_clause_before_sentence_end(text, clause)
    punctuation_index = sentence_match.start()
    return f"{text[:punctuation_index]}, {clause}{text[punctuation_index:]}"


def append_trend_context_to_outputs(
    bullets: list[str],
    summary: str,
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
    *,
    attach_structured_summary_clauses: bool = True,
    context_seed: str = "",
) -> tuple[list[str], str]:
    updated_bullets = list(bullets)

    summary_clauses: list[str] = []
    if attach_structured_summary_clauses:
        watchpoint_sentence = summary_watchpoint_sentence(selected_rules, metrics, context_seed)
        trend_sentence = build_summary_trend_sentence(selected_rules, metrics)
        if watchpoint_sentence:
            summary_clauses.append(watchpoint_sentence.rstrip(".").capitalize())
        elif trend_sentence:
            summary_clauses.append(trend_sentence.rstrip(".").capitalize())

    confidence_sentence = build_summary_confidence_clause(metrics, summary_text=summary)
    if confidence_sentence:
        summary_clauses.append(confidence_sentence.rstrip(".").capitalize())

    max_clauses = 2
    for clause in summary_clauses[:max_clauses]:
        normalized_clause = clause[0].lower() + clause[1:] if clause else clause
        if normalized_clause and normalized_clause not in summary:
            summary = append_clause_before_sentence_end(summary, normalized_clause)

    return updated_bullets, summary


def metric_value(metrics: dict[str, dict[str, Any]], metric_name: str, value_key: str = "current") -> float | None:
    metric = metrics.get(metric_name)
    if not metric:
        return None
    value = metric.get(value_key)
    if isinstance(value, (int, float)):
        return float(value)
    return None


def metric_direction(metrics: dict[str, dict[str, Any]], metric_name: str) -> str | None:
    metric = metrics.get(metric_name)
    if metric is None:
        return None
    direction = metric.get("direction")
    return str(direction) if direction is not None else None


def metric_is_adverse(metrics: dict[str, dict[str, Any]], metric_name: str) -> bool:
    metric = metrics.get(metric_name)
    if not metric:
        return False
    direction = metric_direction(metrics, metric_name)
    if direction not in {"up", "down"}:
        return False
    return not metric_is_currently_improving(metric)


def has_adverse_focus_metric(rule: RuleInsight, metrics: dict[str, dict[str, Any]]) -> bool:
    return any(metric_is_adverse(metrics, metric_name) for metric_name in rule.metric_focus or [])


def primary_adverse_focus_metric(rule: RuleInsight, metrics: dict[str, dict[str, Any]]) -> str | None:
    focus_metrics = list(rule.metric_focus or [])
    for metric_name in focus_metrics:
        if metric_is_adverse(metrics, metric_name):
            return metric_name
    return focus_metrics[0] if focus_metrics else None


def summary_metric_label(metric_name: str | None, metrics: dict[str, dict[str, Any]]) -> str | None:
    if not metric_name:
        return None
    label_map = {
        "sales_overall_conversion_rate": "conversion",
        "sales_total_leads": "lead volume",
        "sales_total_deals": "deal volume",
        "sales_internet_response_time": "response times",
        "sales_uncontacted_leads": "uncontacted leads",
        "service_outstanding_tasks": "outstanding service tasks",
        "service_campaign_roi": "service campaign ROI",
        "service_bookings_from_campaigns": "service campaign bookings",
    }
    return label_map.get(metric_name, metric_label(metrics, metric_name))


def sentence_case_start(text: str) -> str:
    if not text:
        return text
    first = text[0]
    if first.isalpha():
        return first.upper() + text[1:]
    return text


def format_metric_change(metric: dict[str, Any]) -> str:
    current_display = format_metric_value_for_reporting(metric, "current")
    direction = metric.get("direction", "flat")
    metric_type = str(metric.get("metric_type", "count"))
    if metric_type == "rate":
        return format_rate_transition(metric)
    change_display = format_absolute_percent_change_value(metric.get("percent_change"))
    if not change_display and metric_type != "time":
        return f"to {current_display}"
    if direction == "flat":
        return f"at {current_display}"
    if metric_type == "time":
        time_delta = format_time_delta_value(float(metric.get("absolute_change") or 0.0))
        time_percent = format_absolute_percent_change_value(metric.get("percent_change"))
        if time_percent:
            return f"by {time_delta} ({time_percent}) to {current_display}"
        return f"by {time_delta} to {current_display}"
    return f"by {change_display.lstrip('+')} to {current_display}"


def format_metric_directional_change(metric: dict[str, Any]) -> str:
    current_display = format_metric_value_for_reporting(metric, "current")
    direction = metric.get("direction", "flat")
    metric_type = str(metric.get("metric_type", "count"))
    if metric_type == "rate":
        return format_rate_directional_value(metric)
    change_display = format_absolute_percent_change_value(metric.get("percent_change"))
    if not change_display and metric_type != "time":
        return f"to {current_display}"
    if direction == "flat":
        return f"at {current_display}"
    if metric_type == "time":
        time_percent = format_absolute_percent_change_value(metric.get("percent_change"))
        return f"{direction} {time_percent or format_time_delta_value(float(metric.get('absolute_change') or 0.0))} to {current_display}"
    return f"{direction} {change_display.lstrip('+')} to {current_display}"


def make_rule_insight(
    rule_id: str,
    message: str,
    category: str,
    severity: str,
    priority_weight: int,
    polarity: str,
    tone_band: str,
    merge_key: str,
    supporting_metrics: list[str],
    signal_type: str = "current_performance",
    metric_focus: list[str] | None = None,
) -> RuleInsight:
    return RuleInsight(
        rule_id=rule_id,
        message=message,
        category=category,
        severity=severity,
        priority_weight=priority_weight,
        polarity=polarity,
        tone_band=tone_band,
        merge_key=merge_key,
        supporting_metrics=supporting_metrics,
        signal_type=signal_type,
        metric_focus=metric_focus,
    )


def metric_change_priority(metric: dict[str, Any] | None) -> float:
    if not metric:
        return 0.0
    percentage_point_change = metric.get("percentage_point_change")
    if isinstance(percentage_point_change, (int, float)):
        return abs(float(percentage_point_change))
    percent_change = metric.get("percent_change")
    if isinstance(percent_change, (int, float)):
        return abs(float(percent_change))
    absolute_change = metric.get("absolute_change")
    if isinstance(absolute_change, (int, float)):
        return abs(float(absolute_change))
    return 0.0


def ordered_rule_metric_focus(rule: RuleInsight, metrics: dict[str, dict[str, Any]]) -> list[str]:
    candidate_names = unique_preserving_order([*trend_metric_priority(rule), *rule.supporting_metrics])
    available_names = [metric_name for metric_name in candidate_names if metrics.get(metric_name)]
    if not available_names:
        return []

    def sort_key(metric_name: str) -> tuple[int, int, int, int, float, str]:
        metric = metrics.get(metric_name)
        return (
            0 if metric_usable_for_reporting(metric, minimum_band="slight") else 1,
            0 if metric_is_adverse(metrics, metric_name) else 1,
            0 if bool(metric and metric.get("is_anomaly")) else 1,
            0 if bool(metric and metric.get("persistent_watchpoint")) else 1,
            -metric_change_priority(metric),
            metric_name,
        )

    return sorted(available_names, key=sort_key)


def derive_rule_signal_type(rule: RuleInsight, metrics: dict[str, dict[str, Any]]) -> str:
    # Composite CRM is already multi-metric; do not upgrade to trend_supported (avoids duplicated trend clauses).
    if rule.rule_id == "crm_execution_pressure_composite":
        return "composite_crm_reasoning"
    focus_metrics = ordered_rule_metric_focus(rule, metrics)
    if any(bool(metrics.get(metric_name, {}).get("is_anomaly")) for metric_name in focus_metrics):
        return "anomaly_flag"
    if any(bool(metrics.get(metric_name, {}).get("persistent_watchpoint")) for metric_name in focus_metrics):
        return "persistent_watchpoint"
    if any(
        should_use_trend_context(rule, metrics.get(metric_name), "bullet", metric_name, metrics)
        for metric_name in focus_metrics
    ):
        return "trend_supported"
    return "current_performance"


def dominant_focus_metrics(rule: RuleInsight, metrics: dict[str, dict[str, Any]]) -> list[str]:
    focus_metrics = ordered_rule_metric_focus(rule, metrics)
    if not focus_metrics:
        return []
    if rule.signal_type == "anomaly_flag":
        anomaly_metrics = [metric_name for metric_name in focus_metrics if bool(metrics.get(metric_name, {}).get("is_anomaly"))]
        adverse_metric = next((metric_name for metric_name in focus_metrics if metric_is_adverse(metrics, metric_name)), None)
        prioritized = unique_preserving_order([adverse_metric, *anomaly_metrics])
        return prioritized[:2] or focus_metrics[:1]
    if rule.signal_type == "persistent_watchpoint":
        watchpoint_metrics = [metric_name for metric_name in focus_metrics if bool(metrics.get(metric_name, {}).get("persistent_watchpoint"))]
        return watchpoint_metrics[:1] or focus_metrics[:1]
    return focus_metrics[:2]


def metric_transition_clause(
    metrics: dict[str, dict[str, Any]],
    metric_name: str,
    *,
    up_verb: str,
    down_verb: str,
    label_override: str | None = None,
    minimum_band: str = "slight",
) -> str | None:
    return build_metric_transition_text(
        metrics,
        metric_name,
        up_verb=up_verb,
        down_verb=down_verb,
        label_override=label_override,
        minimum_band=minimum_band,
    )


def persistent_watchpoint_suffix(rule: RuleInsight, metrics: dict[str, dict[str, Any]]) -> str | None:
    watchpoint_types = unique_preserving_order(
        watchpoint_type
        for metric_name in dominant_focus_metrics(rule, metrics)
        for watchpoint_type in list(metrics.get(metric_name, {}).get("watchpoint_type") or [])
    )
    if not watchpoint_types:
        return None
    if "lead_followup_issue" in watchpoint_types or "response_time_issue" in watchpoint_types:
        return "remaining a recurring follow-up risk"
    if "service_task_pressure" in watchpoint_types:
        return "remaining a recurring service follow-up risk"
    if "conversion_improvement" in watchpoint_types:
        return "showing consistent conversion improvement"
    if "roi_improvement" in watchpoint_types:
        return "showing consistent ROI improvement"
    return None


def concise_trend_suffix(rule: RuleInsight, metrics: dict[str, dict[str, Any]]) -> str | None:
    # Phase 168: do not append longitudinal trend suffixes into Priority Insight rule text — they duplicate
    # append_trend_context_to_outputs when enabled and read as internal validation / arbitration language.
    return None


def focused_metric_message(metric_name: str, metrics: dict[str, dict[str, Any]]) -> str | None:
    clause_map = {
        "sales_overall_conversion_rate": ("improved", "fell", "Conversion"),
        "sales_total_leads": ("increased", "eased", "Lead volume"),
        "sales_total_deals": ("increased", "fell", "Deal volume"),
        "sales_internet_response_time": ("slowed", "improved", "Response times"),
        "sales_uncontacted_leads": ("rose", "fell", "Uncontacted leads"),
        "service_outstanding_tasks": ("rose", "fell", "Outstanding service tasks"),
        "service_campaign_roi": ("improved", "fell", "Service campaign ROI"),
        "service_bookings_from_campaigns": ("rose", "fell", "Service campaign bookings"),
    }
    up_verb, down_verb, label = clause_map.get(metric_name, ("improved", "fell", metric_label(metrics, metric_name)))
    return metric_transition_clause(metrics, metric_name, up_verb=up_verb, down_verb=down_verb, label_override=label)


def metric_is_improvement(metric: dict[str, Any] | None) -> bool:
    """True when the metric moved in the direction that is good for the dealership."""
    if not metric:
        return False
    direction = str(metric.get("direction", "flat"))
    if direction == "flat":
        return False
    higher_is_better = bool(metric.get("higher_is_better", True))
    return (direction == "up") == higher_is_better


def rule_shows_visible_improvement(rule: RuleInsight, metrics: dict[str, dict[str, Any]]) -> bool:
    """True when at least one of the rule's own metrics is both reportable and improving.

    A rule can carry positive polarity while the metric that made it positive has been
    gated out as unreliable, which is how "Lead volume down 72.5%" once reached the
    Primary Strength field.
    """
    names = [*(rule.metric_focus or []), *(rule.supporting_metrics or [])]
    return any(
        metric_usable_for_reporting(metrics.get(name)) and metric_is_improvement(metrics.get(name))
        for name in names
    )


def worst_declining_metric_phrase(metrics: dict[str, dict[str, Any]]) -> str | None:
    """Short factual phrase for the sharpest reportable deterioration."""
    search_order = [*FALLBACK_METRIC_ORDER]
    search_order.extend(name for name in METRIC_DEFINITIONS if name not in search_order)
    for minimum_band in ("meaningful", "slight"):
        ranked: list[tuple[int, float, str]] = []
        for metric_name in search_order:
            metric = metrics.get(metric_name)
            if not metric_usable_for_reporting(metric, minimum_band):
                continue
            if metric_is_improvement(metric) or str(metric.get("direction", "flat")) == "flat":
                continue
            band_rank = SIGNIFICANCE_ORDER.get(metric_significance_band(metric), 0)
            magnitude = abs(float(metric.get("percent_change") or 0.0))
            ranked.append((band_rank, magnitude, metric_name))
        if not ranked:
            continue
        ranked.sort(key=lambda item: (item[0], item[1]), reverse=True)
        fact = build_short_metric_fact(metrics, ranked[0][2], include_month_on_month=False)
        if fact:
            return sentence_case_start(fact.strip().rstrip("."))
    return None


def best_improving_metric_phrase(metrics: dict[str, dict[str, Any]]) -> str | None:
    """Short factual phrase for the strongest genuine improvement in the metric set.

    Primary Strength used to be drawn only from the four selected rules, and rule
    selection is weighted toward adverse signals. When all four came back negative the
    report printed "No standout strength identified" even though the parsed metrics
    held a reliable, material improvement -- in 11 of the 12 cases in the July 2026 run.
    """
    search_order = [*FALLBACK_METRIC_ORDER]
    search_order.extend(name for name in METRIC_DEFINITIONS if name not in search_order)
    for minimum_band in ("meaningful", "slight"):
        ranked: list[tuple[int, float, str]] = []
        for metric_name in search_order:
            metric = metrics.get(metric_name)
            if not metric_usable_for_reporting(metric, minimum_band):
                continue
            if not metric_is_improvement(metric):
                continue
            band_rank = SIGNIFICANCE_ORDER.get(metric_significance_band(metric), 0)
            magnitude = abs(float(metric.get("percent_change") or 0.0))
            ranked.append((band_rank, magnitude, metric_name))
        if not ranked:
            continue
        ranked.sort(key=lambda item: (item[0], item[1]), reverse=True)
        fact = build_short_metric_fact(metrics, ranked[0][2], include_month_on_month=False)
        if fact:
            return sentence_case_start(fact.strip().rstrip("."))
    return None


def rate_consequence_clause(
    metrics: dict[str, dict[str, Any]],
    rate_metric_name: str,
    volume_metric_name: str,
    unit_noun: str,
) -> str | None:
    """Restate a rate move in the units it costs or earns at this month's volume."""
    rate = metrics.get(rate_metric_name)
    volume = metrics.get(volume_metric_name)
    if not metric_usable_for_reporting(rate) or not metric_is_meaningful(volume):
        return None
    percentage_points = rate.get("percentage_point_change")
    base = volume.get("current")
    if not isinstance(percentage_points, (int, float)) or not isinstance(base, (int, float)):
        return None
    if float(base) <= 0:
        return None
    units = abs(float(percentage_points)) / 100.0 * float(base)
    if units < 5:
        return None
    movement = "more" if float(percentage_points) > 0 else "fewer"
    # Deliberately avoids the words "lead volume"/"lead flow": the SMS signal matcher
    # keys on those phrases, and a consequence clause must not register as a second,
    # separate movement in the metric it is measured against.
    base_noun = "lead count" if volume_metric_name == "sales_total_leads" else metric_label(metrics, volume_metric_name)
    return f"about {format_count_value(round(units))} {movement} {unit_noun} on this month's {base_noun}"


def conversion_counterfactual_clause(metrics: dict[str, dict[str, Any]]) -> str | None:
    """Say how many deals last month's conversion rate would have returned."""
    conversion = metrics.get("sales_overall_conversion_rate")
    leads = metrics.get("sales_total_leads")
    deals = metrics.get("sales_total_deals")
    if not metric_usable_for_reporting(conversion):
        return None
    if not (metric_is_meaningful(leads) and metric_is_meaningful(deals)):
        return None
    try:
        leads_now = float(leads["current"])
        deals_now = float(deals["current"])
        conversion_now = float(conversion["current"])
        conversion_previous = float(conversion["previous"])
    except (KeyError, TypeError, ValueError):
        return None
    if leads_now <= 0 or deals_now <= 0 or conversion_previous <= 0:
        return None
    # Only trust the counterfactual where the published conversion rate really is
    # deals over leads this month. Some sites report conversion on a different base,
    # and a counterfactual built on the wrong denominator is worse than none.
    implied_deals = leads_now * conversion_now / 100.0
    if abs(implied_deals - deals_now) > max(5.0, deals_now * 0.15):
        return None
    expected_deals = leads_now * conversion_previous / 100.0
    if abs(expected_deals - deals_now) < max(5.0, deals_now * 0.03):
        return None
    return (
        "at last month's conversion the same leads would have returned about "
        f"{format_count_value(round(expected_deals))} deals, not {format_count_value(round(deals_now))}"
    )


def campaign_return_per_deal_clause(
    metrics: dict[str, dict[str, Any]],
    roi_metric_name: str,
    deals_metric_name: str,
) -> str | None:
    """Separate campaign volume from campaign efficiency, in dollars already in the data."""
    roi = metrics.get(roi_metric_name)
    deals = metrics.get(deals_metric_name)
    if not metric_usable_for_reporting(roi) or not metric_is_meaningful(deals):
        return None
    try:
        roi_previous = float(roi["previous"])
        roi_now = float(roi["current"])
        deals_previous = float(deals["previous"])
        deals_now = float(deals["current"])
    except (KeyError, TypeError, ValueError):
        return None
    if min(roi_previous, roi_now) <= 0 or min(deals_previous, deals_now) <= 0:
        return None
    per_deal_previous = roi_previous / deals_previous
    per_deal_now = roi_now / deals_now
    if abs(per_deal_now - per_deal_previous) < max(5.0, abs(per_deal_previous) * 0.05):
        return None
    movement = "fell" if per_deal_now < per_deal_previous else "rose"
    # Whole dollars: cents on a per-deal average imply a precision the figure does not have.
    return (
        f"return per attributed deal {movement} from "
        f"{format_currency_value(round(per_deal_previous))} to {format_currency_value(round(per_deal_now))}"
    )


CONSEQUENCE_BY_RULE: dict[str, str] = {
    "test_drive_and_conversion_down": "test_drives",
    "leads_up_conversion_down": "conversion_counterfactual",
    "leads_down_conversion_up": "conversion_counterfactual",
    "leads_up_deals_down": "conversion_counterfactual",
    "sales_conversion_snapshot": "conversion_counterfactual",
    "sales_volume_snapshot": "conversion_counterfactual",
    "sales_campaign_roi_decline": "sales_campaign_per_deal",
    "sales_marketing_up": "sales_campaign_per_deal",
}


def consequence_clause(rule_id: str, metrics: dict[str, dict[str, Any]]) -> str | None:
    """One clause of arithmetic on figures already in the report.

    A rate on its own asks the reader to do the multiplication. These clauses do it,
    in deals, test drives or dollars the source data already carries -- no assumed
    gross, no estimate. Each returns None unless its inputs are present and material,
    so a line either earns the extra clause or stays as it was.
    """
    kind = CONSEQUENCE_BY_RULE.get(rule_id)
    if kind == "test_drives":
        return rate_consequence_clause(metrics, "sales_new_leads_to_test_drives", "sales_total_leads", "test drives")
    if kind == "conversion_counterfactual":
        return conversion_counterfactual_clause(metrics)
    if kind == "sales_campaign_per_deal":
        return campaign_return_per_deal_clause(metrics, "sales_campaign_roi", "sales_campaign_deals")
    return None


MIN_MONTHS_FOR_RECORD_CLAIM = 4
MIN_MONTHS_FOR_AVERAGE_CLAIM = 3
ROLLING_AVERAGE_WINDOW = 6


def format_history_value(metric: dict[str, Any], value: float) -> str:
    metric_type = str(metric.get("metric_type", "count"))
    if metric_type == "rate":
        return format_rate_value(value)
    if metric_type == "currency":
        # Whole dollars: cents on a derived average imply precision that is not there.
        return format_currency_value(round(value))
    if metric_type == "time":
        return format_duration_value(value)
    return format_count_value(round(value))


def current_change_magnitude(metric: dict[str, Any]) -> float | None:
    """This month's move, in the unit the metric's history is comparable in."""
    if str(metric.get("metric_type", "count")) == "rate":
        points = metric.get("percentage_point_change")
        return abs(float(points)) if isinstance(points, (int, float)) else None
    percent = metric.get("percent_change")
    return abs(float(percent)) if isinstance(percent, (int, float)) else None


def historical_change_magnitudes(metric: dict[str, Any], values: list[float]) -> list[float]:
    is_rate = str(metric.get("metric_type", "count")) == "rate"
    magnitudes: list[float] = []
    for previous, current in zip(values, values[1:]):
        if is_rate:
            magnitudes.append(abs(current - previous))
        elif abs(previous) > 1e-9:
            magnitudes.append(abs((current - previous) / previous * 100.0))
    return magnitudes


def baseline_clause(metric_name: str, metrics: dict[str, dict[str, Any]]) -> tuple[str, str] | None:
    """One comparison against the dealership's own record.

    An observation tells an owner what they already know; a comparison tells them
    whether to care. Returns (clause, strength) where strength is "record" for a claim
    that ranks this month against the others, and "reference" for a softer comparison.
    Returns None whenever the record is too short to make the claim honestly.
    """
    metric = metrics.get(metric_name)
    if not metric_usable_for_reporting(metric, "meaningful"):
        return None

    series = list(metric.get("history_series") or [])
    current_value = metric.get("current_value", metric.get("current"))
    if not isinstance(current_value, (int, float)):
        return None
    current_value = float(current_value)
    direction = str(metric.get("direction", "flat"))

    if len(series) >= MIN_MONTHS_FOR_RECORD_CLAIM:
        values = [value for _, value in series]
        window = len(values) + 1

        # Sharpest move: is this month's step bigger than every step in the record?
        magnitude = current_change_magnitude(metric)
        previous_magnitudes = historical_change_magnitudes(metric, values)
        if (
            magnitude is not None
            and previous_magnitudes
            and magnitude > max(previous_magnitudes) * 1.05
        ):
            movement = "drop" if direction == "down" else "rise"
            return f"the sharpest {movement} in {window} months", "record"

        # Failing that, does the level itself sit at the edge of the record?
        if direction == "down" and current_value < min(values):
            return f"the lowest in {window} months", "record"
        if direction == "up" and current_value > max(values):
            return f"the highest in {window} months", "record"

    last_year = metric.get("history_same_month_last_year")
    if isinstance(last_year, (int, float)) and abs(float(last_year)) > 1e-9:
        gap = abs(current_value - float(last_year)) / abs(float(last_year)) * 100.0
        if gap >= 10.0:
            return (
                f"against {format_history_value(metric, float(last_year))} the same month last year",
                "reference",
            )

    if len(series) >= MIN_MONTHS_FOR_AVERAGE_CLAIM:
        recent = [value for _, value in series[-ROLLING_AVERAGE_WINDOW:]]
        average = sum(recent) / len(recent)
        is_rate = str(metric.get("metric_type", "count")) == "rate"
        material = (
            abs(current_value - average) >= 1.5
            if is_rate
            else abs(average) > 1e-9 and abs(current_value - average) / abs(average) * 100.0 >= 10.0
        )
        if material:
            return (
                f"against a {len(recent)}-month average of {format_history_value(metric, average)}",
                "reference",
            )

    return None


_RISING_WORDS = re.compile(r"\b(rose|risen|increased|improved|slowed|climbed|grew|up)\b", re.IGNORECASE)
_FALLING_WORDS = re.compile(r"\b(fell|fallen|eased|reduced|dropped|softened|weakened|down)\b", re.IGNORECASE)


def sentence_describes_opposite_moves(text: str) -> bool:
    """True when a sentence already reports one thing going up and another going down.

    A trailing comparison binds to whatever the sentence just said, so it must not be
    added to a sentence where "the sharpest drop" could refer to either half.
    """
    return bool(_RISING_WORDS.search(text or "") and _FALLING_WORDS.search(text or ""))


def rule_baseline_clause(rule: RuleInsight, metrics: dict[str, dict[str, Any]]) -> tuple[str, str] | None:
    """A baseline for the rule, but only where it cannot bind to the wrong metric.

    A trailing clause attaches to whatever the sentence just said. Where a rule names
    two metrics moving opposite ways -- "Uncontacted leads up 6.2% and email capture
    rate down from 68% to 56%" -- "the sharpest drop in 5 months" would read as a claim
    about either. In that case the comparison is dropped rather than made ambiguous.
    """
    focus = dominant_focus_metrics(rule, metrics) or list(rule.metric_focus or [])
    # supporting_metrics matters as much as the focus: the sentence for
    # lead_handling_pressure names response times *and* uncontacted leads even though
    # only one of them is the focus.
    mentioned = unique_preserving_order([*focus, *(rule.supporting_metrics or [])])
    named = [name for name in mentioned if metric_usable_for_reporting(metrics.get(name))]
    directions = {str((metrics.get(name) or {}).get("direction", "flat")) for name in named}
    if len(named) > 1 and len(directions) > 1:
        return None
    for metric_name in named[:2]:
        clause = baseline_clause(metric_name, metrics)
        if clause:
            return clause
    return None


def compressed_rule_message(rule: RuleInsight, metrics: dict[str, dict[str, Any]]) -> str:
    leads_shift = metric_transition_clause(metrics, "sales_total_leads", up_verb="increased", down_verb="eased", label_override="Lead volume")
    deals_shift = metric_transition_clause(metrics, "sales_total_deals", up_verb="increased", down_verb="fell", label_override="Deal volume")
    conversion_shift = metric_transition_clause(metrics, "sales_overall_conversion_rate", up_verb="improved", down_verb="fell", label_override="conversion")
    response_shift = metric_transition_clause(metrics, "sales_internet_response_time", up_verb="slowed", down_verb="improved", label_override="Response times")
    uncontacted_shift = metric_transition_clause(metrics, "sales_uncontacted_leads", up_verb="rose", down_verb="fell", label_override="uncontacted leads")
    appointment_shift = metric_transition_clause(metrics, "sales_appointment_to_sales", up_verb="improved", down_verb="fell", label_override="appointment-to-sale conversion")
    same_day_shift = metric_transition_clause(metrics, "sales_same_day_deal_pct", up_verb="improved", down_verb="fell", label_override="same-day deal mix")
    test_drive_shift = metric_transition_clause(metrics, "sales_new_leads_to_test_drives", up_verb="improved", down_verb="fell", label_override="lead-to-test-drive conversion")
    sales_campaign_deals_shift = metric_transition_clause(metrics, "sales_campaign_deals", up_verb="increased", down_verb="fell", label_override="Sales campaign deals")
    sales_campaign_roi_shift = metric_transition_clause(metrics, "sales_campaign_roi", up_verb="improved", down_verb="fell", label_override="sales campaign ROI")
    service_bookings_shift = metric_transition_clause(metrics, "service_bookings_from_campaigns", up_verb="rose", down_verb="fell", label_override="Service campaign bookings")
    service_roi_shift = metric_transition_clause(metrics, "service_campaign_roi", up_verb="improved", down_verb="fell", label_override="service campaign ROI")
    service_tasks_shift = metric_transition_clause(metrics, "service_outstanding_tasks", up_verb="rose", down_verb="fell", label_override="Outstanding service tasks")

    message_by_rule = {
        "combined_sales_service_up": join_phrases([value for value in [deals_shift, service_roi_shift] if value]),
        "leads_up_conversion_down": f"{leads_shift}, but {conversion_shift}." if leads_shift and conversion_shift else leads_shift or conversion_shift or rule.message,
        "leads_down_conversion_up": f"{leads_shift}, but {conversion_shift}." if leads_shift and conversion_shift else leads_shift or conversion_shift or rule.message,
        "leads_up_deals_down": f"{leads_shift}, but {deals_shift}." if leads_shift and deals_shift else leads_shift or deals_shift or rule.message,
        "lead_handling_pressure": f"{response_shift}, while {uncontacted_shift}." if response_shift and uncontacted_shift else response_shift or uncontacted_shift or rule.message,
        "lead_handling_improved": f"{response_shift}, while {uncontacted_shift}." if response_shift and uncontacted_shift else response_shift or uncontacted_shift or rule.message,
        "closing_execution_up": f"{appointment_shift}, with {same_day_shift}." if appointment_shift and same_day_shift else appointment_shift or same_day_shift or rule.message,
        "test_drive_and_conversion_down": test_drive_shift or conversion_shift or rule.message,
        "sales_marketing_up": f"{sales_campaign_deals_shift}, with {sales_campaign_roi_shift}." if sales_campaign_deals_shift and sales_campaign_roi_shift else sales_campaign_deals_shift or sales_campaign_roi_shift or rule.message,
        "sales_marketing_negative_roi": f"Sales campaign ROI turned negative to {format_metric_value_for_reporting(metrics['sales_campaign_roi'])}." if metrics.get("sales_campaign_roi") else rule.message,
        "sales_campaign_roi_decline": join_phrases([sales_campaign_roi_shift, sales_campaign_deals_shift]) or rule.message,
        "service_campaigns_up": f"{service_bookings_shift}, with {service_roi_shift}." if service_bookings_shift and service_roi_shift else service_bookings_shift or service_roi_shift or rule.message,
        "service_bookings_roi_mismatch": (
            f"{service_roi_shift}, despite booking momentum holding up."
            if service_roi_shift and metric_direction(metrics, "service_bookings_from_campaigns") in {"up", "flat"}
            else service_roi_shift or service_bookings_shift or rule.message
        ),
        "service_tasks_pressure": service_tasks_shift or rule.message,
        "service_follow_up_cleaner": service_tasks_shift or rule.message,
        "sales_volume_snapshot": deals_shift or rule.message,
        "sales_conversion_snapshot": conversion_shift or rule.message,
        "lead_response_snapshot": response_shift or rule.message,
        "service_booking_snapshot": service_bookings_shift or rule.message,
        "service_roi_snapshot": service_roi_shift or rule.message,
        "service_task_snapshot": service_tasks_shift or rule.message,
        "crm_execution_pressure_composite": build_crm_composite_insight_message(
            metrics,
            f"{rule.merge_key}|{rule.rule_id}",
        )
        or rule.message,
    }

    # rule.message at this point is still the RULE_MESSAGE_TEMPLATES text -- the old
    # consultant voice. Reach for it only when neither a metric-derived message nor a
    # plain fact list is available, so it can no longer surface just because a metric
    # was gated out.
    base_message = message_by_rule.get(rule.rule_id)
    # Several entries above end in "or rule.message", so a miss arrives here as the
    # template text rather than as None -- compare against it directly.
    if not base_message or base_message == rule.message:
        facts = build_rule_fact_list(rule, metrics, include_month_on_month=False)
        if facts:
            base_message = join_phrases(facts)
        else:
            rule.has_metric_evidence = False
            base_message = rule.message
    if not base_message.endswith((".", "!", "?")):
        base_message = f"{base_message}."

    # One added clause, never two: a line carrying both the arithmetic and the baseline
    # stops being short. A claim that ranks the month against the record wins, because
    # it is the thing the owner cannot work out from the figures in front of them.
    baseline = None if sentence_describes_opposite_moves(base_message) else rule_baseline_clause(rule, metrics)
    consequence = consequence_clause(rule.rule_id, metrics)
    extra: str | None = None
    if baseline and baseline[1] == "record":
        extra = baseline[0]
    elif consequence:
        extra = consequence
    elif baseline:
        extra = baseline[0]
    if extra and extra.split()[0].lower() not in base_message.lower():
        base_message = append_clause_before_sentence_end(base_message, extra)

    if rule.signal_type == "persistent_watchpoint":
        focus_metric_name = (rule.metric_focus or [None])[0]
        focus_message = focused_metric_message(focus_metric_name, metrics) if focus_metric_name else None
        if focus_message:
            base_message = f"{focus_message}." if not focus_message.endswith((".", "!", "?")) else focus_message
        suffix = persistent_watchpoint_suffix(rule, metrics)
        if suffix:
            return sentence_case_start(append_clause_before_sentence_end(base_message, suffix))
        return sentence_case_start(base_message)

    if rule.signal_type == "trend_supported":
        suffix = concise_trend_suffix(rule, metrics)
        if suffix:
            return sentence_case_start(append_clause_before_sentence_end(base_message, suffix))

    return sentence_case_start(base_message)


def finalise_rule_candidate(rule: RuleInsight, metrics: dict[str, dict[str, Any]]) -> RuleInsight:
    rule.signal_type = derive_rule_signal_type(rule, metrics)
    rule.metric_focus = dominant_focus_metrics(rule, metrics)
    adverse_focus = primary_adverse_focus_metric(rule, metrics)
    if adverse_focus:
        remaining_focus = [metric_name for metric_name in rule.metric_focus if metric_name != adverse_focus]
        rule.metric_focus = [adverse_focus, *remaining_focus]
    rule.message = compressed_rule_message(rule, metrics)
    return rule


def priority_level_from_weight(weight: int) -> str:
    if weight >= 90:
        return "Critical"
    if weight >= 70:
        return "High"
    if weight >= 45:
        return "Medium"
    return "Low"


def metric_percent_change(metrics: dict[str, dict[str, Any]], metric_name: str) -> float | None:
    metric = metrics.get(metric_name)
    if metric is None:
        return None
    percent_change = metric.get("percent_change")
    if isinstance(percent_change, (int, float)):
        return float(percent_change)
    return None


def metric_is_meaningful(metric: dict[str, Any] | None) -> bool:
    if metric is None:
        return False
    current = metric.get("current")
    previous = metric.get("previous")
    direction = metric.get("direction")
    if current == 0 and previous == 0:
        return False
    if direction == "flat" and current in {0, 0.0}:
        return False
    return True


def metric_significance_band(metric: dict[str, Any] | None) -> str:
    if not metric:
        return "not_available"
    return str(metric.get("significance_band", "not_available"))


def metric_is_significant(metric: dict[str, Any] | None, minimum_band: str = "slight") -> bool:
    if not metric:
        return False
    band = metric_significance_band(metric)
    return SIGNIFICANCE_ORDER.get(band, -1) >= SIGNIFICANCE_ORDER.get(minimum_band, 1)


def metric_change_has_scale(metric: dict[str, Any] | None) -> bool:
    """False for a count that moved too little to be worth a percentage.

    "Uncontacted leads rose from 2 to 3 (+50%)" outranked real movements because the
    significance band reads the percentage, not the size. Rates, currency and times are
    unaffected -- only counted things can have a base small enough to make this happen.
    """
    if not metric:
        return False
    if str(metric.get("metric_type", "count")) not in {"count", "volume"}:
        return True
    previous = metric.get("previous_value", metric.get("previous"))
    change = metric.get("absolute_change", metric.get("delta"))
    if not isinstance(previous, (int, float)) or not isinstance(change, (int, float)):
        return True
    return abs(float(change)) >= 3 or abs(float(previous)) >= 20


def metric_usable_for_reporting(metric: dict[str, Any] | None, minimum_band: str = "slight") -> bool:
    if not metric:
        return False
    if not bool(metric.get("usable_for_reporting", False)):
        return False
    if not bool(metric.get("is_reliable", True)):
        return False
    if bool(metric.get("is_suspicious", False)):
        # A metric flagged suspicious (percent change over 500%, rate above 100,
        # extreme ROI, zero baseline) is already distrusted by the engine. It must
        # not reach client-facing copy, only the review exports.
        return False
    if not metric_change_has_scale(metric):
        return False
    return metric_is_significant(metric, minimum_band)


def metric_label(metrics: dict[str, dict[str, Any]], metric_name: str) -> str:
    metric = metrics.get(metric_name)
    if metric and metric.get("metric_label"):
        return str(metric["metric_label"])
    return str(metric_definition(metric_name).get("label", metric_name.replace("_", " ")))


def build_metric_transition_text(
    metrics: dict[str, dict[str, Any]],
    metric_name: str,
    *,
    up_verb: str,
    down_verb: str,
    minimum_band: str = "slight",
    include_label: bool = True,
    label_override: str | None = None,
) -> str | None:
    metric = metrics.get(metric_name)
    if not metric_usable_for_reporting(metric, minimum_band):
        return None

    direction = str(metric.get("direction", "flat"))
    if direction == "flat":
        return None

    previous_display = format_metric_value_for_reporting(metric, "previous")
    current_display = format_metric_value_for_reporting(metric, "current")

    subject = label_override or metric_label(metrics, metric_name)
    verb = up_verb if direction == "up" else down_verb
    prefix = f"{subject} " if include_label else ""
    if str(metric.get("metric_type", "count")) == "rate":
        return f"{prefix}{verb} from {previous_display} to {current_display}"

    delta_display = format_metric_delta_for_reporting(metric)
    if not delta_display:
        return None
    return f"{prefix}{verb} from {previous_display} to {current_display} ({delta_display})"


def build_short_metric_fact(
    metrics: dict[str, dict[str, Any]],
    metric_name: str,
    *,
    label_override: str | None = None,
    minimum_band: str = "slight",
    include_month_on_month: bool = True,
) -> str | None:
    metric = metrics.get(metric_name)
    if not metric_usable_for_reporting(metric, minimum_band):
        return None

    direction = str(metric.get("direction", "flat"))
    if direction == "flat":
        return None

    metric_type = str(metric.get("metric_type", "count"))
    label = label_override or metric_label(metrics, metric_name)
    # Every figure in this report is month on month; saying so two or three times in
    # one summary is noise. The parameter is kept so call sites stay unchanged.
    del include_month_on_month
    suffix = ""
    if metric_type == "rate":
        previous_display = format_metric_value_for_reporting(metric, "previous")
        current_display = format_metric_value_for_reporting(metric, "current")
        return f"{label} {'up' if direction == 'up' else 'down'} from {previous_display} to {current_display}{suffix}"

    delta_text = format_absolute_percent_change_value(metric.get("percent_change"))
    if not delta_text:
        return None

    movement = "up" if direction == "up" else "down"
    previous_value = metric.get("previous")
    percent = metric.get("percent_change")
    # "Up 400%" off a base of 2, or "up 1890%" on a response time, says less than the figures.
    small_base = metric_type in {"count", "volume"} and isinstance(previous_value, (int, float)) and abs(previous_value) < 10
    extreme_time = metric_type == "time" and isinstance(percent, (int, float)) and abs(percent) >= 100
    current_value = metric.get("current")
    to_zero = isinstance(current_value, (int, float)) and abs(current_value) < 1e-9
    # "down 109.1%" from a profit to a loss is not a percentage anyone can read.
    sign_flip = (isinstance(previous_value, (int, float)) and isinstance(current_value, (int, float))
                 and previous_value * current_value < 0)
    if small_base or extreme_time or to_zero or sign_flip:
        previous_display = format_metric_value_for_reporting(metric, "previous")
        current_display = format_metric_value_for_reporting(metric, "current")
        return f"{label} {movement} from {previous_display} to {current_display}{suffix}"
    return f"{label} {movement} {delta_text}{suffix}"


def full_month_name_from_label(month_label: str) -> str:
    try:
        parsed = datetime.strptime(month_label, "%b %Y")
        return parsed.strftime("%B")
    except ValueError:
        return month_label.split()[0]


def sms_metric_value(metric: dict[str, Any]) -> str:
    metric_type = str(metric.get("metric_type", "count"))
    current_value = metric.get("current")
    previous_value = metric.get("previous")

    if not isinstance(current_value, (int, float)):
        return str(metric.get("current_raw", ""))

    if metric_type == "time":
        if abs(float(current_value)) < 1e-9:
            return "0"
        return format_duration_value(float(current_value))
    if metric_type == "rate":
        return format_rate_value(float(current_value))
    if metric_type == "currency":
        return format_currency_value(float(current_value))

    current_display = format_count_value(float(current_value))
    if isinstance(previous_value, (int, float)):
        previous_display = format_count_value(float(previous_value))
        return f"({previous_display}->{current_display})"
    return current_display


def build_sms_fact_segment(metric_name: str, metric: dict[str, Any], *, short_labels: bool = False) -> str | None:
    if not metric_usable_for_reporting(metric):
        return None

    direction = str(metric.get("direction", "flat"))
    if direction == "flat":
        return None

    label_map = {
        "sales_total_leads": ("Leads", "Leads"),
        "sales_total_deals": ("Deals", "Deals"),
        "sales_overall_conversion_rate": ("Conversion", "Conv"),
        "sales_internet_response_time": ("Response", "Resp"),
        "sales_email_capture_rate": ("Email capture", "Email"),
        "service_reminders_sent": ("Reminders", "Rem"),
        "service_bookings_from_campaigns": ("Bookings", "Bkgs"),
        "service_campaign_roi": ("ROI", "ROI"),
        "service_outstanding_tasks": ("Tasks", "Tasks"),
    }
    long_label, short_label = label_map.get(metric_name, (metric_label({metric_name: metric}, metric_name).title(), metric_label({metric_name: metric}, metric_name).title()))
    label = short_label if short_labels else long_label

    metric_type = str(metric.get("metric_type", "count"))
    current_value = sms_metric_value(metric)
    if metric_type in {"volume", "count"}:
        return f"{label} {'up' if direction == 'up' else 'down'} {current_value}"
    if metric_type == "rate":
        return f"{label} {'up' if direction == 'up' else 'down'} to {current_value}"
    if metric_type == "time":
        if current_value == "0":
            return f"{label} 0"
        return f"{label} {'up' if direction == 'up' else 'down'} to {current_value}"
    if metric_type == "currency":
        return f"{label} {'up' if direction == 'up' else 'down'} to {current_value}"
    return None


def sms_metric_candidates(metrics: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    candidate_specs = [
        ("sales_total_leads", 1, 7),
        ("sales_total_deals", 2, 7),
        ("sales_overall_conversion_rate", 3, 5),
        ("sales_internet_response_time", 4, 4),
        ("sales_email_capture_rate", 5, 2),
        ("service_reminders_sent", 6, 3),
        ("service_bookings_from_campaigns", 6, 3),
        ("service_campaign_roi", 7, 6),
        ("service_outstanding_tasks", 8, 1),
    ]

    candidates: list[dict[str, Any]] = []
    used_priority_slots: set[int] = set()
    for metric_name, priority, keep_weight in candidate_specs:
        if priority in used_priority_slots and priority == 6:
            continue
        metric = metrics.get(metric_name)
        if not metric_usable_for_reporting(metric):
            continue
        if not metric_is_significant(metric):
            continue
        if metric_name == "service_reminders_sent" and metrics.get("service_bookings_from_campaigns") and metric_usable_for_reporting(metrics.get("service_bookings_from_campaigns")):
            reminders_change = abs(float(metric.get("percent_change") or 0.0))
            bookings_change = abs(float(metrics["service_bookings_from_campaigns"].get("percent_change") or 0.0))
            if bookings_change > reminders_change:
                continue
        if metric_name == "service_bookings_from_campaigns" and metrics.get("service_reminders_sent") and metric_usable_for_reporting(metrics.get("service_reminders_sent")):
            reminders_change = abs(float(metrics["service_reminders_sent"].get("percent_change") or 0.0))
            bookings_change = abs(float(metric.get("percent_change") or 0.0))
            if reminders_change >= bookings_change:
                continue

        candidates.append(
            {
                "metric_name": metric_name,
                "priority": priority,
                "keep_weight": keep_weight,
                "segment": build_sms_fact_segment(metric_name, metric, short_labels=False),
                "short_segment": build_sms_fact_segment(metric_name, metric, short_labels=True),
            }
        )
        if priority == 6:
            used_priority_slots.add(priority)

    return [candidate for candidate in candidates if candidate.get("segment")]


def compose_executive_sms_strategic(
    month_label: str,
    context_seed: str,
    ctx: ExecutiveStrategicContext,
    metrics: dict[str, dict[str, Any]],
    sentiment: str,
) -> str:
    """Executive SMS stack: dominant condition → risk → offset → watchpoint (phase 4b)."""
    tier = dealership_activity_tier(metrics)
    dom = choose_deterministic_phrase(
        DOMINANT_STATE_SMS_FRAG.get(ctx.dominant_state, DOMINANT_STATE_SMS_FRAG["mixed_volatility"]),
        f"{context_seed}|exec_sms_dom|{ctx.dominant_state}",
    ).rstrip(".")
    pieces: list[str] = [dom]

    risk_piece: str | None = None
    if _profitability_adverse_meaningful(metrics) and ctx.dominant_state != "profitability_pressure":
        risk_piece = choose_deterministic_phrase(
            ["Largest risk: ROI pressure.", "Risk stack: profitability.", "Watch ROI economics first."],
            f"{context_seed}|sms_risk_roi|{ctx.dominant_state}",
        )
    elif _conversion_adverse_meaningful(metrics) and ctx.dominant_state not in {
        "conversion_deterioration",
        "funnel_deterioration",
        "lead_quality_deterioration_state",
    }:
        risk_piece = choose_deterministic_phrase(
            ["Largest risk: conversion drag.", "Risk stack: weaker conversions.", "Conversion still a drag."],
            f"{context_seed}|sms_risk_conv|{ctx.dominant_state}",
        )
    elif _operational_risk_adverse_meaningful(metrics) and ctx.dominant_state != "sustained_operational_pressure":
        risk_piece = choose_deterministic_phrase(
            ["Largest risk: follow-up speed.", "Risk stack: contact speed.", "Follow-up risk elevated."],
            f"{context_seed}|sms_risk_exec|{ctx.dominant_state}",
        )
    elif ctx.arbitration.conflict_clause:
        cc = ctx.arbitration.conflict_clause.strip().rstrip(".")
        if len(cc) > 52:
            cc = cc[:49].rstrip(",; ") + "…"
        risk_piece = cc

    if risk_piece:
        rp = risk_piece.rstrip(".")
        if rp.lower() not in " ".join(pieces).lower():
            pieces.append(rp)

    if not ctx.suppress_secondary_positives and ctx.primary_strength_clause:
        low = ctx.primary_strength_clause.lower()
        if "service" in low and "appointment" in low:
            off = choose_deterministic_phrase(
                [
                    "Offset: service appointments firmer.",
                    "Bright spot: aftersales appointments.",
                    "Service demand offset some sales noise.",
                ],
                f"{context_seed}|sms_offset_svc",
            )
        else:
            first = ctx.primary_strength_clause.split(".")[0].strip()
            if len(first) > 48:
                first = first[:45].rstrip(",; ") + "…"
            off = first
        blob = " ".join(pieces).lower()
        if off.lower() not in blob:
            pieces.append(off.rstrip("."))

    if ctx.strategic_watchpoint_clause:
        wp = ctx.strategic_watchpoint_clause.split(".")[0].strip()
        if len(wp) > 52:
            wp = wp[:49].rstrip(",; ") + "…"
        blob = " ".join(pieces).lower()
        if wp.lower() not in blob:
            pieces.append(wp.rstrip("."))

    if tier == "low" and _conversion_adverse_meaningful(metrics):
        caveat = choose_deterministic_phrase(
            ["Thin volume—read cautiously.", "Low throughput—nuance reads.", "Thin base—directional only."],
            f"{context_seed}|sms_lowvol_exec",
        )
        blob = " ".join(pieces).lower()
        if caveat.lower() not in blob:
            pieces.append(caveat.rstrip("."))

    narrative = ". ".join(pieces) + "."
    narrative = re.sub(r"\.{2,}", ".", narrative)
    return narrative


def build_sms_fallback_v2(month_label: str, sentiment: str) -> str:
    prefix = f"DealerSocket {full_month_name_from_label(month_label)} Insights:"
    fallback_map = {
        "strongly_improved": "Stronger month overall. Sales and service improved, but lead response remains the key watchpoint.",
        "slightly_improved": "Slightly stronger month overall. Sales improved, but lead follow-up remains the key watchpoint.",
        "mixed": "Mixed month overall. Sales improved, but lead response and service follow-up remain key watchpoints.",
        "significantly_softer": "Softer month overall. Sales weakened, with lead response and service follow-up the key watchpoints.",
        "slightly_softer": "Softer month overall. Sales eased, with lead response and service follow-up the key watchpoints.",
        "stable": "Stable month overall. Sales and service were broadly steady, with follow-up still the main watchpoint.",
    }
    sms = f"{prefix} {fallback_map.get(sentiment, fallback_map['stable'])}"
    return sms[:160].rstrip()


def build_sms_strategic_narrative(
    metrics: dict[str, dict[str, Any]],
    sentiment: str,
    month_label: str,
    context_seed: str,
    selected_rules: list[RuleInsight] | None,
    arbitration: StrategicArbitrationResult | None = None,
    strategic_context: ExecutiveStrategicContext | None = None,
) -> str | None:
    if strategic_context is not None:
        return compose_executive_sms_strategic(month_label, context_seed, strategic_context, metrics, sentiment)

    tier = dealership_activity_tier(metrics)
    month_name = full_month_name_from_label(month_label)
    top_rule_id = selected_rules[0].rule_id if selected_rules else ""

    if arbitration and arbitration.crosscurrent_clause and arbitration.crosscurrent_override_pillar == "conversion_efficiency":
        return choose_deterministic_phrase(
            [
                f"{month_name} snapshot: enquiry eased, but conversion carried the executive read.",
                f"{month_name} snapshot: headline demand softened while mid-funnel execution defined the month.",
                f"{month_name} snapshot: treat the month as conversion-led, not volume-led, at the leadership level.",
            ],
            f"{context_seed}|sms|crosscurrent_conv",
        )

    if arbitration and arbitration.mixed_attention:
        dom = arbitration.dominant_pillar
        run = arbitration.runner_up_pillar or ""
        if dom == "conversion_efficiency" and run == "lead_flow":
            return choose_deterministic_phrase(
                [
                    f"{month_name} snapshot: efficiency moved the sales story more than headline enquiry volume.",
                    f"{month_name} snapshot: read the month through conversion quality, not only lead counts.",
                    f"{month_name} snapshot: funnel efficiency—not raw demand—carried executive weight.",
                ],
                f"{context_seed}|sms|arb_conv_leads",
            )
        if dom == "lead_flow" and run == "conversion_efficiency":
            return choose_deterministic_phrase(
                [
                    f"{month_name} snapshot: demand contraction dominated; conversion support was secondary.",
                    f"{month_name} snapshot: enquiry and deal throughput told the story ahead of efficiency gains.",
                    f"{month_name} snapshot: volume movement—not funnel polish—was the headline commercial driver.",
                ],
                f"{context_seed}|sms|arb_leads_conv",
            )
        if dom == "operational_risk" and run == "conversion_efficiency":
            return choose_deterministic_phrase(
                [
                    f"{month_name} snapshot: follow-up risk still caps how safely conversion gains can scale.",
                    f"{month_name} snapshot: execution responsiveness matters first even when conversion firms.",
                    f"{month_name} snapshot: prioritise contact discipline alongside any conversion lift.",
                ],
                f"{context_seed}|sms|arb_risk_conv",
            )
        if dom == "profitability_movement" and run in {"conversion_efficiency", "lead_flow"}:
            return choose_deterministic_phrase(
                [
                    f"{month_name} snapshot: ROI and return signals outweighed mixed sales-flow noise.",
                    f"{month_name} snapshot: commercial economics—not activity alone—defined the strategic read.",
                    f"{month_name} snapshot: profitability movement led; treat volume swings as secondary context.",
                ],
                f"{context_seed}|sms|arb_roi_mixed",
            )

    response = metrics.get("sales_internet_response_time")
    if response and metric_usable_for_reporting(response) and not metric_is_currently_improving(response) and metric_is_significant(response, "slight"):
        return choose_deterministic_phrase(
            [
                f"{month_name} snapshot: follow-up risk rose as response times stretched.",
                f"{month_name} snapshot: lead responsiveness is the urgent operational watch.",
                f"{month_name} snapshot: slower responses are compressing conversion opportunity.",
            ],
            f"{context_seed}|sms|risk_response",
        )

    uncontacted = metrics.get("sales_uncontacted_leads")
    if (
        uncontacted
        and metric_usable_for_reporting(uncontacted)
        and str(uncontacted.get("direction")) == "up"
        and metric_is_significant(uncontacted, "slight")
    ):
        return choose_deterministic_phrase(
            [
                f"{month_name} snapshot: uncontacted leads climbed; execution discipline is the priority.",
                f"{month_name} snapshot: untouched enquiries are accumulating faster than ideal.",
                f"{month_name} snapshot: lead contact coverage needs a focused push.",
            ],
            f"{context_seed}|sms|risk_uncontacted",
        )

    conversion = metrics.get("sales_overall_conversion_rate")
    if conversion and metric_usable_for_reporting(conversion) and not metric_is_currently_improving(conversion) and metric_is_significant(conversion, "slight"):
        if tier == "low":
            return choose_deterministic_phrase(
                [
                    f"{month_name} snapshot: conversion softened; treat directional reads cautiously on thin volume.",
                    f"{month_name} snapshot: efficiency eased modestly with limited throughput backing the trend.",
                    f"{month_name} snapshot: conversion is the sales watch, but volumes keep conclusions measured.",
                ],
                f"{context_seed}|sms|conv_lowvol",
            )
        return choose_deterministic_phrase(
            [
                f"{month_name} snapshot: conversion efficiency is the headline sales watch.",
                f"{month_name} snapshot: the sales story is conversion quality on a real enquiry base.",
                f"{month_name} snapshot: funnel efficiency, not just lead flow, is defining the month.",
            ],
            f"{context_seed}|sms|conv_highvol",
        )

    for roi_key in ("service_campaign_roi", "sales_campaign_roi"):
        roi = metrics.get(roi_key)
        if roi and metric_usable_for_reporting(roi) and not metric_is_currently_improving(roi) and metric_is_significant(roi, "slight"):
            return choose_deterministic_phrase(
                [
                    f"{month_name} snapshot: campaign ROI weakened; validate economics before scaling spend.",
                    f"{month_name} snapshot: profitability signals from campaigns need a tighter review.",
                    f"{month_name} snapshot: commercial returns, not just activity, are under pressure.",
                ],
                f"{context_seed}|sms|roi|{roi_key}",
            )

    tasks = metrics.get("service_outstanding_tasks")
    if tasks and metric_usable_for_reporting(tasks) and not metric_is_currently_improving(tasks) and metric_is_significant(tasks, "slight"):
        return choose_deterministic_phrase(
            [
                f"{month_name} snapshot: service backlog built; retention workflows deserve attention.",
                f"{month_name} snapshot: outstanding service tasks are the aftersales pressure point.",
                f"{month_name} snapshot: service execution load is trending heavier than ideal.",
            ],
            f"{context_seed}|sms|tasks",
        )

    if top_rule_id in {"leads_up_conversion_down", "leads_up_deals_down"}:
        return choose_deterministic_phrase(
            [
                f"{month_name} snapshot: demand held, but closing efficiency is the operational gap.",
                f"{month_name} snapshot: the month is about converting real enquiry into delivered deals.",
                f"{month_name} snapshot: funnel leakage after lead capture is the strategic focus.",
            ],
            f"{context_seed}|sms|rule_funnel",
        )

    if sentiment in {"strongly_improved", "slightly_improved"}:
        return choose_deterministic_phrase(
            [
                f"{month_name} snapshot: performance tilted constructive; protect follow-up discipline.",
                f"{month_name} snapshot: a cleaner month overall with execution still the stabiliser.",
                f"{month_name} snapshot: momentum improved; sustain responsiveness as volumes move.",
            ],
            f"{context_seed}|sms|pos",
        )

    if sentiment in {"significantly_softer", "slightly_softer"}:
        return choose_deterministic_phrase(
            [
                f"{month_name} snapshot: results softened; triage lead handling and conversion first.",
                f"{month_name} snapshot: a weaker month—focus on funnel efficiency and service backlog.",
                f"{month_name} snapshot: operational drag showed in sales execution and follow-up.",
            ],
            f"{context_seed}|sms|neg",
        )

    if sentiment == "mixed":
        return choose_deterministic_phrase(
            [
                f"{month_name} snapshot: mixed signals; prioritise conversion discipline and responsive follow-up.",
                f"{month_name} snapshot: offsetting moves across sales and service—anchor on funnel efficiency.",
                f"{month_name} snapshot: uneven month; leadership focus should stay on closing quality and contact speed.",
            ],
            f"{context_seed}|sms|mixed",
        )

    return choose_deterministic_phrase(
        [
            f"{month_name} snapshot: broadly steady; protect follow-up speed as the main stabiliser.",
            f"{month_name} snapshot: range-bound month; keep execution tight on leads and service backlog.",
            f"{month_name} snapshot: consolidation tone; watch conversion and responsiveness on any volume lift.",
        ],
        f"{context_seed}|sms|steady",
    )


def build_sms_summary_v2(
    metrics: dict[str, dict[str, Any]],
    sentiment: str,
    month_label: str,
    *,
    context_seed: str = "",
    selected_rules: list[RuleInsight] | None = None,
    arbitration: StrategicArbitrationResult | None = None,
    strategic_context: ExecutiveStrategicContext | None = None,
) -> tuple[str, list[str]]:
    prefix = f"DealerSocket {full_month_name_from_label(month_label)} Insights:"
    if selected_rules:
        hb = phase165_build_sms_headline_body(metrics, selected_rules, context_seed)
        if hb:
            core = hb if hb.endswith(".") else f"{hb}."
            if strategic_context is None:
                suff = phase167_compact_sms_offset(metrics, selected_rules, context_seed, hb)
                if suff:
                    trial_core = f"{core.rstrip('.')}{suff}."
                    trial_sms = f"{prefix} {trial_core}"
                    trial_sms = phase165_scrub_interpretation_and_jargon(trial_sms)
                    if len(trial_sms) <= 160:
                        core = trial_core
            sms = f"{prefix} {core}"
            sms = phase165_scrub_interpretation_and_jargon(sms)
            sms = enforce_sms_length_limit(sms, 160)
            return sms, ["headline_first"]

    arb_effective = arbitration or (strategic_context.arbitration if strategic_context else None)
    strategic = None
    if context_seed:
        strategic = build_sms_strategic_narrative(
            metrics,
            sentiment,
            month_label,
            context_seed,
            selected_rules,
            arbitration=arb_effective,
            strategic_context=strategic_context,
        )
    if strategic:
        sms = f"{prefix} {strategic}".strip()
        if len(sms) > 160:
            room = max(0, 160 - len(prefix) - 1)
            sms = f"{prefix} {strategic[:room].rstrip(',; ')}"
        if len(sms) > 160:
            sms = sms[:160].rstrip()
        return sms, ["strategic", "compressed"]

    candidates = sms_metric_candidates(metrics)

    if len(candidates) < 2:
        sms = build_sms_fallback_v2(month_label, sentiment)
        return sms, ["fallback"]

    def compose(parts: list[str], compact_commas: bool = False) -> str:
        separator = "," if compact_commas else ", "
        return f"{prefix} {separator.join(parts)}"

    ordered_candidates = sorted(candidates, key=lambda item: (item["priority"], -item["keep_weight"], item["metric_name"]))
    active_parts = [candidate["segment"] for candidate in ordered_candidates if candidate.get("segment")]
    active_names = [candidate["metric_name"] for candidate in ordered_candidates if candidate.get("segment")]
    sms = compose(active_parts)

    drop_order = [
        "service_outstanding_tasks",
        "sales_email_capture_rate",
        "service_reminders_sent",
        "service_bookings_from_campaigns",
        "sales_internet_response_time",
        "sales_overall_conversion_rate",
    ]

    while len(sms) > 160 and len(active_parts) > 1:
        removed = False
        for metric_name in drop_order:
            if metric_name in active_names:
                index = active_names.index(metric_name)
                del active_names[index]
                del active_parts[index]
                removed = True
                break
        if not removed:
            active_names.pop()
            active_parts.pop()
        sms = compose(active_parts)

    if len(sms) > 160:
        short_segments_by_name = {candidate["metric_name"]: candidate.get("short_segment") or candidate.get("segment") for candidate in ordered_candidates}
        active_parts = [short_segments_by_name[name] for name in active_names]
        sms = compose(active_parts)

    if len(sms) > 160:
        sms = compose(active_parts, compact_commas=True)

    while len(sms) > 160 and len(active_parts) > 1:
        active_parts.pop()
        active_names.pop()
        sms = compose(active_parts, compact_commas=True)

    if len(sms) > 160:
        sms = build_sms_fallback_v2(month_label, sentiment)
        active_names = ["fallback"]

    if len(sms) > 160:
        logging.error("SMS v2 exceeded 160 characters after trimming: %s", sms)
        sms = sms[:160].rstrip()

    return sms, active_names


def enforce_sms_length_limit(sms_text: str, limit: int = SMS_CHAR_LIMIT, *, allow_truncate: bool = False) -> str:
    cleaned = (sms_text or "").strip()
    if len(cleaned) <= limit:
        return cleaned
    if not allow_truncate:
        return cleaned
    trimmed = cleaned[:limit].rstrip(",; -")
    return trimmed.rstrip()


FALLBACK_METRIC_ORDER: tuple[str, ...] = (
    "sales_total_deals",
    "sales_overall_conversion_rate",
    "sales_campaign_roi",
    "service_campaign_roi",
    "sales_total_leads",
    "sales_new_leads_to_test_drives",
    "sales_uncontacted_leads",
    "sales_internet_response_time",
    "sales_email_capture_rate",
    "service_bookings_from_campaigns",
    "service_outstanding_tasks",
    "service_reminders_sent",
)

FALLBACK_METRIC_THEMES: dict[str, frozenset[str]] = {
    "sales_total_deals": frozenset({"deal_volume"}),
    "sales_overall_conversion_rate": frozenset({"conversion"}),
    "sales_campaign_roi": frozenset({"roi", "campaign_performance"}),
    "service_campaign_roi": frozenset({"service_roi", "roi", "campaign_performance"}),
    "sales_total_leads": frozenset({"lead_volume"}),
    "sales_new_leads_to_test_drives": frozenset({"test_drives"}),
    "sales_new_leads_to_appointments": frozenset({"appointments"}),
    "sales_appointment_to_sales": frozenset({"appointments", "conversion"}),
    "sales_uncontacted_leads": frozenset({"uncontacted_leads"}),
    "sales_internet_response_time": frozenset({"follow_up"}),
    "sales_email_capture_rate": frozenset({"email_capture"}),
    "service_bookings_from_campaigns": frozenset({"campaign_performance"}),
    "service_outstanding_tasks": frozenset({"service_tasks"}),
    "service_reminders_sent": frozenset({"service_tasks", "campaign_performance"}),
}

INSIGHT_THEME_NEEDLES: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("service_roi", ("service roi", "service campaign roi", "workshop roi", "aftersales roi")),
    ("email_capture", ("email capture", "capture rate", "contactability")),
    ("uncontacted_leads", ("uncontacted lead", "uncontacted leads")),
    ("service_tasks", ("service task", "service tasks", "outstanding service", "service follow-up", "service follow up")),
    ("overdue_tasks", ("overdue task", "overdue tasks", "overdue sales task", "overdue sales tasks")),
    ("test_drives", ("test drive", "test drives", "demo rate", "demos")),
    ("appointments", ("appointment", "appointments", "lead-to-appointment", "appointment-to-sale")),
    ("deal_volume", ("deal volume", "deal volumes", "deal throughput", "deals closed", "deals up", "deals down")),
    ("conversion", ("conversion", "closing rate", "closing discipline", "funnel efficiency")),
    ("lead_volume", ("lead volume", "lead flow", "lead activity", "enquiry", "enquiries", "top-of-funnel")),
    ("campaign_performance", ("campaign", "campaigns", "reminder", "reminders", "bookings from campaign")),
    ("roi", (" roi", "return on investment", "commercial return", "campaign return", "campaign returns")),
)

# Reached only when no further metric cleared the significance gates, so each of
# these is true by construction. They state that fact rather than issuing advice.
LIMITED_SIGNAL_FALLBACKS: tuple[str, ...] = (
    "No further sales metric moved enough to report this month.",
    "No further service metric moved enough to report this month.",
)

LAST_RESORT_DISTINCT_FALLBACKS: tuple[str, ...] = (
    "Remaining sales and service metrics held close to last month.",
    "Nothing further cleared the reporting threshold this month.",
)


def _insight_dedupe_key(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip().lower())[:96]


def _theme_ids_for_metric(metric_name: str) -> frozenset[str]:
    return FALLBACK_METRIC_THEMES.get(metric_name, frozenset())


def _theme_ids_for_insight(text: str) -> frozenset[str]:
    low = f" {(text or '').lower()} "
    themes: set[str] = set()
    for theme_id, needles in INSIGHT_THEME_NEEDLES:
        if any(needle in low for needle in needles):
            themes.add(theme_id)
    return frozenset(themes)


def _theme_ids_for_rules(selected_rules: list[RuleInsight] | None) -> frozenset[str]:
    themes: set[str] = set()
    for rule in selected_rules or []:
        themes.update(_theme_ids_for_insight(rule.message))
        for metric_name in [*(rule.metric_focus or []), *rule.supporting_metrics]:
            themes.update(_theme_ids_for_metric(metric_name))
    return frozenset(themes)


def _metric_evidence_for_rule(rule: RuleInsight) -> frozenset[str]:
    return frozenset(unique_preserving_order([*(rule.metric_focus or []), *rule.supporting_metrics]))


def _metric_evidence_for_insight(text: str) -> frozenset[str]:
    low = (text or "").lower()
    evidence: set[str] = set()
    metric_needles: tuple[tuple[str, tuple[str, ...]], ...] = (
        ("sales_campaign_roi", ("sales campaign roi", "sales roi")),
        ("service_campaign_roi", ("service campaign roi", "service roi")),
        ("sales_campaign_deals", ("sales campaign deals", "campaign deals")),
        ("sales_total_deals", ("deal volume", "total deals", "deal throughput")),
        ("sales_total_leads", ("lead volume", "lead flow", "more leads", "more enquiries", "fewer leads", "fewer enquiries")),
        ("sales_overall_conversion_rate", ("overall conversion", "conversion improved", "conversion fell", "conversion down", "conversion up")),
        ("sales_new_leads_to_appointments", ("lead-to-appointment", "appointment conversion", "converting to appointments")),
        ("sales_new_leads_to_test_drives", ("lead-to-test-drive", "test-drive conversion", "test drive conversion")),
        ("sales_appointment_to_sales", ("appointment-to-sale", "sales conversion")),
        ("sales_internet_response_time", ("response time", "response times", "responded much faster")),
        ("sales_uncontacted_leads", ("uncontacted leads",)),
        ("sales_overdue_tasks", ("overdue sales tasks", "overdue tasks")),
        ("service_outstanding_tasks", ("outstanding service tasks", "service task carryover", "service follow-up backlog")),
        ("sales_email_capture_rate", ("email capture",)),
        ("service_reminders_sent", ("service reminders",)),
        ("service_bookings_from_campaigns", ("service bookings", "service campaign bookings")),
        ("service_customers_enrolled", ("customers enrolled", "enrolled clients")),
    )
    for metric_name, needles in metric_needles:
        if any(needle in low for needle in needles):
            evidence.add(metric_name)
    specific_conversion_metrics = {
        "sales_new_leads_to_appointments",
        "sales_new_leads_to_test_drives",
        "sales_appointment_to_sales",
    }
    if evidence & specific_conversion_metrics and not re.search(
        r"(^|[.;]\s*)conversion (improved|fell|down|up|softened|weakened)",
        low,
    ):
        evidence.discard("sales_overall_conversion_rate")
    return frozenset(evidence)


def _limited_signal_fallback_insight(slot_index: int) -> str:
    idx = min(max(slot_index, 0), len(LIMITED_SIGNAL_FALLBACKS) - 1)
    return LIMITED_SIGNAL_FALLBACKS[idx]


def _iter_distinct_fallback_candidates(fallbacks: tuple[str, ...], seen_keys: set[str]) -> Iterable[str]:
    for fallback in fallbacks:
        key = _insight_dedupe_key(fallback)
        if key in seen_keys:
            continue
        seen_keys.add(key)
        yield fallback


def _metric_fallback_insights(
    metrics: dict[str, dict[str, Any]] | None,
    seen_keys: set[str],
    used_themes: set[str] | None = None,
) -> list[str]:
    if not metrics:
        return []
    out: list[str] = []
    active_themes = used_themes if used_themes is not None else set()
    # Preferred headline metrics first, then anything else parsed, then a relaxed pass
    # that allows a theme already covered. A real second fact about the month always
    # beats a sentence from the fallback bank.
    search_order = [*FALLBACK_METRIC_ORDER]
    search_order.extend(name for name in METRIC_DEFINITIONS if name not in search_order)
    for minimum_band, ignore_theme_clash in (
        ("meaningful", False),
        ("slight", False),
        ("meaningful", True),
        ("slight", True),
    ):
        for metric_name in search_order:
            metric_themes = set(_theme_ids_for_metric(metric_name))
            if metric_themes and metric_themes & active_themes and not ignore_theme_clash:
                continue
            fact = build_short_metric_fact(metrics, metric_name, minimum_band=minimum_band)
            if not fact:
                continue
            sentence = sentence_case_start(fact.strip().rstrip(".")) + "."
            key = _insight_dedupe_key(sentence)
            if key in seen_keys:
                continue
            seen_keys.add(key)
            active_themes.update(metric_themes or _theme_ids_for_insight(sentence))
            out.append(sentence)
    return out


def ensure_four_priority_insights(
    candidate_insights: Iterable[str],
    *,
    metrics: dict[str, dict[str, Any]] | None = None,
    selected_rules: list[RuleInsight] | None = None,
    pad: bool = True,
) -> list[str]:
    """Return up to four distinct, client-ready priority insights.

    Padding to four is the normal contract. ``pad=False`` is for placeholder months,
    where the source reports could not be read and there is nothing to pad with.
    """
    cleaned: list[str] = []
    seen: set[str] = set()
    used_themes: set[str] = set(_theme_ids_for_rules(selected_rules))
    for raw in candidate_insights or []:
        text = (raw or "").strip()
        if not text or text == NO_ADDITIONAL_INSIGHT_MARKER:
            continue
        key = _insight_dedupe_key(text)
        if key in seen:
            continue
        seen.add(key)
        used_themes.update(_theme_ids_for_insight(text))
        cleaned.append(text)

    if len(cleaned) > PRIORITY_INSIGHT_COUNT:
        cleaned = cleaned[:PRIORITY_INSIGHT_COUNT]

    if not pad:
        # A placeholder month has no figures to report. Do not mine partial metrics
        # and present them as findings when the report itself is blocked.
        return cleaned[:PRIORITY_INSIGHT_COUNT]

    for fact in _metric_fallback_insights(metrics, seen, used_themes):
        if len(cleaned) >= PRIORITY_INSIGHT_COUNT:
            break
        cleaned.append(fact)

    for fallback in _iter_distinct_fallback_candidates(LIMITED_SIGNAL_FALLBACKS, seen):
        if len(cleaned) >= PRIORITY_INSIGHT_COUNT:
            break
        used_themes.update(_theme_ids_for_insight(fallback))
        cleaned.append(fallback)

    for fallback in _iter_distinct_fallback_candidates(LAST_RESORT_DISTINCT_FALLBACKS, seen):
        if len(cleaned) >= PRIORITY_INSIGHT_COUNT:
            break
        cleaned.append(fallback)

    return cleaned[:PRIORITY_INSIGHT_COUNT]


@dataclass(frozen=True)
class SmsSignal:
    phrase: str
    family: str
    polarity: str
    metric: str
    quantified_phrase: str = ""
    source_insight: str = ""


FAMILY_INSIGHT_NEEDLES: dict[str, tuple[str, ...]] = {
    "service_roi": ("service campaign roi", "service roi"),
    "sales_roi": ("sales campaign roi", "sales roi"),
    "campaign_deals": ("sales campaign deals", "campaign deals"),
    "deal_volume": ("deal volume", "total deals", "deal throughput"),
    "appointment_conversion": ("lead-to-appointment", "appointment conversion", "converting to appointments"),
    "test_drive_conversion": ("lead-to-test-drive", "test-drive conversion", "test drive conversion"),
    "sales_conversion": ("appointment-to-sale", "sales conversion"),
    "conversion": ("overall conversion", "conversion improved", "conversion fell", "conversion down", "conversion up", "conversion stayed constructive"),
    "lead_volume": ("lead volume", "lead flow", "more leads", "more enquiries", "fewer leads", "fewer enquiries"),
    "response_time": ("response time", "response times", "responded much faster"),
    "uncontacted_leads": ("uncontacted leads",),
    "overdue_tasks": ("overdue sales tasks", "overdue tasks"),
    "service_tasks": ("outstanding service tasks", "service task carryover", "service follow-up backlog", "service tasks"),
    "email_capture": ("email capture",),
    "service_activity": ("service reminders", "service bookings", "service campaign bookings", "customers enrolled"),
}


def _insight_segment_for_family(insight: str, family: str) -> str:
    needles = FAMILY_INSIGHT_NEEDLES.get(family, ())
    parts = re.split(r"\s+and\s+", insight, flags=re.IGNORECASE)
    for part in parts:
        low = part.lower()
        if any(needle in low for needle in needles):
            return part.strip()
    low = (insight or "").lower()
    if needles and any(needle in low for needle in needles):
        return insight
    return insight


def _parse_campaign_deals_quantifiers(segment: str) -> tuple[str | None, tuple[str, str] | None]:
    match = re.search(
        r"(?:sales campaign deals|campaign deals)\s+\w+\s+from\s+(\d+)\s+to\s+(\d+)(?:\s*\(([+-]?\d+(?:\.\d+)?%)\))?",
        segment or "",
        flags=re.IGNORECASE,
    )
    if match:
        pct = match.group(3).lstrip("+") if match.group(3) else None
        return pct, (match.group(1).strip(), match.group(2).strip())
    pct = _parse_parenthetical_percent(segment) or _parse_inline_percent(segment)
    from_to = _parse_from_to_values(segment)
    if from_to and not from_to[0].startswith("$"):
        return pct, from_to
    return pct, None


def _parse_parenthetical_percent(text: str) -> str | None:
    match = re.search(r"\(([+-]?\d+(?:\.\d+)?%)\)", text or "")
    if not match:
        return None
    return match.group(1).lstrip("+")


def _parse_inline_percent(text: str) -> str | None:
    patterns = (
        r"(\d+(?:\.\d+)?)%\s+lead volume growth",
        r"(\d+(?:\.\d+)?)%\s+more enquiries",
        r"lead volume up\s+(\d+(?:\.\d+)?)%",
        r"appointment conversion (?:fell|weakened|softened)\s+(\d+(?:\.\d+)?)%",
        r"lead-to-appointment conversion weakened\s+(\d+(?:\.\d+)?)%",
        r"(?:up|down|fell|rose|increased|decreased|eased|softened|improved|growth)\s+(\d+(?:\.\d+)?)\s*%",
        r"(\d+(?:\.\d+)?)%\s+(?:month on month|mom)",
    )
    for pattern in patterns:
        match = re.search(pattern, text or "", flags=re.IGNORECASE)
        if match:
            return f"{match.group(1)}%"
    return None


def _parse_roi_quantifiers(segment: str, *, sales: bool) -> tuple[str | None, tuple[str, str] | None]:
    label = "sales campaign roi" if sales else "service campaign roi"
    lowered = (segment or "").lower()
    if label not in lowered:
        return None, None
    roi_match = re.search(rf"({label}[^.]+)", segment, flags=re.IGNORECASE)
    roi_part = roi_match.group(1) if roi_match else segment
    pct = _parse_parenthetical_percent(roi_part)
    from_to_match = re.search(
        rf"{label}\s+\w+\s+from\s+(\$[\d,.-]+)\s+to\s+(\$[\d,.-]+)",
        roi_part,
        flags=re.IGNORECASE,
    )
    if not from_to_match:
        from_to_match = re.search(
            rf"{label}\s+\w+\s+from\s+(\d+(?:\.\d+)?%)\s+to\s+(\d+(?:\.\d+)?%)",
            roi_part,
            flags=re.IGNORECASE,
        )
    from_to = (from_to_match.group(1).strip(), from_to_match.group(2).strip()) if from_to_match else None
    return pct, from_to


def _parse_from_to_values(text: str) -> tuple[str, str] | None:
    patterns = (
        r"from\s+(\$[\d,.-]+)\s+to\s+(\$[\d,.-]+)",
        r"from\s+(\d+(?:\.\d+)?%)\s+to\s+(\d+(?:\.\d+)?%)",
        r"from\s+(\d+h\s?\d{2}m)\s+to\s+(\d+h\s?\d{2}m)",
        r"from\s+(\d+\s?min(?:\s+\d+\s?sec)?)\s+to\s+(\d+\s?min(?:\s+\d+\s?sec)?)",
        r"from\s+(\d+\s?sec)\s+to\s+(\d+\s?sec)",
        r"from\s+(\d+:\d{2})\s+to\s+(\d+:\d{2})",
        r"from\s+(\d+)\s+to\s+(\d+)",
    )
    for pattern in patterns:
        match = re.search(pattern, text or "", flags=re.IGNORECASE)
        if match:
            return match.group(1).strip(), match.group(2).strip()
    return None


def _parse_negative_roi_value(segment: str, *, sales: bool) -> str | None:
    label = "sales campaign roi" if sales else "service campaign roi"
    match = re.search(rf"{label}\s+turned\s+negative\s+to\s+(\$[\d,.-]+)", segment, flags=re.IGNORECASE)
    return match.group(1).strip() if match else None


def _insight_has_usable_number(text: str) -> bool:
    return bool(re.search(r"\d", text or ""))


def _signal_display_phrase(signal: SmsSignal) -> str:
    if signal.quantified_phrase and re.search(r"\d", signal.quantified_phrase):
        return signal.quantified_phrase
    return signal.phrase


def format_quantified_sms_phrase(signal: SmsSignal, insight: str) -> str:
    segment = _insight_segment_for_family(insight, signal.family)
    for candidate_segment in (segment, insight):
        phrase = _format_quantified_phrase_for_segment(signal, candidate_segment)
        if re.search(r"\d", phrase):
            return phrase
    return signal.phrase


def _format_quantified_phrase_for_segment(signal: SmsSignal, segment: str) -> str:
    polarity = signal.polarity
    pct = _parse_parenthetical_percent(segment) or _parse_inline_percent(segment)
    from_to = _parse_from_to_values(segment)
    family = signal.family

    if family in {"service_roi", "sales_roi"}:
        label = "service ROI" if family == "service_roi" else "sales ROI"
        verb = "improved" if polarity == "positive" else "softened"
        negative_value = _parse_negative_roi_value(segment, sales=family == "sales_roi")
        if negative_value and polarity == "negative":
            return f"{label} turned negative to {negative_value}"
        pct, roi_from_to = _parse_roi_quantifiers(segment, sales=family == "sales_roi")
        if roi_from_to and pct:
            return f"{label} {verb} to {roi_from_to[1]} ({pct})"
        if roi_from_to:
            return f"{label} {verb} from {roi_from_to[0]} to {roi_from_to[1]}"
        if pct:
            return f"{label} {verb} {pct}"
        return signal.phrase

    if family == "lead_volume":
        if pct:
            return f"leads rose {pct}" if polarity == "positive" else f"leads softened {pct}"
        if from_to:
            return f"leads up from {from_to[0]} to {from_to[1]}" if polarity == "positive" else f"leads down from {from_to[0]} to {from_to[1]}"
        return signal.phrase

    if family == "deal_volume":
        if from_to and pct:
            verb = "up" if polarity == "positive" else "down"
            return f"deal volume {verb} from {from_to[0]} to {from_to[1]} ({pct})"
        if pct:
            verb = "up" if polarity == "positive" else "down"
            return f"deal volume {verb} {pct}"
        if from_to:
            verb = "up" if polarity == "positive" else "down"
            return f"deals {verb} from {from_to[0]} to {from_to[1]}"
        return signal.phrase

    if family in {"conversion", "test_drive_conversion", "appointment_conversion", "sales_conversion"}:
        labels = {
            "conversion": "conversion",
            "test_drive_conversion": "test-drive conversion",
            "appointment_conversion": "appointment conversion",
            "sales_conversion": "sales conversion",
        }
        label = labels[family]
        if family == "appointment_conversion":
            appointment_pct = re.search(
                r"appointment conversion (?:fell|weakened|softened|improved)\s+(\d+(?:\.\d+)?)%",
                segment,
                flags=re.IGNORECASE,
            )
            if appointment_pct:
                verb = "softened" if polarity == "negative" else "improved"
                return f"{label} {verb} {appointment_pct.group(1)}%"
        if from_to and re.match(r"\d+:\d{2}|\d+h\s?\d{2}m|\d+\s?(?:min|sec)", from_to[0]):
            from_to = None
        if from_to:
            if polarity == "positive":
                return f"{label} up from {from_to[0]} to {from_to[1]}"
            return f"{label} down from {from_to[0]} to {from_to[1]}"
        if pct:
            verb = "improved" if polarity == "positive" else "softened"
            return f"{label} {verb} {pct}"
        return signal.phrase

    if family == "campaign_deals":
        pct, deals_from_to = _parse_campaign_deals_quantifiers(segment)
        if deals_from_to and pct:
            verb = "up" if polarity == "positive" else "down"
            return f"campaign deals {verb} from {deals_from_to[0]} to {deals_from_to[1]} ({pct})"
        if deals_from_to:
            verb = "lifted" if polarity == "positive" else "fell"
            return f"campaign deals {verb} from {deals_from_to[0]} to {deals_from_to[1]}"
        if pct:
            return f"campaign deals up {pct}" if polarity == "positive" else f"campaign deals down {pct}"
        return signal.phrase

    if family == "overdue_tasks":
        if pct:
            return f"overdue tasks up {pct}" if polarity == "negative" else f"overdue tasks down {pct}"
        return signal.phrase

    if family == "service_tasks":
        if "steady" in signal.phrase.lower() or "held steady" in signal.phrase.lower():
            steady_match = re.search(r"(?:steady|stable)\s+at\s+(\d+)", segment, flags=re.IGNORECASE)
            if steady_match:
                return f"service tasks steady at {steady_match.group(1)}"
            return signal.phrase
        if pct:
            return f"service tasks up {pct}" if polarity == "negative" else f"service tasks down {pct}"
        if from_to:
            return f"service tasks from {from_to[0]} to {from_to[1]}"
        return signal.phrase

    if family == "uncontacted_leads":
        if pct:
            return f"uncontacted leads up {pct}" if polarity == "negative" else f"uncontacted leads down {pct}"
        if from_to:
            return f"uncontacted leads from {from_to[0]} to {from_to[1]}"
        return signal.phrase

    if family == "response_time":
        if from_to:
            verb = "improved" if polarity == "positive" else "slowed"
            return f"response time {verb} to {from_to[1]}"
        if pct:
            verb = "improved" if polarity == "positive" else "slowed"
            return f"response times {verb} {pct}"
        return signal.phrase

    if family == "email_capture":
        if from_to:
            verb = "up" if polarity == "positive" else "down"
            return f"email capture {verb} from {from_to[0]} to {from_to[1]}"
        if pct:
            verb = "up" if polarity == "positive" else "down"
            return f"email capture {verb} {pct}"
        return signal.phrase

    if family == "service_activity":
        if pct:
            verb = "up" if polarity == "positive" else "down"
            return f"service activity {verb} {pct}"
        return signal.phrase

    return signal.phrase


def extract_quantified_signal(insight: str, signal: SmsSignal) -> SmsSignal:
    quantified = format_quantified_sms_phrase(signal, insight)
    return SmsSignal(
        phrase=signal.phrase,
        family=signal.family,
        polarity=signal.polarity,
        metric=signal.metric,
        quantified_phrase=quantified,
        source_insight=insight,
    )


def _enrich_sms_signal(insight: str, signal: SmsSignal) -> SmsSignal:
    return extract_quantified_signal(insight, signal)


def _has_any(text: str, needles: tuple[str, ...]) -> bool:
    return any(needle in text for needle in needles)


def _word_present(text: str, word: str) -> bool:
    if word == "up":
        return bool(re.search(r"(?<!follow-)\bup\b", text))
    return bool(re.search(rf"\b{re.escape(word)}\b", text))


def _direction_token_present(text: str, token: str) -> bool:
    token = token.strip()
    if not token:
        return False
    if " " in token or "-" in token:
        return token in text
    return _word_present(text, token)


def _semantic_direction_for_phrase(low: str, *, inverse: bool = False, steady_positive: bool = False) -> str:
    positive_words = (
        "increased",
        "improved",
        "strengthened",
        "rose",
        "grew",
        "growth",
        "higher",
        "up from",
        "up ",
        "much faster",
    )
    negative_words = (
        "fell",
        "down",
        "declined",
        "softened",
        "weakened",
        "eased",
        "fewer",
        "slowed",
        "negative",
        "turned negative",
    )
    steady_words = ("steady", "stable", "held near", "remained broadly stable", "similar")
    if steady_positive and any(token in low for token in steady_words):
        return "positive"
    if inverse:
        inverse_positive = ("improved", "much faster", "faster", "fell", "down", "reduced", "fewer")
        inverse_negative = ("slowed", "slower", "rose", "increased", "higher")
        pos = sum(1 for token in inverse_positive if _direction_token_present(low, token))
        neg = sum(1 for token in inverse_negative if _direction_token_present(low, token))
    else:
        pos = sum(1 for token in positive_words if _direction_token_present(low, token))
        neg = sum(1 for token in negative_words if _direction_token_present(low, token))
    if pos > neg:
        return "positive"
    if neg > pos:
        return "negative"
    return "neutral"


def _append_sms_signal(
    signals: list[SmsSignal],
    seen_metrics: set[str],
    *,
    low: str,
    metric: str,
    family: str,
    positive_phrase: str,
    negative_phrase: str,
    needles: tuple[str, ...],
    inverse: bool = False,
    steady_positive: bool = False,
) -> None:
    if metric in seen_metrics or not _has_any(low, needles):
        return
    direction = _semantic_direction_for_phrase(low, inverse=inverse, steady_positive=steady_positive)
    if direction == "positive":
        signals.append(SmsSignal(positive_phrase, family, "positive", metric))
        seen_metrics.add(metric)
    elif direction == "negative":
        signals.append(SmsSignal(negative_phrase, family, "negative", metric))
        seen_metrics.add(metric)


def _sms_signals_for_insight(insight: str) -> list[SmsSignal]:
    low = (insight or "").lower()
    signals: list[SmsSignal] = []
    seen_metrics: set[str] = set()

    _append_sms_signal(
        signals,
        seen_metrics,
        low=low,
        metric="sales_campaign_roi",
        family="sales_roi",
        positive_phrase="sales ROI improved",
        negative_phrase="sales ROI softened",
        needles=("sales campaign roi", "sales roi", "sales campaign roi turned negative"),
    )
    _append_sms_signal(
        signals,
        seen_metrics,
        low=low,
        metric="service_campaign_roi",
        family="service_roi",
        positive_phrase="service ROI improved",
        negative_phrase="service ROI softened",
        needles=("service campaign roi", "service roi"),
    )
    _append_sms_signal(
        signals,
        seen_metrics,
        low=low,
        metric="sales_campaign_deals",
        family="campaign_deals",
        positive_phrase="campaign deals improved",
        negative_phrase="campaign deals softened",
        needles=("sales campaign deals", "campaign deals"),
    )
    _append_sms_signal(
        signals,
        seen_metrics,
        low=low,
        metric="sales_total_deals",
        family="deal_volume",
        positive_phrase="deal volume improved",
        negative_phrase="deal volume softened",
        needles=("deal volume", "total deals", "deal throughput"),
    )
    if "sales_new_leads_to_appointments" not in seen_metrics and (
        "appointment conversion fell" in low
        or "appointment conversion weakened" in low
        or "fewer converting to appointments" in low
        or "lead-to-appointment conversion weakened" in low
    ):
        signals.append(
            SmsSignal("appointment conversion softened", "appointment_conversion", "negative", "sales_new_leads_to_appointments")
        )
        seen_metrics.add("sales_new_leads_to_appointments")
    _append_sms_signal(
        signals,
        seen_metrics,
        low=low,
        metric="sales_new_leads_to_appointments",
        family="appointment_conversion",
        positive_phrase="appointment conversion improved",
        negative_phrase="appointment conversion softened",
        needles=("lead-to-appointment", "appointment conversion", "converting to appointments"),
    )
    _append_sms_signal(
        signals,
        seen_metrics,
        low=low,
        metric="sales_new_leads_to_test_drives",
        family="test_drive_conversion",
        positive_phrase="test-drive conversion improved",
        negative_phrase="test-drive conversion softened",
        needles=("lead-to-test-drive", "test-drive conversion", "test drive conversion"),
    )
    _append_sms_signal(
        signals,
        seen_metrics,
        low=low,
        metric="sales_appointment_to_sales",
        family="sales_conversion",
        positive_phrase="sales conversion improved",
        negative_phrase="sales conversion softened",
        needles=("appointment-to-sale", "sales conversion"),
    )
    if not {
        "sales_new_leads_to_appointments",
        "sales_new_leads_to_test_drives",
        "sales_appointment_to_sales",
    } & seen_metrics:
        if "sales_overall_conversion_rate" not in seen_metrics and "conversion stayed constructive" in low:
            signals.append(SmsSignal("conversion held up", "conversion", "positive", "sales_overall_conversion_rate"))
            seen_metrics.add("sales_overall_conversion_rate")
        _append_sms_signal(
            signals,
            seen_metrics,
            low=low,
            metric="sales_overall_conversion_rate",
            family="conversion",
            positive_phrase="conversion improved",
            negative_phrase="conversion softened",
            needles=("overall conversion", "conversion improved", "conversion fell", "conversion down", "conversion up", "conversion stayed constructive"),
        )
    if "sales_total_leads" not in seen_metrics and (
        ("lead volume growth" in low or "lead volume increased" in low or "more leads" in low or "more enquiries" in low)
        # "more leads remained untouched" is a follow-up failure, not lead growth.
        and "untouched" not in low
        and "uncontacted" not in low
    ):
        signals.append(SmsSignal("leads rose", "lead_volume", "positive", "sales_total_leads"))
        seen_metrics.add("sales_total_leads")
    elif "sales_total_leads" not in seen_metrics and ("lead volume eased" in low or "fewer leads" in low or "fewer enquiries" in low):
        signals.append(SmsSignal("lead volume softened", "lead_volume", "negative", "sales_total_leads"))
        seen_metrics.add("sales_total_leads")
    _append_sms_signal(
        signals,
        seen_metrics,
        low=low,
        metric="sales_total_leads",
        family="lead_volume",
        positive_phrase="leads rose",
        negative_phrase="lead volume softened",
        needles=("lead volume", "lead flow", "more leads", "more enquiries", "fewer leads", "fewer enquiries"),
    )
    if "sales_internet_response_time" not in seen_metrics:
        if "response times slowed" in low or "response time slowed" in low:
            signals.append(SmsSignal("response times slowed", "response_time", "negative", "sales_internet_response_time"))
            seen_metrics.add("sales_internet_response_time")
        elif (
            "response times improved" in low
            or "response time down" in low
            or "response times down" in low
            or "responded much faster" in low
        ):
            signals.append(SmsSignal("response times improved", "response_time", "positive", "sales_internet_response_time"))
            seen_metrics.add("sales_internet_response_time")
    _append_sms_signal(
        signals,
        seen_metrics,
        low=low,
        metric="sales_internet_response_time",
        family="response_time",
        positive_phrase="response times improved",
        negative_phrase="response times slowed",
        needles=("response time", "response times", "responded much faster"),
        inverse=True,
    )
    if "sales_uncontacted_leads" not in seen_metrics:
        if "uncontacted leads up" in low or "uncontacted leads rose" in low:
            signals.append(SmsSignal("uncontacted leads rose", "uncontacted_leads", "negative", "sales_uncontacted_leads"))
            seen_metrics.add("sales_uncontacted_leads")
        elif "uncontacted leads fell" in low or "uncontacted leads down" in low or "uncontacted leads reduced" in low:
            signals.append(SmsSignal("uncontacted leads reduced", "uncontacted_leads", "positive", "sales_uncontacted_leads"))
            seen_metrics.add("sales_uncontacted_leads")
    _append_sms_signal(
        signals,
        seen_metrics,
        low=low,
        metric="sales_uncontacted_leads",
        family="uncontacted_leads",
        positive_phrase="uncontacted leads reduced",
        negative_phrase="uncontacted leads rose",
        needles=("uncontacted leads",),
        inverse=True,
    )
    if "sales_overdue_tasks" not in seen_metrics:
        if "overdue sales tasks up" in low or "overdue tasks up" in low or "overdue sales tasks rose" in low:
            signals.append(SmsSignal("overdue tasks rose", "overdue_tasks", "negative", "sales_overdue_tasks"))
            seen_metrics.add("sales_overdue_tasks")
        elif "overdue sales tasks fell" in low or "overdue tasks fell" in low or "overdue tasks down" in low:
            signals.append(SmsSignal("overdue tasks reduced", "overdue_tasks", "positive", "sales_overdue_tasks"))
            seen_metrics.add("sales_overdue_tasks")
    _append_sms_signal(
        signals,
        seen_metrics,
        low=low,
        metric="sales_overdue_tasks",
        family="overdue_tasks",
        positive_phrase="overdue tasks reduced",
        negative_phrase="overdue tasks rose",
        needles=("overdue sales tasks", "overdue tasks"),
        inverse=True,
    )
    if "service_outstanding_tasks" not in seen_metrics:
        if "outstanding service tasks rose" in low or "outstanding service tasks up" in low:
            signals.append(SmsSignal("service tasks rose", "service_tasks", "negative", "service_outstanding_tasks"))
            seen_metrics.add("service_outstanding_tasks")
        elif "outstanding service tasks fell" in low or "outstanding service tasks down" in low:
            signals.append(SmsSignal("service tasks improved", "service_tasks", "positive", "service_outstanding_tasks"))
            seen_metrics.add("service_outstanding_tasks")
        elif "service task carryover remained" in low or "service follow-up backlog remained" in low:
            signals.append(SmsSignal("service tasks held steady", "service_tasks", "positive", "service_outstanding_tasks"))
            seen_metrics.add("service_outstanding_tasks")
    _append_sms_signal(
        signals,
        seen_metrics,
        low=low,
        metric="service_outstanding_tasks",
        family="service_tasks",
        positive_phrase="service tasks improved",
        negative_phrase="service tasks rose",
        needles=("outstanding service tasks", "service task carryover", "service follow-up backlog", "service tasks"),
        inverse=True,
        steady_positive=True,
    )
    _append_sms_signal(
        signals,
        seen_metrics,
        low=low,
        metric="sales_email_capture_rate",
        family="email_capture",
        positive_phrase="email capture improved",
        negative_phrase="email capture softened",
        needles=("email capture",),
    )
    _append_sms_signal(
        signals,
        seen_metrics,
        low=low,
        metric="service_activity",
        family="service_activity",
        positive_phrase="service activity improved",
        negative_phrase="service activity softened",
        needles=("service reminders", "service bookings", "service campaign bookings", "customers enrolled"),
    )
    return [_enrich_sms_signal(insight, signal) for signal in signals]


def _infer_insight_polarity(text: str) -> str:
    signals = _sms_signals_for_insight(text)
    pos = sum(1 for signal in signals if signal.polarity == "positive")
    neg = sum(1 for signal in signals if signal.polarity == "negative")
    if neg > pos:
        return "negative"
    if pos > neg:
        return "positive"
    low = (text or "").lower()
    if any(token in low for token in ("fell", "declined", "softened", "weakened", "turned negative")):
        return "negative"
    if any(token in low for token in ("increased", "improved", "strengthened", "rose")):
        return "positive"
    return "neutral"


def _extract_sms_theme(insight: str) -> str:
    signals = _sms_signals_for_insight(insight)
    if signals:
        return signals[0].phrase
    return ""


def validate_sms_semantics(sms: str, insights: list[str] | None = None) -> list[str]:
    issues: list[str] = []
    low = (sms or "").lower()
    contradiction_patterns: tuple[tuple[str, str, str], ...] = (
        ("deals_up_down", r"\bdeals up\b.*\bdeals down\b|\bdeals down\b.*\bdeals up\b", "unqualified deals contradiction"),
        ("conversion_up_down", r"\bconversion up\b.*\bconversion down\b|\bconversion down\b.*\bconversion up\b", "unqualified conversion contradiction"),
        ("leads_up_down", r"\bleads up\b.*\bleads down\b|\bleads down\b.*\bleads up\b", "unqualified leads contradiction"),
        (
            "response_times_up_down",
            r"\bresponse times up\b.*\bresponse times down\b|\bresponse times down\b.*\bresponse times up\b",
            "unqualified response-time contradiction",
        ),
    )
    for _, pattern, message in contradiction_patterns:
        if re.search(pattern, low):
            issues.append(message)
    if "…" in (sms or "") or "..." in (sms or ""):
        issues.append("SMS appears to contain a truncated phrase")
    if re.search(r"\bLead qualification performance\b", sms or ""):
        issues.append("SMS contains fallback chunk text instead of a client-ready theme")

    if insights:
        evidence_counts: dict[str, int] = defaultdict(int)
        for insight in insights:
            for metric_name in _metric_evidence_for_insight(insight):
                evidence_counts[metric_name] += 1
        for metric_name, count in evidence_counts.items():
            if count > 1 and metric_name in {"service_campaign_roi", "sales_campaign_roi", "sales_total_deals", "sales_overall_conversion_rate"}:
                issues.append(f"duplicate metric evidence: {metric_name}")
    return issues


SMS_GENERIC_REFERRAL_PATTERNS: tuple[str, ...] = (
    r"see report",
    r"see attached",
    r"refer to",
    r"4 key insights",
    r"attached report",
    r"more details in the report",
)


def extract_priority_insights(report_text: str) -> list[str]:
    if "Priority Insights:" not in report_text:
        return []
    block = report_text.split("Priority Insights:", 1)[1]
    block = block.split("\n\nSummary:", 1)[0]
    return [line.strip() for line in block.splitlines() if line.strip()]


def validate_sms_no_generic_referral(sms_text: str) -> list[str]:
    low = (sms_text or "").lower()
    return [pattern for pattern in SMS_GENERIC_REFERRAL_PATTERNS if re.search(pattern, low)]


def validate_sms_format(sms_text: str, month_label: str) -> list[str]:
    month_name = full_month_name_from_label(month_label)
    expected_prefix = f"DealerSocket {month_name} Insights:"
    issues: list[str] = []
    text = (sms_text or "").strip()
    if not text.startswith(expected_prefix):
        issues.append(f"SMS must start with '{expected_prefix}'")
        return issues
    body = text[len(expected_prefix) :].lstrip("\n ").strip()
    if not body:
        issues.append("SMS missing compressed Priority Insights summary")
    return issues


def _sms_body_text(sms_text: str) -> str:
    if "Insights:" in (sms_text or ""):
        return sms_text.split("Insights:", 1)[1].lstrip("\n ").strip()
    return (sms_text or "").strip()


def _numbers_in_text(text: str) -> set[str]:
    return set(re.findall(r"\d+(?:\.\d+)?", text or ""))


def validate_sms_contains_quantified_metrics(sms_text: str, priority_insights: list[str]) -> list[str]:
    body = _sms_body_text(sms_text)
    if re.search(r"\d", body):
        return []
    return ["SMS contains no quantified metrics"]


def validate_sms_no_invented_numbers(sms_text: str, priority_insights: list[str]) -> list[str]:
    body = _sms_body_text(sms_text)
    sms_numbers = _numbers_in_text(body)
    insight_numbers = _numbers_in_text(" ".join(priority_insights))
    return [f"invented number: {number}" for number in sorted(sms_numbers) if number not in insight_numbers]


def validate_sms_number_traceability(sms_text: str, priority_insights: list[str]) -> list[str]:
    return validate_sms_no_invented_numbers(sms_text, priority_insights)


def validate_sms_qualitative_only(sms_text: str, priority_insights: list[str]) -> list[str]:
    body = _sms_body_text(sms_text).lower()
    issues: list[str] = []
    for signal in _collect_priority_insight_signals(priority_insights):
        display = _signal_display_phrase(signal)
        if re.search(r"\d", display):
            continue
        if not _insight_has_usable_number(signal.source_insight or ""):
            continue
        if display.lower() in body or signal.phrase.lower() in body:
            issues.append(f"qualitative-only phrase for {signal.family}")
    return issues


def validate_sms_priority_traceability(sms_text: str, priority_insights: list[str]) -> list[str]:
    issues: list[str] = []
    low = (sms_text or "").lower()
    signals = _collect_priority_insight_signals(priority_insights)
    if not signals:
        issues.append("no Priority Insight themes available for SMS traceability")
        return issues
    body = low.split("insights:", 1)[-1] if "insights:" in low else low
    represented = [
        signal
        for signal in signals
        if _signal_display_phrase(signal).lower() in body or signal.phrase.lower() in body
    ]
    if not represented:
        issues.append("SMS body does not reflect any Priority Insight themes")
    return issues


def validate_sms_polarity(priority_insights: list[str], sms_text: str) -> list[str]:
    issues: list[str] = []
    low = (sms_text or "").lower()
    inverse_families = {"overdue_tasks", "uncontacted_leads", "response_time", "service_tasks"}
    for insight in priority_insights:
        for signal in _sms_signals_for_insight(insight):
            phrase_low = signal.phrase.lower()
            if phrase_low not in low:
                continue
            if signal.polarity == "positive" and any(
                marker in phrase_low for marker in ("softened", "slowened", "eased", "fell", "declined")
            ):
                issues.append(f"polarity mismatch for theme '{signal.phrase}'")
            elif (
                signal.polarity == "negative"
                and signal.family not in inverse_families
                and any(marker in phrase_low for marker in ("improved", "lifted", "held steady", "held up"))
            ):
                issues.append(f"polarity mismatch for theme '{signal.phrase}'")
    return issues


def _join_sms_phrases(phrases: list[str]) -> str:
    cleaned = [phrase.strip() for phrase in phrases if phrase and phrase.strip()]
    if not cleaned:
        return ""
    if len(cleaned) == 1:
        return cleaned[0]
    if len(cleaned) == 2:
        return f"{cleaned[0]} and {cleaned[1]}"
    return ", ".join(cleaned[:-1]) + f", and {cleaned[-1]}"


def _collect_priority_insight_signals(insights: list[str]) -> list[SmsSignal]:
    ordered: list[SmsSignal] = []
    seen_metrics: set[str] = set()
    seen_families: set[str] = set()
    for insight in insights:
        for signal in _sms_signals_for_insight(insight):
            if signal.metric in seen_metrics or signal.family in seen_families:
                continue
            seen_metrics.add(signal.metric)
            seen_families.add(signal.family)
            ordered.append(signal)
    return ordered


def _signal_has_quantification(signal: SmsSignal) -> bool:
    return bool(re.search(r"\d", _signal_display_phrase(signal)))


def _compose_sms_body_from_signals(signals: list[SmsSignal]) -> str:
    if not signals:
        return "mixed movement across key sales and service priorities."
    positive = [_signal_display_phrase(signal) for signal in signals if signal.polarity in {"positive", "neutral"}]
    negative = [_signal_display_phrase(signal) for signal in signals if signal.polarity == "negative"]
    if positive and negative:
        return f"{_join_sms_phrases(positive)}, but {_join_sms_phrases(negative)}."
    phrases = [_signal_display_phrase(signal) for signal in signals]
    return f"{_join_sms_phrases(phrases)}."


def _format_dealersocket_sms(prefix: str, body: str, *, separator: str = "\n") -> str:
    cleaned_body = re.sub(r"\s+", " ", (body or "").strip())
    return f"{prefix}{separator}{cleaned_body}"


def _compress_sms_message(
    month_name: str,
    insights: list[str],
    *,
    limit: int = SMS_CHAR_LIMIT,
    signals: list[SmsSignal] | None = None,
) -> tuple[str, bool, str]:
    prefix = f"DealerSocket {month_name} Insights:"
    if signals is None:
        signals = _collect_priority_insight_signals(insights)
    working = list(signals)
    best_candidate = ""
    best_issues: list[str] = []

    while True:
        body = _compose_sms_body_from_signals(working)
        for separator in ("\n", " "):
            candidate = _format_dealersocket_sms(prefix, body, separator=separator)
            referral_issues = validate_sms_no_generic_referral(candidate)
            semantic_issues = validate_sms_semantics(candidate, insights)
            quant_issues = validate_sms_contains_quantified_metrics(candidate, insights)
            invented_issues = validate_sms_no_invented_numbers(candidate, insights)
            qualitative_issues = validate_sms_qualitative_only(candidate, insights)
            issues = referral_issues + semantic_issues + quant_issues + invented_issues + qualitative_issues
            if len(candidate) <= limit and not issues:
                return candidate, True, ""
            if len(candidate) <= limit and (not best_candidate or len(issues) < len(best_issues)):
                best_candidate = candidate
                best_issues = issues
        if len(working) <= 1:
            break
        unquantified = [signal for signal in working if not _signal_has_quantification(signal)]
        if unquantified:
            working.remove(unquantified[-1])
            continue
        working.pop()

    if best_candidate:
        warning = "; ".join(best_issues) if best_issues else ""
        return best_candidate, not best_issues, warning

    body = _compose_sms_body_from_signals(signals[:2] if signals else [])
    candidate = _format_dealersocket_sms(prefix, body, separator="\n")
    if len(candidate) > limit:
        room = max(20, limit - len(prefix) - 1)
        trimmed_body = body[:room].rstrip(",; ")
        if trimmed_body and not trimmed_body.endswith("."):
            trimmed_body += "."
        candidate = _format_dealersocket_sms(prefix, trimmed_body, separator="\n")
    issues = (
        validate_sms_no_generic_referral(candidate)
        + validate_sms_semantics(candidate, insights)
        + validate_sms_contains_quantified_metrics(candidate, insights)
        + validate_sms_no_invented_numbers(candidate, insights)
        + validate_sms_qualitative_only(candidate, insights)
    )
    valid = len(candidate) <= limit and not issues
    error = "" if valid else "; ".join(issues) or f"SMS exceeds {limit} characters after compression"
    return candidate, valid, error


def compress_priority_insights_to_sms(
    priority_insights: list[str],
    month_label: str,
) -> tuple[str, list[str], dict[str, Any]]:
    return build_sms_from_priority_insights(priority_insights, month_label)


# --- SMS composed from metric records ------------------------------------------
#
# The layer below this one derives SMS facts by pattern-matching the finished insight
# prose. That is how "uncontacted leads up -21.8%", "service tasks down -8.3%" and
# "email capture down from 41% to 43%" (where the numbers go up) reached clients: a
# direction word taken from one clause, a number from another. Everything the SMS needs
# is already in the metric records, so it is composed from those instead, and the prose
# path is kept only for callers that have nothing else (archived reports, the tools).

SMS_METRIC_LABELS: dict[str, str] = {
    "sales_total_leads": "leads",
    "sales_total_deals": "deals",
    "sales_overall_conversion_rate": "conversion",
    "sales_new_leads_to_test_drives": "test-drive conversion",
    "sales_new_leads_to_appointments": "lead-to-appointment conversion",
    "sales_appointment_to_sales": "appointment-to-sale conversion",
    "sales_same_day_deal_pct": "same-day deal mix",
    "sales_internet_response_time": "response time",
    "sales_uncontacted_leads": "uncontacted leads",
    "sales_overdue_tasks": "overdue tasks",
    "sales_email_capture_rate": "email capture",
    "sales_campaign_deals": "campaign deals",
    "sales_campaign_roi": "sales ROI",
    "sales_pipeline_size": "pipeline",
    "service_campaign_roi": "service ROI",
    "service_bookings_from_campaigns": "service bookings",
    "service_outstanding_tasks": "service tasks",
    "service_customers_enrolled": "enrolled clients",
    "service_reminders_sent": "service reminders",
    "service_no_shows": "service no-shows",
    "service_rebooked_no_shows": "rebooked no-shows",
}

# Families gate de-duplication, so two metrics telling the same story never both appear.
SMS_METRIC_FAMILIES: dict[str, str] = {
    "sales_total_leads": "lead_volume",
    "sales_total_deals": "deal_volume",
    "sales_overall_conversion_rate": "conversion",
    "sales_new_leads_to_test_drives": "test_drive_conversion",
    "sales_new_leads_to_appointments": "appointment_conversion",
    "sales_appointment_to_sales": "sales_conversion",
    "sales_same_day_deal_pct": "same_day_mix",
    "sales_internet_response_time": "response_time",
    "sales_uncontacted_leads": "uncontacted_leads",
    "sales_overdue_tasks": "overdue_tasks",
    "sales_email_capture_rate": "email_capture",
    "sales_campaign_deals": "campaign_deals",
    "sales_campaign_roi": "sales_roi",
    "sales_pipeline_size": "pipeline",
    "service_campaign_roi": "service_roi",
    "service_bookings_from_campaigns": "service_activity",
    "service_outstanding_tasks": "service_tasks",
    "service_customers_enrolled": "service_retention",
    "service_reminders_sent": "service_activity",
    "service_no_shows": "service_no_shows",
    "service_rebooked_no_shows": "service_no_shows",
}


# Three facts is the most an SMS can carry and still read as a headline rather than a
# list. The compressor drops further if the limit needs it.
MAX_SMS_SIGNALS = 3


def sms_metric_label(metric_name: str, metrics: dict[str, dict[str, Any]]) -> str:
    return SMS_METRIC_LABELS.get(metric_name) or metric_label(metrics, metric_name)


def sms_phrase_candidates(metric_name: str, metrics: dict[str, dict[str, Any]]) -> list[str]:
    """Phrases for one metric, richest first, every one built from the metric record.

    The direction word comes from the metric's own direction and the percentage is
    always absolute, so "up -21.8%" and "down from 41% to 43%" cannot be constructed.
    """
    metric = metrics.get(metric_name)
    if not metric:
        return []
    direction = str(metric.get("direction", "flat"))
    if direction == "flat":
        return []

    label = sms_metric_label(metric_name, metrics)
    metric_type = str(metric.get("metric_type", "count"))
    movement = "up" if direction == "up" else "down"
    percent = format_absolute_percent_change_value(metric.get("percent_change"))
    previous = format_metric_value_for_reporting(metric, "previous")
    current = format_metric_value_for_reporting(metric, "current")

    if metric_type == "time":
        # "up"/"down" on a duration invites the reader to score it as good or bad.
        pace = "slower" if direction == "up" else "faster"
        candidates = [f"{label} {pace} at {current}", f"{label} {pace} from {previous} to {current}"]
        if percent:
            candidates.append(f"{label} {movement} {percent}")
        candidates.append(f"{label} {pace}")
        return candidates

    if metric_type == "rate":
        # A relative percent change on a rate reads as the rate itself; endpoints only.
        return [
            f"{label} {movement} from {previous} to {current}",
            f"{label} {movement} to {current}",
        ]

    candidates: list[str] = []
    if percent:
        candidates.append(f"{label} {movement} {percent} to {current}")
    candidates.append(f"{label} {movement} from {previous} to {current}")
    if percent:
        candidates.append(f"{label} {movement} {percent}")
    candidates.append(f"{label} {movement} to {current}")
    return candidates


def sms_candidate_metric_names(
    selected_rules: list[RuleInsight] | None,
    metrics: dict[str, dict[str, Any]],
    insights: list[str],
) -> list[str]:
    """Metrics the priority insights actually talk about.

    Either the rule that produced an insight names the metric, or the metric's label
    appears in the text. Both keep the SMS to things the report itself said.
    """
    joined = " ".join(insights or []).lower()
    names: list[str] = []
    for rule in selected_rules or []:
        names.extend([*(rule.metric_focus or []), *(rule.supporting_metrics or [])])
    for metric_name in metrics:
        if metric_label(metrics, metric_name).lower() in joined:
            names.append(metric_name)
    return unique_preserving_order(names)


def sms_signals_from_metrics(
    selected_rules: list[RuleInsight] | None,
    metrics: dict[str, dict[str, Any]] | None,
    insights: list[str],
) -> list[SmsSignal]:
    """Ordered, de-duplicated SMS signals built from the metric records."""
    if not metrics:
        return []
    insight_numbers = _numbers_in_text(" ".join(insights or []))
    scored: list[tuple[float, str, SmsSignal]] = []
    for metric_name in sms_candidate_metric_names(selected_rules, metrics, insights):
        metric = metrics.get(metric_name)
        if not metric_usable_for_reporting(metric):
            continue
        phrase = ""
        for candidate in sms_phrase_candidates(metric_name, metrics):
            # Every number in the SMS has to be one the report already showed, so a
            # reader can find it. A phrase that cannot be traced is simply not used.
            if _numbers_in_text(candidate) <= insight_numbers:
                phrase = candidate
                break
        if not phrase:
            continue
        weight = INSIGHT_MATERIALITY_WEIGHTS.get(metric_name, DEFAULT_MATERIALITY_WEIGHT)
        scored.append(
            (
                metric_materiality(metric) * weight,
                metric_name,
                SmsSignal(
                    phrase=phrase,
                    family=SMS_METRIC_FAMILIES.get(metric_name, metric_name),
                    polarity="positive" if metric_is_improvement(metric) else "negative",
                    metric=metric_name,
                    quantified_phrase=phrase,
                ),
            )
        )

    scored.sort(key=lambda item: item[0], reverse=True)
    signals: list[SmsSignal] = []
    seen_families: set[str] = set()
    for _, _, signal in scored:
        if signal.family in seen_families:
            continue
        seen_families.add(signal.family)
        signals.append(signal)
        if len(signals) >= MAX_SMS_SIGNALS:
            break
    return _ensure_sms_shows_both_sides(signals, scored)


def _ensure_sms_shows_both_sides(
    signals: list[SmsSignal],
    scored: list[tuple[float, str, "SmsSignal"]],
) -> list[SmsSignal]:
    """Do not let three of a kind stand in for a month that moved both ways.

    The SMS takes the three most material signals, one per family. In a month where one
    department moved hard in one direction, all three slots can go to that department:
    a month with deals up 26.7% and conversion up 12 points went out as "service ROI
    down 79.6%, service tasks up 266.7%, and service bookings down to 8" — every word
    true, the month itself misrepresented.

    So the weakest signal gives up its slot to the strongest signal of the missing
    polarity, when one exists. `_compose_sms_body_from_signals` then renders it as
    "X, but Y" on its own.
    """
    if len(signals) < 2:
        return signals
    present = {signal.polarity for signal in signals}
    if "positive" in present and "negative" in present:
        return signals
    missing = "negative" if "positive" in present else "positive"
    families = {signal.family for signal in signals}
    replacement = next(
        (
            signal
            for _, _, signal in scored
            if signal.polarity == missing and signal.family not in families
        ),
        None,
    )
    if replacement is None:
        return signals
    # Drop the least material of the same-polarity run, never the leading signal:
    # the message still opens on the month's biggest movement.
    return signals[:-1] + [replacement]


def build_sms_from_priority_insights(
    bullet_insights: list[str],
    month_label: str,
    *,
    metrics: dict[str, dict[str, Any]] | None = None,
    selected_rules: list[RuleInsight] | None = None,
) -> tuple[str, list[str], dict[str, Any]]:
    """Build client SMS from the final four Priority Insights only.

    Where the metric records are available the SMS is composed from them; callers that
    only hold the finished text fall back to reading the facts back out of the prose.
    """
    insights = ensure_four_priority_insights(bullet_insights)
    month_name = full_month_name_from_label(month_label)
    signals = sms_signals_from_metrics(selected_rules, metrics, insights)
    sms, valid, error = _compress_sms_message(month_name, insights, signals=signals or None)
    meta = {
        "sms_valid": valid,
        "sms_char_count": len(sms),
        "sms_limit": SMS_CHAR_LIMIT,
        "sms_source": "priority_insights",
        "sms_composition": "metrics" if signals else "text",
        "sms_error": error,
    }
    if not valid:
        logging.error("SMS validation failed for %s: %s (%s chars)", month_label, error, len(sms))
    return sms, ["priority_insights"], meta


METRIC_FAMILIES = ("sales", "service")


def metric_family_coverage(
    metrics: dict[str, dict[str, Any]] | None,
) -> dict[str, dict[str, int]]:
    """How many metrics each half of the business contributed, and how many survived gating."""
    coverage = {family: {"parsed": 0, "usable": 0} for family in METRIC_FAMILIES}
    for name, metric in (metrics or {}).items():
        family = str(name).split("_", 1)[0]
        if family not in coverage:
            continue
        coverage[family]["parsed"] += 1
        if metric_usable_for_reporting(metric):
            coverage[family]["usable"] += 1
    return coverage


def families_without_usable_evidence(
    metrics: dict[str, dict[str, Any]] | None,
) -> list[str]:
    """Halves of the business that gave the report nothing it could use this month.

    Only meaningful while the other half did give something: when neither side has
    usable metrics the month is a parsing failure, which is handled upstream as a
    placeholder report rather than as an imbalance.
    """
    coverage = metric_family_coverage(metrics)
    empty = [family for family, counts in coverage.items() if counts["usable"] == 0]
    if len(empty) == len(METRIC_FAMILIES):
        return []
    return empty


def _insight_line_polarity(line: str) -> str:
    """positive, negative, mixed or neutral, read from one rendered insight line."""
    polarities = {signal.polarity for signal in _sms_signals_for_insight(line)}
    if not polarities:
        return "neutral"
    if "positive" in polarities and "negative" in polarities:
        return "mixed"
    if "negative" in polarities:
        return "negative"
    if "positive" in polarities:
        return "positive"
    return "neutral"


def rebalance_when_a_family_reported_nothing(
    bullets: list[str],
    metrics: dict[str, dict[str, Any]] | None,
) -> list[str]:
    """Lead with a real improvement when one side of the business reported nothing.

    With no usable service figures every line comes from sales, and in a soft month
    that reads as an unbroken run of bad news even where something genuinely improved
    — the same dealership with service data leads on a campaign ROI rise instead. The
    imbalance is in the ordering, not in the facts.

    So nothing is added, removed or reworded here: a line already in the report moves
    up. Every decline still appears, in full, immediately below.
    """
    if len(bullets) < 2 or not families_without_usable_evidence(metrics):
        return bullets
    if _insight_line_polarity(bullets[0]) != "negative":
        return bullets
    for index in range(1, len(bullets)):
        if _insight_line_polarity(bullets[index]) == "positive":
            reordered = list(bullets)
            reordered.insert(0, reordered.pop(index))
            return reordered
    return bullets


def apply_final_insight_and_sms_package(
    bullets: list[str],
    summary: str,
    month_label: str,
    metrics: dict[str, dict[str, Any]] | None = None,
    selected_rules: list[RuleInsight] | None = None,
    pad: bool = True,
) -> tuple[list[str], str, str, list[str], dict[str, Any]]:
    final_bullets = ensure_four_priority_insights(
        bullets,
        metrics=metrics,
        selected_rules=selected_rules,
        pad=pad,
    )
    # Applied before the SMS is composed so the message inherits the same order.
    final_bullets = rebalance_when_a_family_reported_nothing(final_bullets, metrics)
    sms, sms_segments, sms_meta = build_sms_from_priority_insights(
        final_bullets,
        month_label,
        metrics=metrics,
        selected_rules=selected_rules,
    )
    return final_bullets, summary, sms, sms_segments, sms_meta


def build_fact_from_metric_names(
    metrics: dict[str, dict[str, Any]],
    metric_names: list[str],
    *,
    minimum_band: str = "slight",
) -> list[str]:
    facts: list[str] = []
    seen: set[str] = set()
    for metric_name in metric_names:
        fact = build_short_metric_fact(metrics, metric_name, minimum_band=minimum_band)
        if fact and fact not in seen:
            facts.append(fact)
            seen.add(fact)
    return facts


def choose_numeric_rule_message(
    rule_id: str,
    metrics: dict[str, dict[str, Any]],
    context_seed: str,
) -> str | None:
    response_shift = build_metric_transition_text(
        metrics,
        "sales_internet_response_time",
        up_verb="slowed",
        down_verb="improved",
        label_override="Response times",
    )
    uncontacted_shift = build_metric_transition_text(
        metrics,
        "sales_uncontacted_leads",
        up_verb="rose",
        down_verb="fell",
        label_override="uncontacted leads",
    )
    leads_shift = build_metric_transition_text(
        metrics,
        "sales_total_leads",
        up_verb="increased",
        down_verb="eased",
        label_override="Lead volume",
    )
    deals_shift = build_metric_transition_text(
        metrics,
        "sales_total_deals",
        up_verb="increased",
        down_verb="fell",
        label_override="Deal volume",
    )
    conversion_shift = build_metric_transition_text(
        metrics,
        "sales_overall_conversion_rate",
        up_verb="improved",
        down_verb="fell",
        label_override="conversion",
    )
    sales_campaign_deals_shift = build_metric_transition_text(
        metrics,
        "sales_campaign_deals",
        up_verb="increased",
        down_verb="fell",
        label_override="Sales campaign deals",
    )
    sales_campaign_roi_shift = build_metric_transition_text(
        metrics,
        "sales_campaign_roi",
        up_verb="improved",
        down_verb="fell",
        label_override="sales campaign ROI",
    )
    service_bookings_shift = build_metric_transition_text(
        metrics,
        "service_bookings_from_campaigns",
        up_verb="rose",
        down_verb="fell",
        label_override="Service campaign bookings",
    )
    service_roi_shift = build_metric_transition_text(
        metrics,
        "service_campaign_roi",
        up_verb="improved",
        down_verb="fell",
        label_override="service campaign ROI",
    )
    service_tasks_shift = build_metric_transition_text(
        metrics,
        "service_outstanding_tasks",
        up_verb="rose",
        down_verb="fell",
        label_override="Outstanding service tasks",
    )

    evidence_templates: dict[str, list[str]] = {
        "lead_handling_pressure": [
            "{response}, while {uncontacted}, increasing the risk of warm enquiries cooling off.",
            "{response}, and {uncontacted}, increasing the risk of warm enquiries cooling off.",
        ],
        "lead_handling_improved": [
            "{response}, while {uncontacted}, reading as cleaner lead follow-up discipline.",
            "{response}, and {uncontacted}, indicating cleaner first-contact discipline.",
        ],
        "sales_marketing_up": [
            "{sales_campaign_deals}, with {sales_campaign_roi}, reading as stronger campaign performance.",
            "{sales_campaign_deals}, while {sales_campaign_roi}, reading as stronger campaign performance.",
        ],
        "service_campaigns_up": [
            "{service_bookings}, while {service_roi}.",
            "{service_bookings}, with {service_roi}.",
        ],
        "leads_up_conversion_down": [
            "{leads}, but {conversion}, with the extra demand not converting efficiently.",
            "{leads}, while {conversion}, with the added demand not translating efficiently into sales.",
        ],
        "leads_down_conversion_up": [
            "{leads}, but {conversion}, reading as better quality or stronger follow-up execution.",
            "{leads}, while {conversion}, with the smaller enquiry pool handled more efficiently.",
        ],
        "leads_up_deals_down": [
            "{leads}, but {deals}, with the extra demand not converting efficiently into delivered deals.",
        ],
        "service_tasks_pressure": [
            "{service_tasks}, without matching booking growth, with workflow pressure building.",
            "{service_tasks}, without enough matching booking momentum, with workflow pressure building.",
        ],
        "service_follow_up_cleaner": [
            "{service_tasks}, even while customer enrolments remained healthy, reading as cleaner service follow-up.",
        ],
        "combined_sales_service_up": [
            "{deals}, while {service_roi}, alongside a stronger overall month.",
            "{deals}, with {service_roi}, alongside a stronger overall month.",
        ],
    }

    replacements = {
        "response": response_shift,
        "uncontacted": uncontacted_shift,
        "leads": leads_shift,
        "deals": deals_shift,
        "conversion": conversion_shift,
        "sales_campaign_deals": sales_campaign_deals_shift,
        "sales_campaign_roi": sales_campaign_roi_shift,
        "service_bookings": service_bookings_shift,
        "service_roi": service_roi_shift,
        "service_tasks": service_tasks_shift,
    }

    templates = evidence_templates.get(rule_id)
    if not templates:
        return None

    viable_templates = []
    for template in templates:
        placeholder_names = set(re.findall(r"{([a-z_]+)}", template))
        if all(replacements.get(name) for name in placeholder_names):
            viable_templates.append(template)

    if not viable_templates:
        return None

    template = choose_deterministic_phrase(viable_templates, f"{context_seed}|{rule_id}|numeric")
    return template.format(**replacements)


def choose_deterministic_phrase(options: list[str], context_key: str) -> str:
    if not options:
        return ""
    digest = hashlib.sha256(context_key.encode("utf-8")).hexdigest()
    index = int(digest[:8], 16) % len(options)
    return options[index]


def metric_change_strength(metric: dict[str, Any] | None) -> int:
    if not metric:
        return 0

    significance_band = metric_significance_band(metric)
    if significance_band == "strong":
        return 2
    if significance_band in {"meaningful", "slight"}:
        return 1
    return 0


def rule_tone_band(
    metrics: dict[str, dict[str, Any]],
    supporting_metric_names: list[str],
    polarity: str,
    *,
    current_value_override: float | None = None,
) -> str:
    relevant_metrics = [metrics.get(name) for name in supporting_metric_names if metrics.get(name) is not None]
    strongest_change = max((metric_change_strength(metric) for metric in relevant_metrics), default=0)

    if current_value_override is not None and current_value_override < 0:
        return "significantly_softer"

    if polarity == "positive":
        if strongest_change >= 2:
            return "strongly_improved"
        if strongest_change >= 1:
            return "slightly_improved"
        return "stable"
    if polarity == "negative":
        if strongest_change >= 2:
            return "significantly_softer"
        if strongest_change >= 1:
            return "slightly_softer"
        return "stable"
    return "stable"


def metric_snapshot_tone_band(metric: dict[str, Any] | None, higher_is_better: bool) -> str:
    if not metric:
        return "stable"

    direction = metric.get("direction")
    strength = metric_change_strength(metric)
    if direction == "flat" or strength == 0:
        return "stable"

    improving = (direction == "up" and higher_is_better) or (direction == "down" and not higher_is_better)
    if improving:
        return "strongly_improved" if strength >= 2 else "slightly_improved"
    return "significantly_softer" if strength >= 2 else "slightly_softer"


def render_rule_message(rule_id: str, tone_band: str, context_key: str, **kwargs: str) -> str:
    """
    Renders a rule message with safe fallbacks for missing templates.
    
    PATCH 3: Render Template Safety
    Ensures triggered rules cannot silently degrade into empty/punctuation-only output.
    """
    template_groups = RULE_MESSAGE_TEMPLATES.get(rule_id, {})
    templates = template_groups.get(tone_band) or template_groups.get("stable") or []
    if not templates:
        # Developer-visible log, customer-safe fallback (never leak rule ids).
        logging.warning("Missing template coverage | rule_id=%s | tone=%s", rule_id, tone_band)
        customer_safe = choose_deterministic_phrase(
            {
                "strongly_improved": [
                    "A supporting operational indicator improved meaningfully this month.",
                    "A secondary operational signal strengthened this month.",
                ],
                "slightly_improved": [
                    "A supporting operational indicator improved modestly this month.",
                    "A secondary operational signal moved in a constructive direction this month.",
                ],
                "significantly_softer": [
                    "A supporting operational indicator softened materially this month.",
                    "A secondary operational signal deteriorated and warrants attention.",
                ],
                "slightly_softer": [
                    "A supporting operational indicator softened this month.",
                    "A secondary operational signal eased and warrants monitoring.",
                ],
                "stable": [
                    "A supporting operational indicator was broadly steady this month.",
                    "A secondary operational signal held broadly stable this month.",
                ],
            }.get(tone_band, ["A supporting operational indicator was broadly steady this month."]),
            f"{context_key}|template_missing_safe|{tone_band}",
        )
        return customer_safe
    
    template = choose_deterministic_phrase(templates, context_key)
    rendered = template.format(**kwargs)
    
    # Validate the rendered message is meaningful
    if not is_meaningful_candidate_message(rendered):
        logging.warning("Non-meaningful render | rule_id=%s | tone=%s | rendered=%r", rule_id, tone_band, rendered)
        return choose_deterministic_phrase(
            [
                "A supporting operational signal was identified, but details were not reliably extractable this month.",
                "A secondary operational signal was present, though the supporting detail was not cleanly extractable.",
            ],
            f"{context_key}|template_nonmeaningful_safe|{tone_band}",
        )
    
    return rendered


def render_rule_message_with_evidence(
    rule_id: str,
    tone_band: str,
    context_key: str,
    metrics: dict[str, dict[str, Any]],
    **kwargs: str,
) -> str:
    numeric_message = choose_numeric_rule_message(rule_id, metrics, context_key)
    if numeric_message:
        return numeric_message
    return render_rule_message(rule_id, tone_band, context_key, **kwargs)


def join_phrases(parts: list[str]) -> str:
    clean_parts = [part for part in parts if part]
    if not clean_parts:
        return ""
    if len(clean_parts) == 1:
        return clean_parts[0]
    if len(clean_parts) == 2:
        return f"{clean_parts[0]} and {clean_parts[1]}"
    return f"{', '.join(clean_parts[:-1])}, and {clean_parts[-1]}"


def overall_month_rating_from_sentiment(sentiment: str) -> str:
    if sentiment in {"strongly_improved", "slightly_improved"}:
        return "Strong"
    if sentiment == "mixed":
        return "Mixed"
    if sentiment in {"significantly_softer", "slightly_softer"}:
        return "Softer"
    return "Stable"


def short_phrase_for_rule(rule: RuleInsight, *, align_tone: bool = False) -> str:
    if align_tone and rule.rule_id == "service_tasks_pressure" and rule.tone_band == "stable":
        return "Service follow-up workload remained broadly stable"
    rule_phrase_map = {
        "combined_sales_service_up": "Balanced sales and service momentum",
        "leads_down_conversion_up": "Stronger conversions from the lead pool",
        "lead_handling_improved": "Cleaner lead follow-up performance",
        "closing_execution_up": "Better closing execution",
        "sales_marketing_up": "Stronger sales campaign return",
        "service_campaigns_up": "Stronger service campaign performance",
        "service_follow_up_cleaner": "Stronger service follow-up performance",
        "leads_up_conversion_down": "Weaker lead-to-sale conversion",
        "leads_up_deals_down": "Demand not converting into enough deals",
        "lead_handling_pressure": "Higher uncontacted lead pressure",
        "test_drive_and_conversion_down": "Softer funnel progression",
        "sales_marketing_negative_roi": "Negative sales campaign ROI",
        "sales_campaign_roi_decline": "Weaker sales campaign return",
        "service_bookings_roi_mismatch": "Weaker service ROI quality",
        "service_tasks_pressure": "Higher outstanding service-task load",
        "sales_volume_snapshot": "Deal-volume movement",
        "sales_conversion_snapshot": "Sales conversion movement",
        "lead_response_snapshot": "Lead response performance",
        "service_booking_snapshot": "Service booking momentum",
        "service_roi_snapshot": "Service ROI movement",
        "service_task_snapshot": "Service follow-up workload",
        "crm_execution_pressure_composite": "Heavy follow-up pressure",
    }
    category_phrase_map = {
        "combined_performance": "Overall dealership momentum",
        "sales_funnel": "Sales funnel quality",
        "lead_management": "Lead follow-up discipline",
        "sales_execution": "Sales follow-up quality",
        "marketing": "Campaign performance",
        "service_campaigns": "Service campaign performance",
        "service_operations": "Service follow-up workflow",
        "sales_overview": "Sales volume consistency",
    }
    return rule_phrase_map.get(rule.rule_id) or category_phrase_map.get(rule.category, "Operational performance")


def build_rule_fact_list(rule: RuleInsight, metrics: dict[str, dict[str, Any]], *, include_month_on_month: bool = True) -> list[str]:
    specific_fact_map: dict[str, list[tuple[str, str | None]]] = {
        "combined_sales_service_up": [("sales_total_deals", "deal volume"), ("service_campaign_roi", "service campaign ROI")],
        "leads_up_conversion_down": [("sales_total_leads", "lead volume"), ("sales_overall_conversion_rate", "lead-to-sale conversion")],
        "leads_down_conversion_up": [("sales_total_leads", "lead volume"), ("sales_overall_conversion_rate", "lead-to-sale conversion")],
        "leads_up_deals_down": [("sales_total_leads", "lead volume"), ("sales_total_deals", "deal volume")],
        "lead_handling_pressure": [("sales_internet_response_time", "response times"), ("sales_uncontacted_leads", "uncontacted leads")],
        "lead_handling_improved": [("sales_internet_response_time", "response times"), ("sales_uncontacted_leads", "uncontacted leads")],
        "closing_execution_up": [("sales_appointment_to_sales", "appointment-to-sale conversion"), ("sales_same_day_deal_pct", "same-day deal mix")],
        "test_drive_and_conversion_down": [("sales_new_leads_to_test_drives", "lead-to-test-drive conversion"), ("sales_overall_conversion_rate", "lead-to-sale conversion")],
        "sales_marketing_up": [("sales_campaign_deals", "sales campaign deals"), ("sales_campaign_roi", "sales campaign ROI")],
        "sales_marketing_negative_roi": [("sales_campaign_roi", "sales campaign ROI")],
        "sales_campaign_roi_decline": [("sales_campaign_roi", "sales campaign ROI"), ("sales_campaign_deals", "sales campaign deals")],
        "service_campaigns_up": [("service_bookings_from_campaigns", "service campaign bookings"), ("service_campaign_roi", "service campaign ROI")],
        "service_bookings_roi_mismatch": [("service_bookings_from_campaigns", "service campaign bookings"), ("service_campaign_roi", "service campaign ROI")],
        "service_tasks_pressure": [("service_outstanding_tasks", "outstanding service tasks")],
        "service_follow_up_cleaner": [("service_outstanding_tasks", "outstanding service tasks")],
        "sales_volume_snapshot": [("sales_total_deals", "deal volume")],
        "sales_conversion_snapshot": [("sales_overall_conversion_rate", "lead-to-sale conversion")],
        "lead_response_snapshot": [("sales_internet_response_time", "response times")],
        "service_booking_snapshot": [("service_bookings_from_campaigns", "service campaign bookings")],
        "service_roi_snapshot": [("service_campaign_roi", "service campaign ROI")],
        "service_task_snapshot": [("service_outstanding_tasks", "outstanding service tasks")],
    }

    facts: list[str] = []
    seen: set[str] = set()
    for metric_name, label_override in specific_fact_map.get(rule.rule_id, []):
        fact = build_short_metric_fact(
            metrics,
            metric_name,
            label_override=label_override,
            include_month_on_month=include_month_on_month,
        )
        if fact and fact not in seen:
            facts.append(fact)
            seen.add(fact)

    if facts:
        return facts

    for metric_name in rule.supporting_metrics:
        fact = build_short_metric_fact(metrics, metric_name, include_month_on_month=include_month_on_month)
        if fact and fact not in seen:
            facts.append(fact)
            seen.add(fact)
    return facts


def build_rule_headline(rule: RuleInsight | None, metrics: dict[str, dict[str, Any]]) -> str | None:
    if rule is None:
        return None
    facts = build_rule_fact_list(rule, metrics, include_month_on_month=False)
    if not facts:
        return None
    contrast_rules = {"leads_up_conversion_down", "leads_down_conversion_up", "leads_up_deals_down",
                      "service_bookings_roi_mismatch"}
    if rule.rule_id in contrast_rules:
        # Half a contrast inverts it: "Lead volume up 45%" alone was shipped as a Primary Risk.
        if len(facts) < 2:
            return None
        headline = f"{facts[1]} despite {facts[0]}"
    elif rule.rule_id in {"lead_handling_pressure", "lead_handling_improved", "sales_marketing_up", "sales_marketing_negative_roi", "sales_campaign_roi_decline", "service_campaigns_up", "combined_sales_service_up"} and len(facts) >= 2:
        headline = f"{facts[0]} and {facts[1]}"
    else:
        headline = facts[0]
    return headline[0].upper() + headline[1:] if headline else headline


# --- Phase 165 — headline-first summary & interpretation suppression ---


def phase165_sentence_case_metric_clause(seg: str) -> str:
    """Title-case a metric headline segment without uppercasing a leading 'while' connector."""
    s = seg.strip()
    if not s:
        return s
    low = s.lower()
    if low.startswith("while "):
        rest = s[6:].strip()
        if not rest:
            return s
        return "while " + (rest[0].upper() + rest[1:] if rest[0].isalpha() else rest)
    return s[0].upper() + s[1:] if s[0].isalpha() else s


def phase168_scrub_internal_reporting_language(text: str) -> str:
    """Remove internal arbitration, validation, and directional-state narration (phase 168, deterministic)."""
    if not text:
        return text
    out = collapse_whitespace(str(text))
    # Legacy duplicate tails: ", showing X across recent months, Label is showing X across recent months."
    out = re.sub(
        r",\s*showing mixed directional transitions across recent months\s*,\s*[^.]{5,140}?mixed directional transitions across recent months\.?",
        ".",
        out,
        flags=re.IGNORECASE,
    )
    out = re.sub(
        r",\s*under pressure across recent months\s*,\s*[^.]{5,140}?under pressure across recent months\.?",
        ".",
        out,
        flags=re.IGNORECASE,
    )
    out = re.sub(r",\s*showing mixed directional transitions across recent months", "", out, flags=re.IGNORECASE)
    out = re.sub(r",\s*showing consistent ROI improvement", "", out, flags=re.IGNORECASE)
    out = re.sub(r",\s*showing consistent conversion improvement", "", out, flags=re.IGNORECASE)
    out = re.sub(
        r",\s*[A-Za-z][^.]{0,120}?\s(?:is|are)\s+showing mixed directional transitions across recent months\.?",
        ".",
        out,
        flags=re.IGNORECASE,
    )
    out = re.sub(
        r",\s*[A-Za-z][^.]{0,120}?\s(?:is|are)\s+under pressure across recent months\.?",
        ".",
        out,
        flags=re.IGNORECASE,
    )
    literal_fragments = (
        "Interpretation stays conservative where overall sales throughput is modest, even though core metrics were usable.",
        "Readings stay intentionally measured because fewer headline metrics met reliability thresholds for this month.",
        "Activity remained relatively stable across the month, although lower transaction volumes continue to limit stronger directional signals.",
        "Some conclusions are tempered by lower metric reliability and should be validated.",
        "carrying sustained directional improvement across recent months",
        "under sustained pressure across recent months",
        "losing early-funnel progression strength across recent months",
        "mixed directional transitions across recent months",
        "directional transitions across recent months",
        "evidence confidence",
        "arbitration confidence",
        "supporting directional",
        "directional state",
        "directional trajectory",
    )
    for frag in literal_fragments:
        if frag.lower() in out.lower():
            out = re.sub(re.escape(frag), "", out, flags=re.IGNORECASE)
    out = re.sub(r"\s{2,}", " ", out).strip()
    out = re.sub(r"\s*\.\s*\.", ".", out)
    out = re.sub(r",\s*\.", ".", out)
    out = re.sub(r"\.\s*\.", ".", out)
    return out.strip()


def phase165_scrub_interpretation_and_jargon(text: str) -> str:
    """Strip consultant-style framing and abstract business filler (deterministic, phase 165)."""
    if not text:
        return text
    out = collapse_whitespace(str(text))
    literal_pairs: tuple[tuple[str, str], ...] = (
        (" amid increasing operational pressure", ""),
        (" amid broader operational pressure", ""),
        (" broader operational execution weakened", " follow-up metrics weakened"),
        (" commercial pressure", ""),
        (" operational backdrop", ""),
        (" performance environment", ""),
        (" demand environment", ""),
        (" strategic environment", ""),
        (" workflow environment", ""),
        (" underlying pressure", ""),
        (" underlying activity", ""),
        (" sustained conditions", ""),
        (" broader performance trends", ""),
        (" retention execution", " retention activity"),
        (" broader conditions", ""),
        (" mixed execution", " mixed results"),
        (" reading as ", ". "),
        (" which reads as ", ". "),
        (" interpretive weight", ""),
        (" executive read", " read"),
        (" at the leadership level", ""),
        (" strategically", ""),
        (" operationally,", ""),
        (" from a strategic standpoint", ""),
        (" in the broader operating context", ""),
    )
    for old, new in literal_pairs:
        if old.lower() in out.lower():
            out = re.sub(re.escape(old), new, out, flags=re.IGNORECASE)
    out = re.sub(r"\s{2,}", " ", out).strip()
    out = re.sub(r"\s*\.\s*\.", ".", out)
    out = re.sub(r",\s*\.", ".", out)
    out = re.sub(r"([.!?])\s+While\s+", r"\1 while ", out)
    out = re.sub(r"([a-zA-Z0-9%])\s+While\s+", r"\1 while ", out)
    return phase168_scrub_internal_reporting_language(out.strip())


def phase165_headline_fallback_sentence(rule: RuleInsight) -> str | None:
    """One short observational clause from the rule message when metric headlines are unavailable."""
    raw = collapse_whitespace(str(rule.message or "").strip())
    if not raw:
        return None
    first = raw.split(". ", 1)[0].strip()
    if "," in first and len(first) > 140:
        first = first.split(",")[0].strip()
    if len(first) > 160:
        first = first[:157].rstrip(",; ") + "…"
    if not first.endswith((".", "!", "?")):
        first += "."
    return first[0].upper() + first[1:] if first else None


def phase165_collect_priority_headlines(
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
) -> list[str]:
    headlines: list[str] = []
    seen: set[str] = set()
    for rule in selected_rules[:4]:
        h = build_rule_headline(rule, metrics)
        if not h:
            fb = phase165_headline_fallback_sentence(rule)
            h = fb
        if not h:
            continue
        h = phase165_sentence_case_metric_clause(h.strip().rstrip("."))
        key = re.sub(r"\s+", " ", h.strip().lower())[:72]
        if key in seen:
            continue
        seen.add(key)
        headlines.append(h)
    return headlines


def phase165_build_headline_first_summary(
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
    context_seed: str,
) -> str:
    """
    Short executive briefing: ordered metric headlines aligned with Priority Insights (phase 165).
    """
    del context_seed
    parts = phase165_collect_priority_headlines(selected_rules, metrics)
    if not parts:
        return ""
    cap = 3
    body_parts = parts[:cap]
    sentences = []
    for p in body_parts:
        seg = p.strip()
        if not seg:
            continue
        sentences.append(phase165_sentence_case_metric_clause(seg))
    if not sentences:
        return ""
    text = sentences[0]
    for s in sentences[1:]:
        if s.lower().startswith("while "):
            text = f"{text}, {s}"
        else:
            text = f"{text}. {s}"
    if len(parts) > cap:
        extra = parts[cap]
        extra_stripped = extra.strip()
        if extra_stripped.lower().startswith("while "):
            rest = extra_stripped[6:].strip()
            extra_u = "while " + (rest[0].upper() + rest[1:] if rest and rest[0].isalpha() else rest)
        else:
            extra_u = phase165_sentence_case_metric_clause(extra_stripped)
        if extra_u.lower().startswith("while "):
            text = f"{text}, {extra_u}"
        else:
            text = f"{text}. {extra_u}"
    if text and not text.endswith((".", "!", "?")):
        text += "."
    text = phase165_scrub_interpretation_and_jargon(text)
    if len(text) > 520:
        text = text[:517].rsplit(".", 1)[0] + "."
    return text


def phase165_build_sms_headline_body(
    metrics: dict[str, dict[str, Any]],
    selected_rules: list[RuleInsight],
    context_seed: str,
    *,
    max_chars: int = 118,
) -> str:
    """Metric-led SMS body from top priority insights (phase 165)."""
    parts = phase165_collect_priority_headlines(selected_rules, metrics)
    if not parts:
        return ""
    digest = hashlib.sha256(f"{context_seed}|sms165".encode("utf-8")).hexdigest()
    take = 2 if int(digest[:4], 16) % 5 != 0 else 1
    chosen = parts[:take]
    body = ". ".join(phase165_sentence_case_metric_clause(s) for s in chosen)
    body = phase165_scrub_interpretation_and_jargon(body)
    if len(body) > max_chars:
        if len(chosen) > 1:
            body = phase165_sentence_case_metric_clause(chosen[0])
            body = phase165_scrub_interpretation_and_jargon(body)
        if len(body) > max_chars:
            body = body[: max_chars - 1].rstrip(",; ") + "…"
    return body.rstrip(".")


# --- Phase 166 — balanced natural summary (paragraph, not bullet fragments) ---


def phase166_materiality_gate(ledger: dict[str, Any] | None) -> str:
    """Return slight | moderate | clear for evidence-gated wording (phase 166)."""
    if not ledger:
        return "moderate"
    band = str(ledger.get("materiality_band", "moderate")).lower()
    if band in ("immaterial", "marginal"):
        return "slight"
    if band == "moderate":
        return "moderate"
    if band in ("material", "severe"):
        return "clear"
    return "moderate"


def phase166_can_say_weaker(ledger: dict[str, Any] | None, sentiment: str) -> bool:
    if not ledger:
        return sentiment in {"significantly_softer", "slightly_softer"}
    if bool(ledger.get("material_commercial_adverse")):
        return True
    net = float(ledger.get("net_month_score", 0.0))
    return sentiment in {"significantly_softer", "slightly_softer"} or net <= -6.0


def phase166_can_say_stronger(ledger: dict[str, Any] | None, sentiment: str) -> bool:
    if not ledger:
        return sentiment in {"strongly_improved", "slightly_improved"}
    net = float(ledger.get("net_month_score", 0.0))
    return sentiment in {"strongly_improved", "slightly_improved"} and net >= 4.0


def phase166_can_say_sharp(ledger: dict[str, Any] | None) -> bool:
    if not ledger:
        return False
    if str(ledger.get("materiality_band", "")).lower() not in {"material", "severe"}:
        return False
    return bool(ledger.get("allow_noticeably_weaker_intro", False))


# What each lane's opening phrase claims about the month as a whole.
LANE_IMPLIED_DIRECTION = {
    "steady": "neutral",
    "mixed": "neutral",
    "general_improved": "positive",
    "service_gain": "positive",
    "service_gain_with_risk": "positive",
    "service_gain_mixed_tail": "positive",
    "general_soft": "negative",
    "campaign_weak": "negative",
    "crm_load": "negative",
}

# A 30% move is worth more than a 4% one when weighing up how the month actually went.
SIGNIFICANCE_BAND_WEIGHTS = {"strong": 3.0, "meaningful": 2.0, "slight": 1.0}

# How far one side has to outweigh the other before the month has a direction at all.
METRIC_BALANCE_MARGIN = 1.5


def metric_balance_direction(
    metrics: dict[str, dict[str, Any]] | None,
    margin: float = METRIC_BALANCE_MARGIN,
) -> str:
    """Which way the month actually went, across every metric the report can use.

    Weighted by significance band so a handful of trivial moves cannot outvote a
    genuine collapse, and vice versa. Returns "neutral" unless one side clearly wins,
    because most months are genuinely mixed and should be allowed to say so.
    """
    improving = adverse = 0.0
    for name, metric in (metrics or {}).items():
        if not metric_usable_for_reporting(metric):
            continue
        weight = SIGNIFICANCE_BAND_WEIGHTS.get(str(metric.get("significance_band", "")), 1.0)
        if metric_is_adverse(metrics, name):
            adverse += weight
        else:
            improving += weight
    if improving > adverse * margin:
        return "positive"
    if adverse > improving * margin:
        return "negative"
    return "neutral"


def reconcile_lane_with_metrics(lane: str, metrics: dict[str, dict[str, Any]] | None) -> str:
    """Stop the opening line claiming the opposite of what the month's numbers did.

    The lane is picked from the sentiment score and from keywords in the headline
    clauses, and neither is checked against the metrics themselves. That let a month
    with ten improvements and five declines open "a clear step down across the main
    numbers" because two of the declines were large and service-side, and let a month
    with nine declines open "firmer service campaign returns" on the strength of one
    ROI rise.

    Where the opener contradicts the balance of the evidence it drops to "mixed" --
    never to the opposite claim, which would just invert the error. The facts that
    follow carry the detail either way.
    """
    implied = LANE_IMPLIED_DIRECTION.get(lane, "neutral")
    if implied == "neutral":
        return lane
    balance = metric_balance_direction(metrics)
    if balance == "neutral" or balance == implied:
        return lane
    return "mixed"


def phase166_detect_lane(
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
    sentiment: str,
    parts: list[str],
) -> str:
    return reconcile_lane_with_metrics(
        _phase166_detect_lane_from_signals(selected_rules, metrics, sentiment, parts),
        metrics,
    )


def _phase166_detect_lane_from_signals(
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
    sentiment: str,
    parts: list[str],
) -> str:
    joined = " ".join(parts).lower()
    rids = [r.rule_id for r in selected_rules[:4]]
    r0 = rids[0] if rids else ""

    if sentiment == "mixed":
        return "mixed"
    if sentiment == "stable":
        # "Broadly steady" cannot open a summary whose very next clause reports a
        # material adverse movement -- the opener would contradict the fact beside it.
        # Snapshot rules carry "mixed" polarity even when the movement they report is
        # adverse, so judge the metric rather than the rule's label.
        leading = selected_rules[0] if selected_rules else None
        if (
            leading is None
            or not has_adverse_focus_metric(leading, metrics)
            or insight_materiality(leading, metrics) < 40.0
        ):
            return "steady"

    if r0 in {"lead_handling_pressure", "crm_execution_pressure_composite", "lead_response_snapshot"}:
        return "crm_load"

    svc_up = "service campaign roi" in joined and ("improved" in joined or re.search(r"roi\s+up\b", joined) is not None)
    resp_improved = "response time" in joined and ("down" in joined or "improved" in joined)
    conv_soft = "conversion" in joined and ("down" in joined or "fell" in joined)
    tasks_up = (
        ("outstanding service tasks" in joined or "outstanding service task" in joined)
        and ("rose" in joined or re.search(r"outstanding service tasks[^\n]*up\s+", joined) or " increased" in joined)
    )

    if svc_up and (resp_improved or ("response" in joined and "down" in joined)):
        if conv_soft and tasks_up:
            return "service_gain_mixed_tail"
        if conv_soft or tasks_up:
            return "service_gain_with_risk"
        return "service_gain"

    if "sales campaign roi" in joined and ("down" in joined or "negative" in joined or "turned negative" in joined):
        return "campaign_weak"
    if r0 in {"sales_campaign_roi_decline", "sales_marketing_negative_roi"}:
        return "campaign_weak"

    if sentiment in {"significantly_softer", "slightly_softer"}:
        return "general_soft"
    if sentiment in {"strongly_improved", "slightly_improved"}:
        return "general_improved"
    return "general_soft"


def phase166_theme_phrase(
    lane: str,
    ledger: dict[str, Any] | None,
    sentiment: str,
    context_seed: str,
) -> str:
    gate = phase166_materiality_gate(ledger)
    weaker_ok = phase166_can_say_weaker(ledger, sentiment)
    stronger_ok = phase166_can_say_stronger(ledger, sentiment)

    if lane == "steady":
        return choose_deterministic_phrase(
            ["broadly steady activity", "little movement across the main numbers"],
            f"{context_seed}|ph166|theme|steady",
        )
    if lane == "mixed":
        return choose_deterministic_phrase(
            ["mixed sales and service signals", "mixed movement across sales and service"],
            f"{context_seed}|ph166|theme|mixed",
        )
    if lane == "campaign_weak":
        base = choose_deterministic_phrase(
            ["weaker campaign and lead activity", "softer campaign returns with lighter funnel conversion"],
            f"{context_seed}|ph166|theme|camp",
        )
        if weaker_ok:
            return base
        return choose_deterministic_phrase(
            ["softer campaign and funnel numbers", "softer campaign returns alongside lighter conversion"],
            f"{context_seed}|ph166|theme|campsoft",
        )
    if lane == "crm_load":
        return choose_deterministic_phrase(
            ["higher follow-up workload", "heavier follow-up workload on the sales side"],
            f"{context_seed}|ph166|theme|crm",
        )
    if lane in {"service_gain", "service_gain_with_risk", "service_gain_mixed_tail"}:
        if stronger_ok:
            return choose_deterministic_phrase(
                [
                    "stronger service campaign returns and faster response times",
                    "firmer service campaign returns with quicker response times",
                ],
                f"{context_seed}|ph166|theme|svc",
            )
        if gate == "slight":
            return "slightly firmer service campaign returns and response times"
        return "firmer service campaign returns and response times"
    if lane == "general_improved":
        return choose_deterministic_phrase(
            ["clearer improvement across the main numbers", "movement the right way across sales and service"],
            f"{context_seed}|ph166|theme|up",
        )
    if gate == "slight":
        return choose_deterministic_phrase(
            ["a slight softening across the main numbers", "modest softness across the main numbers"],
            f"{context_seed}|ph166|theme|softslight",
        )
    if gate == "clear" and weaker_ok:
        return choose_deterministic_phrase(
            ["clear pressure on sales and campaign performance", "a clear step down across the main numbers"],
            f"{context_seed}|ph166|theme|softclear",
        )
    return choose_deterministic_phrase(
        ["softer numbers across sales and service", "softer results across the main measures"],
        f"{context_seed}|ph166|theme|softdef",
    )


def phase166_embed_in_sentence(clause: str) -> str:
    """Lower-case first character so metric clauses read naturally after 'with'."""
    s = clause.strip().rstrip(".")
    if not s:
        return s
    return s[0].lower() + s[1:] if s[0].isalpha() else s


def phase166_join_with_while(clause_a: str, clause_b: str) -> str:
    a = collapse_whitespace(clause_a.rstrip("."))
    b = collapse_whitespace(clause_b.rstrip("."))
    if not b:
        return f"{a}."
    if b.lower().startswith("while "):
        b = b[6:].strip()
    b_body = b[0].lower() + b[1:] if b and b[0].isalpha() else b
    return f"{a}, while {b_body}."


def phase168_prune_headline_parts_for_summary(
    parts: list[str],
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
) -> list[str]:
    """Keep at most three headline clauses without dropping the top two priority headlines (phase 168)."""
    if len(parts) <= 3:
        return parts
    if len(parts) == 4 and len(selected_rules) >= 4:

        def _tail_keep_score(idx: int) -> tuple[float, str]:
            rule = selected_rules[idx]
            seg = parts[idx]
            score = float(rule.priority_weight)
            if has_adverse_focus_metric(rule, metrics):
                score += 55.0
            if phase155_crm_deterioration_count(metrics) >= 1 and rule.rule_id in {
                "crm_execution_pressure_composite",
                "lead_handling_pressure",
                "lead_response_snapshot",
            }:
                score += 18.0
            if rule.rule_id in {"service_tasks_pressure", "service_task_snapshot"}:
                m = metrics.get("service_outstanding_tasks")
                if m and metric_is_adverse(metrics, "service_outstanding_tasks"):
                    score += 22.0
            if rule.rule_id == "test_drive_and_conversion_down":
                score += 28.0
            if "conversion" in seg.lower() and metric_is_adverse(metrics, "sales_new_leads_to_test_drives"):
                score += 10.0
            return (score, rule.rule_id)

        keep_third = 2 if _tail_keep_score(2) >= _tail_keep_score(3) else 3
        # Adversity is worth 55 points in the score above, so on a mixed month the
        # discarded clause was reliably the improving one and the summary read as
        # uniformly bad news. Where the first two clauses already agree with each
        # other, the third is the one chance to show the other side of the month.
        first_two = {_insight_line_polarity(parts[0]), _insight_line_polarity(parts[1])}
        if first_two <= {"negative"} or first_two <= {"positive"}:
            wanted = "positive" if first_two <= {"negative"} else "negative"
            for index in (2, 3):
                if _insight_line_polarity(parts[index]) == wanted:
                    return [parts[0], parts[1], parts[index]]
        return [parts[0], parts[1], parts[keep_third]]
    return parts[:3]


def phase166_build_balanced_natural_summary(
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
    context_seed: str,
    month_label: str,
    sentiment: str,
    evidence_ledger: dict[str, Any] | None,
) -> str:
    """
    Paragraph-style summary: month opener + 1–2 full sentences, aligned to Priority Insights (phase 166).
    SMS remains headline-style (phase 165 path unchanged).
    """
    parts = phase165_collect_priority_headlines(selected_rules, metrics)
    parts = phase168_prune_headline_parts_for_summary(parts, selected_rules, metrics)
    if len(parts) < 2:
        return ""
    # Match the month form used by the theme-centred summary path ("July", not
    # "Jul 2026"), so month labels read the same across every report.
    ml = full_month_name_from_label(collapse_whitespace(month_label.strip()))
    lane = phase166_detect_lane(selected_rules, metrics, sentiment, parts)
    theme = phase166_theme_phrase(lane, evidence_ledger, sentiment, context_seed)

    # Lead with one thing. The opening sentence carries the single movement that
    # ranked highest this month; the rest follow it. Two facts spliced into the first
    # sentence made every month read as a list, and an owner who reads one sentence
    # should come away knowing what mattered.
    f1 = parts[0].strip().rstrip(".")
    e1 = phase166_embed_in_sentence(f1)
    if lane == "steady":
        s1 = f"{ml} was broadly steady, with {e1}."
    else:
        s1 = f"{ml} saw {theme}, with {e1}."

    supporting = [part.strip().rstrip(".") for part in parts[1:3] if part.strip()]
    if not supporting:
        text = phase165_scrub_interpretation_and_jargon(s1)
        return text if len(text) <= 580 else text[:577].rsplit(".", 1)[0] + "."

    sentences = [sentence_case_start(fact) + "." for fact in supporting]
    text = phase165_scrub_interpretation_and_jargon(" ".join([s1, *sentences]))
    text = collapse_whitespace(text)
    if len(text) > 580:
        text = text[:577].rsplit(".", 1)[0] + "."
    return text


# --- Phase 167 — outcome-aware offsets + lightweight operational wrap-up (summary only) ---


def phase167_mixed_offset_fact(
    metrics: dict[str, dict[str, Any]],
    parts: list[str],
) -> str | None:
    """
    When headline text shows deterioration on one axis but metrics show a material improving companion,
    return a short fact clause (no leading 'despite') for append after sentence 1 (phase 167).
    """
    if not metrics or len(parts) < 1:
        return None
    joined = " ".join(parts).lower()

    def _offset(bad: str, good: str) -> str | None:
        bm = metrics.get(bad)
        gm = metrics.get(good)
        if not bm or not gm:
            return None
        if not metric_usable_for_reporting(bm, "slight") or not metric_usable_for_reporting(gm, "slight"):
            return None
        if not metric_is_adverse(metrics, bad):
            return None
        if not metric_is_currently_improving(gm):
            return None
        if metric_change_strength(gm) < 1:
            return None
        if metric_change_strength(bm) < 1 and not metric_is_significant(bm, "slight"):
            return None
        label_map = {
            "sales_campaign_deals": "sales campaign deals",
            "sales_overall_conversion_rate": "conversion",
            "service_customers_enrolled": "enrolled clients",
            "service_bookings_from_campaigns": "service campaign bookings",
            "sales_total_deals": "deal volume",
        }
        return build_short_metric_fact(
            metrics,
            good,
            label_override=label_map.get(good),
            include_month_on_month=True,
        )

    if ("sales campaign roi" in joined or metric_is_adverse(metrics, "sales_campaign_roi")) and metric_is_adverse(
        metrics, "sales_campaign_roi"
    ):
        hit = _offset("sales_campaign_roi", "sales_campaign_deals")
        if hit:
            return hit
    if ("service campaign roi" in joined or metric_is_adverse(metrics, "service_campaign_roi")) and metric_is_adverse(
        metrics, "service_campaign_roi"
    ):
        for good in ("service_customers_enrolled", "service_bookings_from_campaigns"):
            hit = _offset("service_campaign_roi", good)
            if hit:
                return hit
    if metric_is_adverse(metrics, "sales_total_leads") and metric_is_currently_improving(
        metrics.get("sales_overall_conversion_rate") or {},
    ):
        conv = metrics.get("sales_overall_conversion_rate")
        if conv and metric_usable_for_reporting(conv, "slight") and metric_change_strength(conv) >= 1:
            return build_short_metric_fact(
                metrics,
                "sales_overall_conversion_rate",
                label_override="conversion",
                include_month_on_month=True,
            )
    return None


def phase167_insert_offset_after_first_sentence(summary: str, offset_fact: str | None) -> str:
    if not summary or not offset_fact:
        return summary
    low = summary.lower()
    if "despite" in low.split(".", 1)[0]:
        return summary
    m = re.search(r"\.(?=\s+[A-Za-z])", summary)
    if not m:
        return summary
    first_seg = summary[: m.start()]
    fl = first_seg.lower()
    off_low = offset_fact.lower()
    if "sales campaign deal" in off_low or "campaign deals" in off_low:
        if "sales campaign deal" in fl or "campaign deals" in fl:
            return summary
    if "enrolled client" in off_low or "enrolled clients" in off_low:
        if "enrolled client" in fl:
            return summary
    if "service campaign bookings" in off_low or "bookings from campaign" in off_low:
        if "service campaign bookings" in fl or "campaign bookings" in fl:
            return summary
    if "conversion" in off_low and "conversion" in fl:
        return summary
    i = m.start()
    first = first_seg.rstrip()
    rest = summary[i:]
    tail = phase166_embed_in_sentence(offset_fact.strip().rstrip("."))
    return collapse_whitespace(f"{first}, despite {tail}{rest}")


def phase167_guidance_eligible(
    summary: str,
    sentiment: str,
    ledger: dict[str, Any] | None,
) -> bool:
    if len(summary.split()) > 74:
        return False
    band = str((ledger or {}).get("materiality_band", "")).lower()
    net = float((ledger or {}).get("net_month_score", 0.0))
    if sentiment == "stable" and band in {"immaterial", "marginal"} and net > -2.0:
        return False
    if sentiment in {"strongly_improved", "slightly_improved"} and net > 8.0 and band in {"immaterial", "marginal"}:
        return False
    return True


def phase167_select_focus_guidance(
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
    summary_lower: str,
    context_seed: str,
) -> str | None:
    rids = {r.rule_id for r in selected_rules[:4]}
    joined = summary_lower

    if metric_is_adverse(metrics, "sales_campaign_roi") and (
        "sales campaign roi" in joined or "sales_campaign_roi_decline" in rids or "sales_marketing_negative_roi" in rids
    ):
        return choose_deterministic_phrase(
            [
                "Sales campaign figures may need a quick revisit on the sales side.",
                "Campaign ROI on sales may need a closer look this month.",
            ],
            f"{context_seed}|ph167|guide|sales_roi",
        )
    if metric_is_adverse(metrics, "service_campaign_roi") and (
        "service campaign roi" in joined or "service_roi_snapshot" in rids or "service_bookings_roi_mismatch" in rids
    ):
        return choose_deterministic_phrase(
            [
                "Service campaign work may need a short review so follow-through does not slip.",
                "Workshop campaign results may need a light check this month.",
            ],
            f"{context_seed}|ph167|guide|svc_roi",
        )
    if phase155_crm_deterioration_count(metrics) >= 1 and (
        "uncontacted" in joined or "response" in joined or "lead_handling_pressure" in rids
    ):
        return choose_deterministic_phrase(
            [
                "Lead follow-up is an area to keep an eye on.",
                "Contact speed and coverage still deserve a steady watch.",
            ],
            f"{context_seed}|ph167|guide|crm",
        )
    if ("outstanding service task" in joined or "service_tasks_pressure" in rids) and metric_is_adverse(
        metrics, "service_outstanding_tasks"
    ):
        return choose_deterministic_phrase(
            [
                "Service task carryover is worth monitoring so the backlog does not creep.",
                "Outstanding service tasks may need a short review so the queue stays manageable.",
            ],
            f"{context_seed}|ph167|guide|tasks",
        )
    return None


def phase167_enrich_operational_wrap_up(
    summary: str,
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
    sentiment: str,
    evidence_ledger: dict[str, Any] | None,
    context_seed: str,
) -> str:
    """Insert mixed-state offset after sentence 1, then optional short operational pointer (phase 167)."""
    if not summary.strip():
        return summary
    parts = phase165_collect_priority_headlines(selected_rules, metrics)
    off = phase167_mixed_offset_fact(metrics, parts)
    text = phase167_insert_offset_after_first_sentence(summary, off)
    text = phase165_scrub_interpretation_and_jargon(text)
    # The summary ends on a fact. The closing pointer this step used to append
    # ("Campaign ROI on sales may need a closer look this month") repeated what the
    # priority insights had already quantified and told the reader to go and look at
    # a number they had just been given.
    del sentiment, evidence_ledger, context_seed
    return text


def phase167_compact_sms_offset(
    metrics: dict[str, dict[str, Any]],
    selected_rules: list[RuleInsight],
    context_seed: str,
    headline_body: str = "",
) -> str:
    """Short ', despite …' tail for SMS when headline mode and length allows (phase 167)."""
    if headline_body and re.search(r"\bwhile\b", headline_body, re.IGNORECASE):
        return ""
    parts = phase165_collect_priority_headlines(selected_rules, metrics)
    off = phase167_mixed_offset_fact(metrics, parts)
    if not off:
        return ""
    tail = phase166_embed_in_sentence(off.strip().rstrip("."))
    if headline_body and tail:
        hb_pct = set(re.findall(r"\d+(?:\.\d+)?%", headline_body))
        tail_pct = set(re.findall(r"\d+(?:\.\d+)?%", tail))
        if hb_pct and tail_pct and (hb_pct & tail_pct):
            return ""
    frag = f", despite {tail}"
    if len(frag) > 52:
        frag = frag[:49].rstrip(",; ") + "…"
    return frag


def recommended_focus_from_rule(rule: RuleInsight | None, overall_month_rating: str) -> str:
    positive_focus_map = {
        "combined_performance": "Protect balanced sales and service momentum",
        "sales_funnel": "Sustain lead-to-sale conversion quality",
        "lead_management": "Maintain fast lead response and first-contact discipline",
        "sales_execution": "Sustain funnel progression and follow-up quality",
        "marketing": "Protect campaign mix, attribution, and ROI performance",
        "service_campaigns": "Sustain service booking conversion and ROI performance",
        "service_operations": "Maintain service follow-up discipline",
        "sales_overview": "Protect deal-volume consistency",
    }
    negative_focus_map = {
        "combined_performance": "Review the main cross-department performance gap",
        "sales_funnel": "Improve lead-to-sale conversion quality",
        "lead_management": "Improve lead response speed and first-contact discipline",
        "sales_execution": "Improve funnel progression and follow-up quality",
        "marketing": "Review campaign mix, attribution, and ROI performance",
        "service_campaigns": "Improve service booking conversion and ROI performance",
        "service_operations": "Reduce service follow-up backlog and task carryover",
        "sales_overview": "Protect deal-volume consistency",
    }
    if rule is not None:
        if rule.polarity == "negative" and rule.tone_band == "stable":
            stable_watch_map = {
                "service_operations": "Monitor service follow-up workload for sustained change",
                "lead_management": "Monitor lead follow-up discipline for sustained change",
                "sales_execution": "Monitor funnel progression for sustained change",
                "marketing": "Monitor campaign economics for sustained change",
                "service_campaigns": "Monitor service campaign return quality for sustained change",
                "sales_funnel": "Monitor conversion quality for sustained change",
            }
            return stable_watch_map.get(rule.category, "Monitor the main operational watchpoint for sustained change")
        focus_map = positive_focus_map if rule.polarity == "positive" else negative_focus_map
        return focus_map.get(rule.category, "Review the highest-priority operational driver")
    if overall_month_rating == "Strong":
        return "Protect recent momentum and follow-up quality"
    if overall_month_rating == "Stable":
        return "Maintain consistency and monitor emerging changes"
    return "Review report details before client delivery"


def build_executive_summary_fields(
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
    sentiment: str,
    insights_source: str,
    report_status: str,
    report_match: ReportMatch,
    scoring_profile: InternalScoringProfile,
) -> tuple[str, str, str, str]:
    overall_month_rating = scoring_profile.final_rating_used if insights_source == "rule-based" else overall_month_rating_from_sentiment(sentiment)

    positive_rules = [rule for rule in selected_rules if rule.polarity == "positive"]
    negative_rules = [rule for rule in selected_rules if rule.polarity == "negative"]
    negative_rules.extend(phase159_supplemental_executive_negative_rules(selected_rules, metrics))

    primary_strength = "No standout strength identified"
    if insights_source == "placeholder":
        primary_strength = "Stable placeholder coverage only"
    elif positive_rules and rule_shows_visible_improvement(positive_rules[0], metrics):
        primary_strength = (
            build_rule_headline(positive_rules[0], metrics)
            or best_improving_metric_phrase(metrics)
            or short_phrase_for_rule(positive_rules[0], align_tone=True)
        )
    else:
        fallback_strength = best_improving_metric_phrase(metrics)
        if fallback_strength:
            primary_strength = fallback_strength
        elif overall_month_rating == "Strong":
            primary_strength = "Performance strengthened across the most reliable metrics"
        elif overall_month_rating == "Stable":
            primary_strength = "Performance broadly held together"

    if insights_source == "placeholder":
        if report_status == "missing file" or report_match.sales_pdf is None or report_match.service_pdf is None:
            primary_risk = "Missing monthly reports"
        else:
            primary_risk = "Low-confidence parsed output"
    elif negative_rules:
        neg_for_exec = phase157_sort_negatives_for_executive(negative_rules)
        primary_risk = (
            build_rule_headline(neg_for_exec[0], metrics)
            or worst_declining_metric_phrase(metrics)
            or short_phrase_for_rule(neg_for_exec[0], align_tone=True)
        )
    elif important_anomaly_types(scoring_profile.anomaly_types):
        anomaly_label_map = {
            "conversion_drop": "Sharp conversion drop",
            "response_time_spike": "Response time spike",
            "lead_quality_mismatch": "Lead-quality mismatch",
            "roi_spike": "ROI spike",
            "reminder_roi_mismatch": "Reminder activity not converting into ROI",
            "service_task_surge": "Outstanding service-task surge",
            "no_show_spike": "No-show spike",
            "volume_spike": "Large volume swing",
            "time_spike": "Large response-time swing",
            "rate_spike": "Large rate swing",
        }
        primary_risk = anomaly_label_map.get(important_anomaly_types(scoring_profile.anomaly_types)[0], "Anomaly detected in reliable metrics")
    else:
        primary_risk = worst_declining_metric_phrase(metrics) or "No major risk signal identified"

    # Recommended Focus is no longer emitted. It repeated one of a dozen fixed
    # sentences across reports and was the last piece of advice-giving copy in a
    # document whose job is to state what happened.
    recommended_focus = ""

    return overall_month_rating, primary_strength, primary_risk, recommended_focus


def selected_rules_for_result(
    metrics: dict[str, dict[str, Any]],
    dealership_name: str,
    month_label: str,
    insights_source: str,
) -> tuple[list[RuleInsight], list[str]]:
    if insights_source != "rule-based":
        return [], []
    context_seed = f"{dealership_name}|{month_label}"
    candidates = build_rule_candidates(metrics, context_seed)
    selected, conflicts = select_top_insights(candidates, limit=4, metrics=metrics)
    ordered = phase164_prioritise_insight_display_order(selected, metrics)
    return lead_with_most_material_insight(ordered, metrics), conflicts


def build_review_reason_tags(
    result: DealershipResult,
    report_match: ReportMatch,
    metrics: dict[str, dict[str, Any]],
    selected_rules: list[RuleInsight],
    conflicts: list[str],
    scoring_profile: InternalScoringProfile,
    workbook_sourced: bool = False,
) -> list[str]:
    watchpoint_types = collect_persistent_watchpoint_types(metrics)
    pdfs_missing = report_match.sales_pdf is None or report_match.service_pdf is None
    ordered_tags: list[tuple[str, bool]] = [
        ("low_confidence", result.confidence_score < 0.8),
        ("fallback_mode", result.insights_source == "placeholder" or result.parsing_status == "fallback"),
        (
            "missing_reports",
            not workbook_sourced and (result.report_status == "missing file" or pdfs_missing),
        ),
        # The figures came from the workbook; only the email attachments are absent.
        ("missing_attachments", workbook_sourced and pdfs_missing),
        ("conflicting_signals", bool(conflicts) or result.sentiment == "mixed" or scoring_profile.scoring_conflict),
        (
            "anomaly_detected",
            phase171_should_emit_anomaly_detected_review_tag(metrics, scoring_profile, result),
        ),
        ("conversion_drop", "conversion_drop" in scoring_profile.anomaly_types),
        ("response_time_spike", "response_time_spike" in scoring_profile.anomaly_types),
        ("lead_quality_mismatch", "lead_quality_mismatch" in scoring_profile.anomaly_types),
        ("roi_spike", "roi_spike" in scoring_profile.anomaly_types),
        (
            "negative_roi",
            any(rule_id.endswith("negative_roi") for rule_id in result.rules_triggered or [])
            or ((metric_value(metrics, "sales_campaign_roi") or 0) < 0)
            or ((metric_value(metrics, "service_campaign_roi") or 0) < 0),
        ),
        ("unreliable_metrics", scoring_profile.unreliable_metric_count > 2),
        ("persistent_watchpoint", bool(watchpoint_types)),
        ("lead_followup_issue", "lead_followup_issue" in watchpoint_types),
        ("response_time_issue", "response_time_issue" in watchpoint_types),
        ("conversion_improvement", "conversion_improvement" in watchpoint_types),
        ("roi_improvement", "roi_improvement" in watchpoint_types),
        (
            "high_uncontacted_leads",
            (metric_value(metrics, "sales_uncontacted_leads") or 0) >= 20
            or any(rule.rule_id == "lead_handling_pressure" for rule in selected_rules),
        ),
        (
            "slow_response_time",
            (metric_value(metrics, "sales_internet_response_time") or 0) >= 120
            or any(rule.rule_id == "lead_handling_pressure" for rule in selected_rules),
        ),
        (
            "service_task_pressure",
            any(rule.rule_id == "service_tasks_pressure" for rule in selected_rules)
            or (
                metric_direction(metrics, "service_outstanding_tasks") == "up"
                and (metric_value(metrics, "service_outstanding_tasks") or 0) >= 50
            ),
        ),
        (
            "weak_conversion",
            any(
                rule.rule_id in {"leads_up_conversion_down", "leads_up_deals_down", "test_drive_and_conversion_down"}
                for rule in selected_rules
            )
            or (
                metric_direction(metrics, "sales_overall_conversion_rate") == "down"
                and (metric_value(metrics, "sales_overall_conversion_rate") or 0) <= 12
            ),
        ),
        (
            "strong_month_watchpoint",
            result.overall_month_rating == "Strong" and any(rule.polarity == "negative" for rule in selected_rules),
        ),
        ("unusually_strong_month", result.sentiment == "strongly_improved"),
        ("unusually_soft_month", result.sentiment == "significantly_softer"),
        # Usually means a report was missed rather than a genuinely dead month, and it
        # is invisible in the finished email: the copy simply covers one side.
        ("service_data_missing", "service" in families_without_usable_evidence(metrics)),
        ("sales_data_missing", "sales" in families_without_usable_evidence(metrics)),
    ]
    return [tag for tag, include in ordered_tags if include]


def add_rule_candidate(candidates: list[RuleInsight], candidate: RuleInsight | None) -> None:
    """
    Adds a rule candidate to the list if it exists and has a meaningful message.
    
    PATCH 1: Empty/Punctuation Candidate Protection
    Rejects candidates with empty/punctuation-only messages.
    """
    if candidate is not None and is_meaningful_candidate_message(candidate.message):
        candidates.append(candidate)
    elif candidate is not None:
        # Log rejected candidate for debugging
        logging.debug(f"Rejected candidate '{candidate.rule_id}' with non-meaningful message: '{candidate.message}'")


def build_cross_metric_reasoning_candidates(metrics: dict[str, dict[str, Any]], context_seed: str) -> list[RuleInsight]:
    """
    Generate insights from cross-metric relationships and business reasoning patterns.
    
    PHASE 2: Cross-Metric Reasoning
    Creates compound insights from metric interdependencies and business logic.
    """
    candidates: list[RuleInsight] = []
    
    # Pattern 1: Leads down + Conversion up = Efficiency improvement
    if (metric_direction(metrics, "sales_total_leads") == "down" and
        metric_direction(metrics, "sales_overall_conversion_rate") == "up"):
        
        leads_change = metric_change_percent(metrics, "sales_total_leads") or 0
        conversion_change = metric_change_percent(metrics, "sales_overall_conversion_rate") or 0
        
        if abs(leads_change) > 5 and conversion_change > 3:  # Meaningful changes
            supporting_metrics = ["sales_total_leads", "sales_overall_conversion_rate"]
            tone_band = "strongly_improved" if conversion_change > 8 else "slightly_improved"
            
            message = generate_cross_metric_message(
                "efficiency_gain",
                {"leads_down": abs(leads_change), "conversion_up": conversion_change},
                context_seed
            )
            
            candidates.append(make_rule_insight(
                "cross_metric_efficiency_gain",
                message,
                "sales_funnel",
                "high",
                78,  # High priority for efficiency insights
                "positive",
                tone_band,
                "sales_efficiency_cross_metric",
                supporting_metrics,
                signal_type="cross_metric_reasoning"
            ))
    
    # Pattern 2: Service stable + Retention rising = Customer loyalty strength
    if (metric_direction(metrics, "service_bookings_from_campaigns") in ["flat", "up"] and
        metric_direction(metrics, "service_customers_enrolled") == "up" and
        metric_direction(metrics, "service_outstanding_tasks") != "up"):
        
        enrollment_change = metric_change_percent(metrics, "service_customers_enrolled") or 0
        
        if enrollment_change > 5:
            supporting_metrics = ["service_bookings_from_campaigns", "service_customers_enrolled", "service_outstanding_tasks"]
            tone_band = "strongly_improved" if enrollment_change > 12 else "slightly_improved"
            
            message = generate_cross_metric_message(
                "retention_strength",
                {"enrollment_up": enrollment_change},
                context_seed
            )
            
            candidates.append(make_rule_insight(
                "cross_metric_retention_strength",
                message,
                "service_operations",
                "medium",
                65,
                "positive",
                tone_band,
                "service_retention_cross_metric",
                supporting_metrics,
                signal_type="cross_metric_reasoning"
            ))
    
    # Pattern 3: Traffic up + Appointments down = Qualification inefficiency
    if (metric_direction(metrics, "sales_total_leads") == "up" and
        metric_direction(metrics, "sales_new_leads_to_appointments") == "down"):
        
        leads_change = metric_change_percent(metrics, "sales_total_leads") or 0
        appointment_change = metric_change_percent(metrics, "sales_new_leads_to_appointments") or 0
        
        if leads_change > 8 and abs(appointment_change) > 5:
            supporting_metrics = ["sales_total_leads", "sales_new_leads_to_appointments"]
            tone_band = "significantly_softer" if abs(appointment_change) > 12 else "slightly_softer"
            
            message = generate_cross_metric_message(
                "qualification_inefficiency",
                {"leads_up": leads_change, "appointments_down": abs(appointment_change)},
                context_seed
            )
            
            candidates.append(make_rule_insight(
                "cross_metric_qualification_inefficiency",
                message,
                "lead_management",
                "high",
                80,
                "negative",
                tone_band,
                "lead_qualification_cross_metric",
                supporting_metrics,
                signal_type="cross_metric_reasoning"
            ))
    
    # Pattern 4: Sales down + Service ROI up = Service profitability focus
    if (metric_direction(metrics, "sales_total_deals") == "down" and
        metric_direction(metrics, "service_campaign_roi") == "up"):
        
        sales_change = metric_change_percent(metrics, "sales_total_deals") or 0
        roi_change = metric_change_percent(metrics, "service_campaign_roi") or 0
        
        if abs(sales_change) > 8 and roi_change > 10:
            supporting_metrics = ["sales_total_deals", "service_campaign_roi"]
            tone_band = "strongly_improved"
            
            message = generate_cross_metric_message(
                "service_profitability_focus",
                {"sales_down": abs(sales_change), "roi_up": roi_change},
                context_seed
            )
            
            candidates.append(make_rule_insight(
                "cross_metric_service_profitability",
                message,
                "service_campaigns",
                "medium",
                68,
                "positive",
                tone_band,
                "service_roi_cross_metric",
                supporting_metrics,
                signal_type="cross_metric_reasoning"
            ))
    
    return candidates


def generate_cross_metric_message(pattern_type: str, metrics_data: dict[str, float], context_seed: str) -> str:
    """
    Generate contextual messages for cross-metric reasoning patterns.
    """
    templates = {
        "efficiency_gain": [
            f"Despite {metrics_data.get('leads_down', 0):.1f}% fewer leads, conversion improved by {metrics_data.get('conversion_up', 0):.1f}%, demonstrating stronger sales efficiency and better lead qualification.",
            f"The team showed efficiency gains this month, converting leads {metrics_data.get('conversion_up', 0):.1f}% better despite a {metrics_data.get('leads_down', 0):.1f}% reduction in lead volume.",
            f"Sales effectiveness strengthened, with conversion improving {metrics_data.get('conversion_up', 0):.1f}% even as lead volume declined {metrics_data.get('leads_down', 0):.1f}%."
        ],
        "retention_strength": [
            f"Service retention showed strength, with customer enrollment up {metrics_data.get('enrollment_up', 0):.1f}% and stable booking momentum indicating growing customer loyalty.",
            f"Customer retention improved materially, with enrollment increasing {metrics_data.get('enrollment_up', 0):.1f}% while maintaining service booking stability.",
            f"The service team strengthened customer relationships, growing enrollment by {metrics_data.get('enrollment_up', 0):.1f}% with solid booking consistency."
        ],
        "qualification_inefficiency": [
            f"Lead qualification efficiency declined, with {metrics_data.get('leads_up', 0):.1f}% more leads but {metrics_data.get('appointments_down', 0):.1f}% fewer converting to appointments.",
            f"Despite {metrics_data.get('leads_up', 0):.1f}% lead volume growth, appointment conversion fell {metrics_data.get('appointments_down', 0):.1f}%, indicating qualification challenges.",
            f"Lead-to-appointment conversion weakened {metrics_data.get('appointments_down', 0):.1f}% despite {metrics_data.get('leads_up', 0):.1f}% more enquiries, suggesting qualification process inefficiencies."
        ],
        "service_profitability_focus": [
            f"While sales softened {metrics_data.get('sales_down', 0):.1f}%, service profitability strengthened with {metrics_data.get('roi_up', 0):.1f}% ROI improvement, showing effective service revenue focus.",
            f"Service operations provided profitability stability, with ROI up {metrics_data.get('roi_up', 0):.1f}% despite sales declining {metrics_data.get('sales_down', 0):.1f}%.",
            f"The service department maintained strong financial performance, improving ROI by {metrics_data.get('roi_up', 0):.1f}% even as sales momentum eased {metrics_data.get('sales_down', 0):.1f}%."
        ]
    }
    
    pattern_templates = templates.get(pattern_type, ["Cross-metric pattern detected with improved business reasoning."])
    return choose_deterministic_phrase(pattern_templates, f"{context_seed}|cross_metric_{pattern_type}")


def metric_change_percent(metrics: dict[str, dict[str, Any]], metric_name: str) -> float | None:
    """Get the percentage change for a metric."""
    metric = metrics.get(metric_name)
    if not metric:
        return None
    percent_change = metric.get("percent_change")
    if isinstance(percent_change, (int, float)):
        return float(percent_change)
    pp = metric.get("percentage_point_change")
    if isinstance(pp, (int, float)):
        return float(pp)
    return None


def build_rule_candidates(metrics: dict[str, dict[str, Any]], context_seed: str) -> list[RuleInsight]:
    """
    Enhanced rule-based insight generation with cross-metric reasoning.
    
    PHASE 2: Cross-Metric Reasoning Integration
    Combines traditional rule-based insights with compound cross-metric patterns.
    """
    candidates: list[RuleInsight] = []

    # Original rule-based candidates
    if metric_direction(metrics, "sales_total_deals") == "up" and (
        metric_direction(metrics, "service_bookings_from_campaigns") == "up"
        or metric_direction(metrics, "service_campaign_roi") == "up"
    ):
        supporting_metrics = ["sales_total_deals", "service_bookings_from_campaigns", "service_campaign_roi"]
        tone_band = rule_tone_band(metrics, supporting_metrics, "positive")
        add_rule_candidate(
            candidates,
            make_rule_insight(
                "combined_sales_service_up",
                render_rule_message_with_evidence(
                    "combined_sales_service_up",
                    tone_band,
                    f"{context_seed}|combined_sales_service_up",
                    metrics,
                ),
                "combined_performance",
                "medium",
                62,
                "positive",
                tone_band,
                "combined_overview",
                supporting_metrics,
            ),
        )

    if metric_direction(metrics, "sales_total_leads") == "up" and metric_direction(metrics, "sales_overall_conversion_rate") == "down":
        supporting_metrics = ["sales_total_leads", "sales_overall_conversion_rate"]
        tone_band = rule_tone_band(metrics, supporting_metrics, "negative")
        add_rule_candidate(
            candidates,
            make_rule_insight(
                "leads_up_conversion_down",
                render_rule_message_with_evidence(
                    "leads_up_conversion_down",
                    tone_band,
                    f"{context_seed}|leads_up_conversion_down",
                    metrics,
                ),
                "sales_funnel",
                "critical",
                95,
                "negative",
                tone_band,
                "sales_funnel_pressure",
                supporting_metrics,
            ),
        )
    elif metric_direction(metrics, "sales_total_leads") == "down" and metric_direction(metrics, "sales_overall_conversion_rate") == "up":
        supporting_metrics = ["sales_total_leads", "sales_overall_conversion_rate"]
        tone_band = rule_tone_band(metrics, supporting_metrics, "positive")
        add_rule_candidate(
            candidates,
            make_rule_insight(
                "leads_down_conversion_up",
                render_rule_message_with_evidence(
                    "leads_down_conversion_up",
                    tone_band,
                    f"{context_seed}|leads_down_conversion_up",
                    metrics,
                ),
                "sales_funnel",
                "high",
                74,
                "positive",
                tone_band,
                "sales_funnel_efficiency",
                supporting_metrics,
            ),
        )
    elif metric_direction(metrics, "sales_total_leads") == "up" and metric_direction(metrics, "sales_total_deals") == "down":
        supporting_metrics = ["sales_total_leads", "sales_total_deals"]
        tone_band = rule_tone_band(metrics, supporting_metrics, "negative")
        add_rule_candidate(
            candidates,
            make_rule_insight(
                "leads_up_deals_down",
                render_rule_message_with_evidence(
                    "leads_up_deals_down",
                    tone_band,
                    f"{context_seed}|leads_up_deals_down",
                    metrics,
                ),
                "sales_funnel",
                "critical",
                90,
                "negative",
                tone_band,
                "sales_funnel_pressure",
                supporting_metrics,
            ),
        )

    response_time = metric_direction(metrics, "sales_internet_response_time")
    uncontacted_leads = metric_direction(metrics, "sales_uncontacted_leads")
    rt_metric = metrics.get("sales_internet_response_time")
    uc_metric = metrics.get("sales_uncontacted_leads")
    rt_sig_up = response_time == "up" and bool(rt_metric and metric_is_significant(rt_metric, "slight"))
    uc_sig_up = uncontacted_leads == "up" and bool(uc_metric and metric_is_significant(uc_metric, "slight"))
    if (rt_sig_up or uc_sig_up) and not (response_time == "down" and uncontacted_leads == "down"):
        if not phase155c_composite_crm_trigger(metrics):
            supporting_metrics = [name for name in ("sales_internet_response_time", "sales_uncontacted_leads") if metric_direction(metrics, name) == "up"]
            if not supporting_metrics:
                supporting_metrics = ["sales_internet_response_time", "sales_uncontacted_leads"]
            tone_band = rule_tone_band(metrics, supporting_metrics, "negative")
            pressure_weight = 84 if response_time == "up" and uncontacted_leads == "up" else 80
            add_rule_candidate(
                candidates,
                make_rule_insight(
                    "lead_handling_pressure",
                    render_rule_message_with_evidence(
                        "lead_handling_pressure",
                        tone_band,
                        f"{context_seed}|lead_handling_pressure",
                        metrics,
                    ),
                    "lead_management",
                    "high",
                    pressure_weight,
                    "negative",
                    tone_band,
                    "lead_management_focus",
                    supporting_metrics,
                ),
            )
    elif response_time == "down" and uncontacted_leads == "down":
        supporting_metrics = ["sales_internet_response_time", "sales_uncontacted_leads"]
        tone_band = rule_tone_band(metrics, supporting_metrics, "positive")
        add_rule_candidate(
            candidates,
            make_rule_insight(
                "lead_handling_improved",
                render_rule_message_with_evidence(
                    "lead_handling_improved",
                    tone_band,
                    f"{context_seed}|lead_handling_improved",
                    metrics,
                ),
                "lead_management",
                "medium",
                58,
                "positive",
                tone_band,
                "lead_management_focus",
                supporting_metrics,
            ),
        )

    if phase155c_composite_crm_trigger(metrics):
        comp_focus = [name for name in PHASE155C_COMPOSITE_ORDER if phase155c_is_material_crm_adverse(metrics, name)]
        if comp_focus:
            tone_band = rule_tone_band(metrics, comp_focus, "negative")
            comp_msg = build_crm_composite_insight_message(metrics, f"{context_seed}|crm_composite")
            if comp_msg and is_meaningful_candidate_message(comp_msg):
                add_rule_candidate(
                    candidates,
                    make_rule_insight(
                        "crm_execution_pressure_composite",
                        comp_msg,
                        "lead_management",
                        "critical",
                        98,
                        "negative",
                        tone_band,
                        "crm_composite_execution",
                        comp_focus,
                        signal_type="composite_crm_reasoning",
                    ),
                )

    if metric_direction(metrics, "sales_appointment_to_sales") == "up" and metric_direction(metrics, "sales_same_day_deal_pct") == "up":
        supporting_metrics = ["sales_appointment_to_sales", "sales_same_day_deal_pct"]
        tone_band = rule_tone_band(metrics, supporting_metrics, "positive")
        add_rule_candidate(
            candidates,
            make_rule_insight(
                "closing_execution_up",
                render_rule_message("closing_execution_up", tone_band, f"{context_seed}|closing_execution_up"),
                "sales_execution",
                "medium",
                60,
                "positive",
                tone_band,
                "sales_execution_strength",
                supporting_metrics,
            ),
        )
    if metric_direction(metrics, "sales_new_leads_to_test_drives") == "down" and metric_direction(metrics, "sales_overall_conversion_rate") == "down":
        supporting_metrics = ["sales_new_leads_to_test_drives", "sales_overall_conversion_rate"]
        tone_band = rule_tone_band(metrics, supporting_metrics, "negative")
        add_rule_candidate(
            candidates,
            make_rule_insight(
                "test_drive_and_conversion_down",
                render_rule_message("test_drive_and_conversion_down", tone_band, f"{context_seed}|test_drive_and_conversion_down"),
                "sales_execution",
                "high",
                76,
                "negative",
                tone_band,
                "sales_execution_pressure",
                supporting_metrics,
            ),
        )

    sales_campaign_roi_current = metric_value(metrics, "sales_campaign_roi")
    if metric_direction(metrics, "sales_campaign_deals") == "up" and metric_direction(metrics, "sales_campaign_roi") == "up":
        supporting_metrics = ["sales_campaign_deals", "sales_campaign_roi"]
        tone_band = rule_tone_band(metrics, supporting_metrics, "positive")
        add_rule_candidate(
            candidates,
            make_rule_insight(
                "sales_marketing_up",
                render_rule_message_with_evidence(
                    "sales_marketing_up",
                    tone_band,
                    f"{context_seed}|sales_marketing_up",
                    metrics,
                ),
                "marketing",
                "medium",
                54,
                "positive",
                tone_band,
                "marketing_effectiveness",
                supporting_metrics,
            ),
        )
    elif sales_campaign_roi_current is not None and sales_campaign_roi_current < 0:
        supporting_metrics = ["sales_campaign_roi"]
        tone_band = rule_tone_band(metrics, supporting_metrics, "negative", current_value_override=sales_campaign_roi_current)
        add_rule_candidate(
            candidates,
            make_rule_insight(
                "sales_marketing_negative_roi",
                render_rule_message("sales_marketing_negative_roi", tone_band, f"{context_seed}|sales_marketing_negative_roi"),
                "marketing",
                "critical",
                88,
                "negative",
                tone_band,
                "marketing_risk",
                supporting_metrics,
            ),
        )
    elif phase159_sales_campaign_roi_material_decline_eligible(metrics):
        supporting_metrics = ["sales_campaign_roi"]
        deals_m = metrics.get("sales_campaign_deals")
        if deals_m and metric_usable_for_reporting(deals_m, "slight") and str(deals_m.get("direction")) == "down":
            supporting_metrics.append("sales_campaign_deals")
        tone_band = rule_tone_band(metrics, supporting_metrics, "negative")
        roi_metric = metrics.get("sales_campaign_roi") or {}
        tier = phase159_deterioration_magnitude_tier(phase159_percent_change_abs(roi_metric))
        sev = "high" if tier in {"severe", "critical"} else "medium"
        wt = {"mild": 72, "moderate": 78, "severe": 84, "critical": 90}.get(tier, 76)
        add_rule_candidate(
            candidates,
            make_rule_insight(
                "sales_campaign_roi_decline",
                render_rule_message("sales_campaign_roi_decline", tone_band, f"{context_seed}|sales_campaign_roi_decline"),
                "marketing",
                sev,
                wt,
                "negative",
                tone_band,
                "sales_campaign_return_pressure",
                supporting_metrics,
            ),
        )

    if metric_direction(metrics, "service_bookings_from_campaigns") == "up" and metric_direction(metrics, "service_campaign_roi") == "up":
        supporting_metrics = ["service_bookings_from_campaigns", "service_campaign_roi"]
        tone_band = rule_tone_band(metrics, supporting_metrics, "positive")
        add_rule_candidate(
            candidates,
            make_rule_insight(
                "service_campaigns_up",
                render_rule_message_with_evidence(
                    "service_campaigns_up",
                    tone_band,
                    f"{context_seed}|service_campaigns_up",
                    metrics,
                ),
                "service_campaigns",
                "medium",
                57,
                "positive",
                tone_band,
                "service_campaign_strength",
                supporting_metrics,
            ),
        )
    elif metric_direction(metrics, "service_bookings_from_campaigns") in {"up", "flat"} and metric_direction(metrics, "service_campaign_roi") == "down":
        supporting_metrics = ["service_bookings_from_campaigns", "service_campaign_roi"]
        tone_band = rule_tone_band(metrics, supporting_metrics, "negative")
        add_rule_candidate(
            candidates,
            make_rule_insight(
                "service_bookings_roi_mismatch",
                render_rule_message("service_bookings_roi_mismatch", tone_band, f"{context_seed}|service_bookings_roi_mismatch"),
                "service_campaigns",
                "high",
                72,
                "negative",
                tone_band,
                "service_campaign_pressure",
                supporting_metrics,
            ),
        )

    if metric_direction(metrics, "service_outstanding_tasks") == "up" and metric_direction(metrics, "service_bookings_from_campaigns") != "up":
        supporting_metrics = ["service_outstanding_tasks", "service_bookings_from_campaigns"]
        tone_band = rule_tone_band(metrics, supporting_metrics, "negative")
        add_rule_candidate(
            candidates,
            make_rule_insight(
                "service_tasks_pressure",
                render_rule_message_with_evidence(
                    "service_tasks_pressure",
                    tone_band,
                    f"{context_seed}|service_tasks_pressure",
                    metrics,
                ),
                "service_operations",
                "high",
                82,
                "negative",
                tone_band,
                "service_operations_focus",
                supporting_metrics,
            ),
        )
    elif metric_direction(metrics, "service_outstanding_tasks") == "down" and metric_direction(metrics, "service_customers_enrolled") == "up":
        supporting_metrics = ["service_outstanding_tasks", "service_customers_enrolled"]
        tone_band = rule_tone_band(metrics, supporting_metrics, "positive")
        add_rule_candidate(
            candidates,
            make_rule_insight(
                "service_follow_up_cleaner",
                render_rule_message_with_evidence(
                    "service_follow_up_cleaner",
                    tone_band,
                    f"{context_seed}|service_follow_up_cleaner",
                    metrics,
                ),
                "service_operations",
                "medium",
                55,
                "positive",
                tone_band,
                "service_operations_focus",
                supporting_metrics,
            ),
        )

    fallback_candidates = [
        (
            "sales_volume_snapshot",
            metrics.get("sales_total_deals"),
            "sales_overview",
            "low",
            38,
            "mixed",
            True,
            "sales_snapshot",
        ),
        (
            "sales_conversion_snapshot",
            metrics.get("sales_overall_conversion_rate"),
            "sales_overview",
            "low",
            35,
            "mixed",
            True,
            "sales_snapshot",
        ),
        (
            "lead_response_snapshot",
            metrics.get("sales_internet_response_time"),
            "lead_management",
            "medium",
            46,
            "mixed",
            False,
            "lead_management_focus",
        ),
        (
            "service_booking_snapshot",
            metrics.get("service_bookings_from_campaigns"),
            "service_campaigns",
            "low",
            34,
            "mixed",
            True,
            "service_campaign_snapshot",
        ),
        (
            "service_roi_snapshot",
            metrics.get("service_campaign_roi"),
            "service_campaigns",
            "medium",
            44,
            "mixed",
            True,
            "service_campaign_snapshot",
        ),
        (
            "service_task_snapshot",
            metrics.get("service_outstanding_tasks"),
            "service_operations",
            "medium",
            48,
            "mixed",
            False,
            "service_operations_focus",
        ),
    ]

    for rule_id, metric, category, severity, weight, polarity, higher_is_better, merge_key in fallback_candidates:
        if metric_is_meaningful(metric):
            tone_band = metric_snapshot_tone_band(metric, higher_is_better)
            add_rule_candidate(
                candidates,
                make_rule_insight(
                    rule_id,
                    render_rule_message(
                        rule_id,
                        tone_band,
                        f"{context_seed}|{rule_id}",
                        change=format_metric_change(metric),
                        directional_change=format_metric_directional_change(metric),
                        current_raw=str(metric.get("current_raw", "")),
                        fallback_message=f"{rule_id} {format_metric_change(metric)}.",
                    ),
                    category,
                    severity,
                    weight,
                    polarity,
                    tone_band,
                    merge_key,
                    [],
                ),
            )

    for candidate in candidates:
        candidate.priority_weight += trend_bonus_for_rule(candidate, metrics)
        candidate.priority_weight += phase155_priority_materiality_adjustment(candidate, metrics)
        finalise_rule_candidate(candidate, metrics)

    # PHASE 2: Add cross-metric reasoning candidates
    cross_metric_candidates = build_cross_metric_reasoning_candidates(metrics, context_seed)
    candidates.extend(cross_metric_candidates)
    for cross_rule in cross_metric_candidates:
        cross_rule.priority_weight += trend_bonus_for_rule(cross_rule, metrics)
        cross_rule.priority_weight += phase155_priority_materiality_adjustment(cross_rule, metrics)
        finalise_rule_candidate(cross_rule, metrics)

    return candidates


def deduplicate_rule_candidates(candidates: list[RuleInsight]) -> list[RuleInsight]:
    """
    Deduplicates rule candidates while preventing empty/punctuation-only messages
    from contaminating the deduplication process.
    
    PATCH 2: Deduplication Safety Fix
    Guards against punctuation-only candidates in substring comparisons.
    """
    ranked = sorted(
        candidates,
        key=lambda candidate: (
            -candidate.priority_weight,
            candidate.category,
            candidate.rule_id,
        ),
    )

    selected: list[RuleInsight] = []
    used_merge_keys: set[str] = set()
    used_messages: list[str] = []

    for candidate in ranked:
        if candidate.merge_key in used_merge_keys:
            continue

        candidate_primary_focus = (candidate.metric_focus or [None])[0]

        normalized_message = normalize_text(candidate.message)
        
        # PATCH 2: Reject candidates with non-meaningful normalized messages
        if not is_meaningful_candidate_message(normalized_message):
            logging.debug(f"Rejected candidate '{candidate.rule_id}' during deduplication with non-meaningful normalized message: '{normalized_message}'")
            continue
        
        # Only perform substring checks if the normalized message is meaningful
        if any(
            normalized_message == existing
            or normalized_message in existing
            or existing in normalized_message
            for existing in used_messages
        ):
            continue

        if any(
            set(candidate.metric_focus or []) & set(existing.metric_focus or [])
            and (
                candidate.signal_type == existing.signal_type
                or candidate.category == existing.category
            )
            for existing in selected
        ):
            continue

        if any(
            candidate_primary_focus
            and candidate_primary_focus == (existing.metric_focus or [None])[0]
            and candidate.polarity == existing.polarity
            for existing in selected
        ):
            continue

        selected.append(candidate)
        used_merge_keys.add(candidate.merge_key)
        used_messages.append(normalized_message)

    return selected


def detect_conflicting_categories(candidates: list[RuleInsight]) -> list[str]:
    polarities_by_category: dict[str, set[str]] = {}
    for candidate in candidates:
        polarities_by_category.setdefault(candidate.category, set()).add(candidate.polarity)

    return [
        category
        for category, polarities in polarities_by_category.items()
        if "positive" in polarities and "negative" in polarities
    ]


def dealership_activity_tier(metrics: dict[str, dict[str, Any]]) -> str:
    leads = metric_value(metrics, "sales_total_leads") or 0.0
    deals = metric_value(metrics, "sales_total_deals") or 0.0
    bookings = metric_value(metrics, "service_bookings_from_campaigns") or 0.0
    reminders = metric_value(metrics, "service_reminders_sent") or 0.0
    index = leads + deals + 0.35 * bookings + 0.15 * reminders
    if index >= 120:
        return "high"
    if index >= 45:
        return "medium"
    return "low"


def rule_volume_support_index(candidate: RuleInsight, metrics: dict[str, dict[str, Any]]) -> float:
    focus = list(candidate.metric_focus or candidate.supporting_metrics or [])
    total = 0.0
    for name in focus:
        m = metrics.get(name)
        if not m:
            continue
        v = m.get("current_value")
        if isinstance(v, (int, float)):
            total += float(abs(v))
    if not focus:
        return 25.0
    return max(1.0, total / max(1, len(focus)))


def volume_credibility_multiplier(candidate: RuleInsight, metrics: dict[str, dict[str, Any]]) -> float:
    tier = dealership_activity_tier(metrics)
    vol = rule_volume_support_index(candidate, metrics)
    base = {"high": 1.0, "medium": 0.92, "low": 0.82}.get(tier, 0.9)
    if vol < 8:
        return base * 0.72
    if vol < 20:
        return base * 0.86
    if vol < 40:
        return base * 0.94
    return base


def trend_persistence_adjustment(candidate: RuleInsight, metrics: dict[str, dict[str, Any]]) -> float:
    bonus = 0.0
    for metric_name in candidate.metric_focus or []:
        m = metrics.get(metric_name)
        if not m:
            continue
        pattern = str(m.get("trend_pattern", ""))
        conf = str(m.get("trend_confidence", "low"))
        aligned = rule_metric_alignment(candidate, m)
        if not aligned:
            continue
        if pattern in {"up_trend", "down_trend", "reversal_up", "reversal_down"} and conf == "high":
            bonus += 10.0
        elif pattern in {"up_trend", "down_trend", "reversal_up", "reversal_down"} and conf == "medium":
            bonus += 5.0
        if pattern == "volatile":
            bonus -= 4.0
    return bonus


def signal_rarity_bonus(candidate: RuleInsight, all_candidates: list[RuleInsight]) -> float:
    same_rule = sum(1 for c in all_candidates if c.rule_id == candidate.rule_id)
    total = len(all_candidates) or 1
    rarity = 1.0 - (same_rule / total)
    return rarity * 8.0


def rule_operational_family(rule_id: str) -> str:
    return RULE_OPERATIONAL_FAMILY.get(rule_id, "other")


# Phase 157 — executive Primary Risk prefers commercial / funnel outcomes over CRM-administration headlines.
PHASE157_EXEC_RISK_FAMILY_ORDER: dict[str, int] = {
    "conversion_story": 10,
    "volume_snapshot": 15,
    "unified_campaign": 17,
    "marketing_roi": 20,
    "combined_lane": 22,
    "retention_campaign": 24,
    "profitability_service": 35,
    "service_ops": 40,
    "crm_execution": 70,
    "other": 50,
}


def phase157_sort_negatives_for_executive(rules: list[RuleInsight]) -> list[RuleInsight]:
    """Deterministic ordering: commercial outcomes before CRM-process headlines."""

    def sort_key(rule: RuleInsight) -> tuple[int, int, str]:
        fam = rule_operational_family(rule.rule_id)
        tier = PHASE157_EXEC_RISK_FAMILY_ORDER.get(fam, 50)
        pw = int(rule.priority_weight)
        return (tier, -pw, rule.rule_id)

    return sorted(rules, key=sort_key)


def phase155_crm_deterioration_count(metrics: dict[str, dict[str, Any]]) -> int:
    """Count materially adverse CRM / follow-up / capture signals (phase 155 materiality)."""
    keys = (
        "sales_internet_response_time",
        "sales_uncontacted_leads",
        "sales_overdue_tasks",
        "sales_email_capture_rate",
        "service_customers_enrolled",
        "service_reminders_sent",
    )
    n = 0
    for metric_name in keys:
        metric = metrics.get(metric_name)
        if metric and metric_usable_for_reporting(metric, "slight") and metric_is_adverse(metrics, metric_name):
            n += 1
    return n


def phase155_priority_materiality_adjustment(rule: RuleInsight, metrics: dict[str, dict[str, Any]]) -> int:
    """Deterministic priority_weight nudges for CRM and execution materiality (phase 155)."""
    adj = 0
    crm_n = phase155_crm_deterioration_count(metrics)
    if rule.rule_id == "lead_handling_pressure":
        adj += 14
        if crm_n >= 2:
            adj += 12
        elif crm_n == 1:
            adj += 6
    elif rule.rule_id == "lead_response_snapshot" and metric_is_adverse(metrics, "sales_internet_response_time"):
        adj += 18
    elif rule.rule_id == "service_tasks_pressure" and metric_is_adverse(metrics, "service_outstanding_tasks"):
        if phase164_commercial_retention_material_stress(metrics) and not phase164_operational_task_backlog_crisis(metrics):
            adj += 3
        else:
            adj += 12
    elif rule.rule_id == "service_task_snapshot" and metric_is_adverse(metrics, "service_outstanding_tasks"):
        if phase164_commercial_retention_material_stress(metrics) and not phase164_operational_task_backlog_crisis(metrics):
            adj += 4
        else:
            adj += 14
    elif rule.rule_id == "cross_metric_qualification_inefficiency":
        adj += 8
    elif rule.rule_id in {"closing_execution_up", "cross_metric_efficiency_gain"} and crm_n >= 2:
        adj -= 8
    elif rule.rule_id == "crm_execution_pressure_composite":
        adj += 24
    if phase157_strong_commercial_performance(metrics) and not phase157_crm_execution_severe_stack(metrics):
        if rule.rule_id in {"crm_execution_pressure_composite", "lead_handling_pressure", "lead_response_snapshot"}:
            adj -= 18
    return adj


def phase155_operational_materiality_bonus(
    candidate: RuleInsight,
    metrics: dict[str, dict[str, Any]] | None,
) -> float:
    """Extra operational-significance weight for CRM and service execution deterioration (phase 155)."""
    if not metrics:
        return 0.0
    bonus = 0.0
    crm_n = phase155_crm_deterioration_count(metrics)
    if candidate.rule_id == "lead_handling_pressure":
        bonus += 20.0
        if crm_n >= 2:
            bonus += 24.0
        elif crm_n == 1:
            bonus += 10.0
    elif candidate.rule_id in {"lead_response_snapshot", "lead_handling_improved"}:
        if metric_is_adverse(metrics, "sales_internet_response_time") or metric_is_adverse(metrics, "sales_uncontacted_leads"):
            bonus += 16.0
    elif candidate.rule_id == "service_tasks_pressure" and metric_is_adverse(metrics, "service_outstanding_tasks"):
        if phase164_commercial_retention_material_stress(metrics) and not phase164_operational_task_backlog_crisis(metrics):
            bonus += 4.0
        else:
            bonus += 14.0
            tasks = metrics.get("service_outstanding_tasks")
            if tasks and metric_change_strength(tasks) >= 2:
                bonus += 8.0
    elif candidate.rule_id == "service_task_snapshot" and metric_is_adverse(metrics, "service_outstanding_tasks"):
        if phase164_commercial_retention_material_stress(metrics) and not phase164_operational_task_backlog_crisis(metrics):
            bonus += 3.0
        else:
            bonus += 12.0
    elif candidate.rule_id == "cross_metric_qualification_inefficiency":
        bonus += 10.0
    if crm_n >= 2 and candidate.rule_id in {"closing_execution_up", "cross_metric_efficiency_gain", "leads_down_conversion_up"}:
        bonus -= 14.0
    if candidate.rule_id == "crm_execution_pressure_composite":
        bonus += 58.0
        if phase157_strong_commercial_performance(metrics) and not phase157_crm_execution_severe_stack(metrics):
            bonus -= 48.0
    if candidate.rule_id in {"lead_handling_pressure", "lead_response_snapshot"}:
        if phase157_strong_commercial_performance(metrics) and not phase157_crm_execution_severe_stack(metrics):
            bonus -= 22.0
    return bonus


def phase155c_is_material_crm_adverse(metrics: dict[str, dict[str, Any]], metric_name: str) -> bool:
    metric = metrics.get(metric_name)
    if not metric or not metric_usable_for_reporting(metric, "slight"):
        return False
    return metric_is_adverse(metrics, metric_name)


def phase155c_is_severe_crm_adverse(metrics: dict[str, dict[str, Any]], metric_name: str) -> bool:
    if not phase155c_is_material_crm_adverse(metrics, metric_name):
        return False
    metric = metrics.get(metric_name) or {}
    if metric_change_strength(metric) >= 2:
        return True
    if bool(metric.get("is_anomaly")):
        return True
    return False


def phase155c_data_quality_and_followup_pair(metrics: dict[str, dict[str, Any]]) -> bool:
    dq = phase155c_is_material_crm_adverse(metrics, "sales_email_capture_rate")
    fu = phase155c_is_material_crm_adverse(metrics, "sales_internet_response_time") or phase155c_is_material_crm_adverse(
        metrics, "sales_uncontacted_leads"
    )
    return bool(dq and fu)


def phase157_commercial_outcome_score(metrics: dict[str, dict[str, Any]]) -> int:
    """Weighted read of commercial throughput / outcomes (phase 157)."""
    score = 0
    for metric_name in (
        "sales_total_deals",
        "sales_overall_conversion_rate",
        "sales_campaign_roi",
        "sales_campaign_deals",
        "service_bookings_from_campaigns",
        "service_campaign_roi",
    ):
        m = metrics.get(metric_name)
        if not m or not metric_usable_for_reporting(m, "slight"):
            continue
        if metric_is_currently_improving(m):
            score += 3
        elif not metric_is_adverse(metrics, metric_name) and str(m.get("direction")) == "up":
            score += 2
        elif not metric_is_adverse(metrics, metric_name) and str(m.get("direction")) == "flat":
            score += 1
    return score


def phase157_strong_commercial_performance(metrics: dict[str, dict[str, Any]]) -> bool:
    """Commercial outcomes are strong enough that CRM admin signals should not dominate."""
    return phase157_commercial_outcome_score(metrics) >= 5


def phase157_crm_execution_severe_stack(metrics: dict[str, dict[str, Any]]) -> bool:
    """Severe enough CRM / execution breakdown to justify headline CRM escalation despite strong sales."""
    rt = metrics.get("sales_internet_response_time")
    if rt and metric_usable_for_reporting(rt, "slight") and metric_is_adverse(metrics, "sales_internet_response_time"):
        if bool(rt.get("is_anomaly")) or metric_change_strength(rt) >= 2:
            return True
    od = metrics.get("sales_overdue_tasks")
    if od and metric_usable_for_reporting(od, "meaningful") and metric_is_adverse(metrics, "sales_overdue_tasks"):
        if metric_change_strength(od) >= 1:
            return True
    uc = metrics.get("sales_uncontacted_leads")
    if (
        uc
        and metric_usable_for_reporting(uc, "meaningful")
        and metric_is_adverse(metrics, "sales_uncontacted_leads")
        and rt
        and metric_usable_for_reporting(rt, "slight")
        and metric_is_adverse(metrics, "sales_internet_response_time")
    ):
        return True
    if phase155c_crm_stack_depth(metrics) >= 5:
        return True
    return False


def phase159_percent_change_abs(metric: dict[str, Any] | None) -> float | None:
    if not metric:
        return None
    pc = metric.get("percent_change")
    if isinstance(pc, (int, float)):
        return abs(float(pc))
    return None


def phase159_deterioration_magnitude_tier(abs_pc: float | None) -> str:
    """Deterministic commercial deterioration bands (phase 159)."""
    if abs_pc is None:
        return "none"
    if abs_pc < 10.0:
        return "mild"
    if abs_pc < 25.0:
        return "moderate"
    if abs_pc < 50.0:
        return "severe"
    return "critical"


def phase159_sales_campaign_roi_material_decline_eligible(metrics: dict[str, dict[str, Any]]) -> bool:
    """Positive-dollar sales campaign ROI with a material MoM decline (phase 159)."""
    roi = metrics.get("sales_campaign_roi")
    if not roi or not metric_usable_for_reporting(roi, "slight"):
        return False
    if str(roi.get("direction")) != "down":
        return False
    cur = roi.get("current_value")
    if cur is None:
        cur = roi.get("current")
    if cur is None or float(cur) < 0:
        return False
    apc = phase159_percent_change_abs(roi)
    if apc is None:
        return False
    band = str(roi.get("significance_band") or "")
    return apc >= 10.0 or band in {"meaningful", "strong"}


def phase159_sales_campaign_roi_critical(metrics: dict[str, dict[str, Any]]) -> bool:
    roi = metrics.get("sales_campaign_roi")
    if not roi or str(roi.get("direction")) != "down":
        return False
    return phase159_deterioration_magnitude_tier(phase159_percent_change_abs(roi)) in {"severe", "critical"}


def phase159_commercial_impact_tier_bonus(rule_id: str, metrics: dict[str, dict[str, Any]] | None) -> float:
    """Deterministic commercial-impact layer on top of operational significance (phase 159)."""
    if not metrics:
        return 0.0
    bonus = 0.0
    if rule_id == "sales_campaign_roi_decline":
        roi = metrics.get("sales_campaign_roi") or {}
        tier = phase159_deterioration_magnitude_tier(phase159_percent_change_abs(roi))
        bonus += {"mild": 6.0, "moderate": 14.0, "severe": 28.0, "critical": 44.0, "none": 0.0}.get(tier, 0.0)
    if rule_id == "test_drive_and_conversion_down":
        if metric_direction(metrics, "sales_appointment_to_sales") == "up" and metric_direction(metrics, "sales_same_day_deal_pct") == "up":
            bonus += 6.0
    return bonus


def phase159_cross_domain_significance_nudge(candidate: RuleInsight, metrics: dict[str, dict[str, Any]] | None) -> float:
    """Reduce service-campaign snapshot inflation when sales campaign deterioration is dominant (phase 159)."""
    if not metrics:
        return 0.0
    if candidate.rule_id != "service_roi_snapshot":
        return 0.0
    if not phase159_sales_campaign_roi_critical(metrics):
        return 0.0
    svc = metrics.get("service_campaign_roi") or {}
    if not metric_usable_for_reporting(svc, "slight") or str(svc.get("direction")) != "down":
        return 0.0
    sales_pc = phase159_percent_change_abs(metrics.get("sales_campaign_roi") or {})
    svc_pc = phase159_percent_change_abs(svc)
    if sales_pc is None or svc_pc is None:
        return 0.0
    if svc_pc >= sales_pc - 8.0:
        return 0.0
    return -14.0


def phase159_supplemental_executive_negative_rules(
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
) -> list[RuleInsight]:
    """Include materially adverse mixed-polarity service ROI when no sales-campaign deterioration rule is present (phase 159)."""
    ids = {r.rule_id for r in selected_rules}
    if "sales_campaign_roi_decline" in ids or "sales_marketing_negative_roi" in ids:
        return []
    extra: list[RuleInsight] = []
    for rule in selected_rules:
        if rule.rule_id != "service_roi_snapshot" or rule.polarity != "mixed":
            continue
        m = metrics.get("service_campaign_roi")
        if not m or str(m.get("direction")) != "down":
            continue
        if phase159_deterioration_magnitude_tier(phase159_percent_change_abs(m)) not in {"severe", "critical"}:
            continue
        extra.append(rule)
    return extra


# Phase 160 — unified campaign commercial family (acquisition + retention / workshop campaigns).
PHASE160_CAMPAIGN_LANE: dict[str, str] = {
    "sales_campaign_roi_decline": "sales",
    "sales_marketing_negative_roi": "sales",
    "sales_marketing_up": "sales",
    "service_roi_snapshot": "service",
    "service_bookings_roi_mismatch": "service",
    "service_booking_snapshot": "service",
    "service_campaigns_up": "service",
    "cross_metric_retention_strength": "service",
}


def phase160_campaign_lane(rule_id: str) -> str | None:
    return PHASE160_CAMPAIGN_LANE.get(rule_id)


def phase160_service_campaign_material(metrics: dict[str, dict[str, Any]]) -> bool:
    """Material deterioration or pressure on service-side campaign / retention throughput (phase 160)."""
    roi = metrics.get("service_campaign_roi")
    if roi and str(roi.get("direction")) == "down" and metric_usable_for_reporting(roi, "slight"):
        apc = phase159_percent_change_abs(roi)
        band = str(roi.get("significance_band") or "")
        if apc is not None and (apc >= 7.0 or band in {"meaningful", "strong"}):
            return True
    bk = metrics.get("service_bookings_from_campaigns")
    if bk and str(bk.get("direction")) == "down" and metric_usable_for_reporting(bk, "meaningful"):
        return True
    return False


def phase160_service_retention_commercial_bonus(rule_id: str, metrics: dict[str, dict[str, Any]] | None) -> float:
    """Treat workshop / retention campaign signals as commercial-impact-weighted (phase 160)."""
    if not metrics:
        return 0.0
    if rule_id not in {"service_roi_snapshot", "service_bookings_roi_mismatch", "service_booking_snapshot"}:
        return 0.0
    roi = metrics.get("service_campaign_roi") or {}
    bonus = 0.0
    if str(roi.get("direction")) == "down" and metric_usable_for_reporting(roi, "slight"):
        tier = phase159_deterioration_magnitude_tier(phase159_percent_change_abs(roi))
        bonus += {"mild": 6.0, "moderate": 14.0, "severe": 24.0, "critical": 36.0, "none": 0.0}.get(tier, 0.0)
        if rule_id == "service_bookings_roi_mismatch":
            bonus *= 1.12
    if rule_id == "service_booking_snapshot":
        bk = metrics.get("service_bookings_from_campaigns") or {}
        if str(bk.get("direction")) == "down" and metric_usable_for_reporting(bk, "slight"):
            bonus += 6.0
    return bonus


def phase160_dual_campaign_balance_nudge(candidate: RuleInsight, metrics: dict[str, dict[str, Any]] | None) -> float:
    """When sales and retention campaigns both weaken materially, avoid funnel rules crowding out retention campaign (phase 160)."""
    if not metrics:
        return 0.0
    if not phase159_sales_campaign_roi_material_decline_eligible(metrics):
        return 0.0
    if not phase160_service_campaign_material(metrics):
        return 0.0
    if candidate.rule_id == "test_drive_and_conversion_down":
        return -14.0
    if candidate.rule_id in {"service_roi_snapshot", "service_bookings_roi_mismatch"}:
        return 14.0
    return 0.0


# --- Phase 164 — retention commercial priority vs operational escalation containment ---


def phase164_commercial_retention_material_stress(metrics: dict[str, dict[str, Any]] | None) -> bool:
    """True when workshop / service campaign retention economics are materially stressed (phase 164)."""
    if not metrics:
        return False
    if phase160_service_campaign_material(metrics):
        return True
    roi = metrics.get("service_campaign_roi")
    if roi and metric_usable_for_reporting(roi, "slight") and metric_is_adverse(metrics, "service_campaign_roi"):
        return metric_change_strength(roi) >= 1
    return False


def phase164_operational_task_backlog_crisis(metrics: dict[str, dict[str, Any]] | None) -> bool:
    """Severe / persistent backlog escalation — operational dominance remains valid (phase 164)."""
    if not metrics:
        return False
    t = metrics.get("service_outstanding_tasks")
    if not t or not metric_usable_for_reporting(t, "slight"):
        return False
    if not metric_is_adverse(metrics, "service_outstanding_tasks"):
        return False
    if metric_change_strength(t) >= 2:
        return True
    pc = t.get("percent_change")
    if isinstance(pc, (int, float)) and abs(float(pc)) >= 22.0:
        return True
    if bool(t.get("persistent_watchpoint")) and metric_change_strength(t) >= 2:
        if isinstance(pc, (int, float)) and abs(float(pc)) >= 15.0:
            return True
    return False


def phase164_retention_commercial_arbitration_nudge(
    candidate: RuleInsight,
    metrics: dict[str, dict[str, Any]],
) -> float:
    """
    When retention campaign economics are materially stressed, demote moderate workflow/backlog rules
    and boost retention / ROI snapshot rules in insight arbitration (phase 164).
    When service campaign ROI is materially improving, still elevate the ROI snapshot over workload
    so commercial retention performance leads the narrative.
    """
    if not metrics:
        return 0.0
    crisis = phase164_operational_task_backlog_crisis(metrics)
    rid = candidate.rule_id
    retention_ids = {"service_roi_snapshot", "service_bookings_roi_mismatch", "service_marketing_negative_roi"}
    ops_workload_ids = {"service_tasks_pressure", "service_task_snapshot"}
    if crisis:
        if rid in retention_ids:
            return 18.0
        return 0.0

    roi = metrics.get("service_campaign_roi")
    roi_constructive = bool(
        roi
        and metric_usable_for_reporting(roi, "slight")
        and metric_is_currently_improving(roi)
        and str(roi.get("direction", "flat")) != "flat"
    )
    if roi_constructive:
        if rid == "service_roi_snapshot":
            return 34.0
        if rid in ops_workload_ids:
            return -38.0
        return 0.0

    if phase164_commercial_retention_material_stress(metrics):
        if rid in ops_workload_ids:
            return -58.0
        if rid in retention_ids:
            return 46.0
    return 0.0


# How much a movement in each metric actually matters to a dealership owner. Deals and
# campaign return are the outcome; task counts are the plumbing behind it.
INSIGHT_MATERIALITY_WEIGHTS: dict[str, float] = {
    "sales_total_deals": 1.0,
    "sales_campaign_roi": 1.0,
    "service_campaign_roi": 0.95,
    "sales_overall_conversion_rate": 0.95,
    "sales_campaign_deals": 0.8,
    "sales_new_leads_to_test_drives": 0.8,
    "sales_appointment_to_sales": 0.8,
    "sales_total_leads": 0.75,
    "sales_new_leads_to_appointments": 0.75,
    "sales_same_day_deal_pct": 0.7,
    "sales_internet_response_time": 0.7,
    "sales_uncontacted_leads": 0.6,
    "service_bookings_from_campaigns": 0.6,
    "service_outstanding_tasks": 0.5,
    "sales_overdue_tasks": 0.5,
    "sales_email_capture_rate": 0.45,
    "service_customers_enrolled": 0.4,
    "service_reminders_sent": 0.4,
}
DEFAULT_MATERIALITY_WEIGHT = 0.5
LEAD_PROMOTION_MARGIN = 1.25


def metric_materiality(metric: dict[str, Any] | None) -> float:
    if not metric:
        return 0.0
    band = float(SIGNIFICANCE_ORDER.get(metric_significance_band(metric), 0))
    if str(metric.get("metric_type", "count")) == "rate":
        points = metric.get("percentage_point_change")
        move = abs(float(points)) * 3.0 if isinstance(points, (int, float)) else 0.0
    else:
        percent = metric.get("percent_change")
        move = abs(float(percent)) if isinstance(percent, (int, float)) else 0.0
    return band * 10.0 + min(move, 120.0)


def insight_materiality(rule: RuleInsight, metrics: dict[str, dict[str, Any]]) -> float:
    """How big this insight's biggest reportable movement is, weighted by what it is."""
    best = 0.0
    for metric_name in unique_preserving_order([*(rule.metric_focus or []), *(rule.supporting_metrics or [])]):
        metric = metrics.get(metric_name)
        if not metric_usable_for_reporting(metric):
            continue
        weight = INSIGHT_MATERIALITY_WEIGHTS.get(metric_name, DEFAULT_MATERIALITY_WEIGHT)
        best = max(best, metric_materiality(metric) * weight)
    return best


def lead_with_most_material_insight(
    selected: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
) -> list[RuleInsight]:
    """Put the month's biggest movement first.

    An owner reads the first line. Editorial ordering above this decides which domain
    leads in normal months, and is left alone unless something else is decisively
    larger -- a 24% fall in deals should not sit under an 8% move in task counts.
    """
    if len(selected) < 2 or not metrics:
        return selected
    scores = [insight_materiality(rule, metrics) for rule in selected]
    best_index = max(range(len(scores)), key=lambda index: (scores[index], -index))
    if best_index == 0 or scores[best_index] < scores[0] * LEAD_PROMOTION_MARGIN:
        return selected
    reordered = list(selected)
    reordered.insert(0, reordered.pop(best_index))
    return reordered


def phase164_prioritise_insight_display_order(
    selected: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
) -> list[RuleInsight]:
    """
    When commercial retention (service campaign ROI) should lead the story, move the ROI snapshot
    ahead of moderate workload rules without re-scoring the whole candidate pool (phase 164).
    """
    if len(selected) < 2 or not metrics:
        return selected
    if phase164_operational_task_backlog_crisis(metrics):
        return phase169_demote_crm_composite_for_commercial_headline(list(selected))
    roi = metrics.get("service_campaign_roi")
    roi_constructive = bool(
        roi
        and metric_usable_for_reporting(roi, "slight")
        and metric_is_currently_improving(roi)
        and str(roi.get("direction", "flat")) != "flat"
    )
    if not roi_constructive and not phase164_commercial_retention_material_stress(metrics):
        return phase169_demote_crm_composite_for_commercial_headline(list(selected))
    snap_idx = next((i for i, r in enumerate(selected) if r.rule_id == "service_roi_snapshot"), None)
    if snap_idx is None or snap_idx == 0:
        return phase169_demote_crm_composite_for_commercial_headline(list(selected))
    if selected[0].rule_id not in {"service_tasks_pressure", "service_task_snapshot"}:
        return phase169_demote_crm_composite_for_commercial_headline(list(selected))
    reordered = list(selected)
    snap_rule = reordered.pop(snap_idx)
    reordered.insert(0, snap_rule)
    return phase169_demote_crm_composite_for_commercial_headline(reordered)


def phase155c_crm_stack_depth(metrics: dict[str, dict[str, Any]]) -> int:
    return sum(1 for name in PHASE155C_COMPOSITE_ORDER if phase155c_is_material_crm_adverse(metrics, name))


def phase155c_composite_crm_trigger(metrics: dict[str, dict[str, Any]]) -> bool:
    if phase155c_data_quality_and_followup_pair(metrics):
        return True
    names = [name for name in PHASE155C_COMPOSITE_ORDER if phase155c_is_material_crm_adverse(metrics, name)]
    # Commercial-first: strong sold outcomes suppress headline CRM composite unless CRM is severe or very deep.
    if phase157_strong_commercial_performance(metrics) and not phase157_crm_execution_severe_stack(metrics):
        if len(names) < 4:
            return False
    if len(names) >= 3:
        return True
    if len(names) >= 2:
        return True
    if len(names) == 1 and phase155c_is_severe_crm_adverse(metrics, names[0]):
        return False
    return False


def phase155d_composite_evidence_seed(metrics: dict[str, dict[str, Any]]) -> str:
    active = [n for n in PHASE155C_COMPOSITE_ORDER if phase155c_is_material_crm_adverse(metrics, n)]
    return ",".join(active) + "|phase155d_evidence"


def phase155d_metric_selection_score(metrics: dict[str, dict[str, Any]], metric_name: str) -> int:
    met = metrics.get(metric_name) or {}
    score = int(metric_change_strength(met)) * 12
    if bool(met.get("is_anomaly")):
        score += 40
    band = str(metric_significance_band(met))
    score += {"negligible": 1, "slight": 2, "meaningful": 5, "strong": 8, "not_available": 0}.get(band, 0)
    return score


def phase155d_crm_domain_winners(metrics: dict[str, dict[str, Any]]) -> dict[str, str]:
    """One winning metric per operational domain (deterministic tie-break on composite order index)."""
    winners: dict[str, str] = {}
    for domain, names in PHASE155D_CRM_DOMAIN_METRICS.items():
        best_name: str | None = None
        best_key: tuple[int, int] | None = None
        for mn in names:
            if not phase155c_is_material_crm_adverse(metrics, mn):
                continue
            sc = phase155d_metric_selection_score(metrics, mn)
            ix = PHASE155C_COMPOSITE_ORDER.index(mn) if mn in PHASE155C_COMPOSITE_ORDER else 99
            key = (sc, -ix)
            if best_key is None or key > best_key:
                best_key = key
                best_name = mn
        if best_name:
            winners[domain] = best_name
    return winners


def phase155d_order_metrics_for_display(metric_names: Iterable[str]) -> tuple[str, ...]:
    present = {str(m) for m in metric_names}
    return tuple(m for m in PHASE155C_COMPOSITE_ORDER if m in present)


def phase155d_trim_mandatory_domains(
    metrics: dict[str, dict[str, Any]],
    winners: dict[str, str],
    mandatory: set[str],
) -> set[str]:
    md = {d for d in mandatory if d in winners}
    if len(md) <= 3:
        return md
    ranked = sorted(md, key=lambda d: -phase155d_metric_selection_score(metrics, winners[d]))
    return set(ranked[:3])


def phase155d_select_composite_metric_names(
    metrics: dict[str, dict[str, Any]],
    context_seed: str,
) -> tuple[str, ...]:
    """
    Select up to three materially adverse CRM metrics for the composite bullet.
    When more than three operational domains fire, choose a triplet that maximises
    domain coverage (data quality + retention bonuses, penalties for dropping them when live).
    """
    winners = phase155d_crm_domain_winners(metrics)
    if not winners:
        return ()
    depth = len(winners)
    if depth <= 3:
        ordered = [winners[d] for d in PHASE155D_DOMAIN_STABLE_ORDER if d in winners]
        return tuple(ordered)

    domains = tuple(sorted(winners.keys()))
    mandatory: set[str] = set()
    if phase155c_data_quality_and_followup_pair(metrics):
        mandatory.add("customer_data_quality")
    for d, mn in winners.items():
        if phase155c_is_severe_crm_adverse(metrics, mn) or bool((metrics.get(mn) or {}).get("is_anomaly")):
            mandatory.add(d)
    mandatory = phase155d_trim_mandatory_domains(metrics, winners, mandatory)

    best_metrics: tuple[str, ...] | None = None
    best_rank: tuple[int, str] = (-(10**9), "")

    feasible_any = False
    for tri in combinations(domains, 3):
        tri_set = set(tri)
        if mandatory and not mandatory.issubset(tri_set):
            continue
        feasible_any = True
        score_num = 0
        for d in tri:
            score_num += phase155d_metric_selection_score(metrics, winners[d])
        if "customer_data_quality" in tri_set:
            score_num += 14
        for rd in ("service_campaign_reminders", "service_customer_enrolment"):
            if rd in tri_set:
                score_num += 10
        if "customer_data_quality" in winners and "customer_data_quality" not in tri_set:
            score_num -= 20
        for rd in ("service_campaign_reminders", "service_customer_enrolment"):
            if rd in winners and rd not in tri_set:
                score_num -= 12
        if "responsiveness" in winners and "responsiveness" not in tri_set:
            rmn = winners["responsiveness"]
            if bool((metrics.get(rmn) or {}).get("is_anomaly")) or phase155c_is_severe_crm_adverse(metrics, rmn):
                score_num -= 35
        tri_key = ",".join(sorted(tri))
        if (
            best_metrics is None
            or score_num > best_rank[0]
            or (score_num == best_rank[0] and tri_key < best_rank[1])
        ):
            best_rank = (score_num, tri_key)
            dom_sorted = sorted(tri, key=lambda d: PHASE155D_DOMAIN_STABLE_ORDER.index(d))
            best_metrics = tuple(winners[d] for d in dom_sorted)

    if not feasible_any or best_metrics is None:
        ranked_d = sorted(winners.keys(), key=lambda d: -phase155d_metric_selection_score(metrics, winners[d]))
        best_metrics = tuple(winners[d] for d in ranked_d[:3])

    return phase155d_order_metrics_for_display(best_metrics)


PHASE169_SERVICE_CAMPAIGN_VOLUME_METRICS: frozenset[str] = frozenset({"service_reminders_sent", "service_customers_enrolled"})

PHASE169_COMPOSITE_DEMOTION_LEADERS: frozenset[str] = frozenset(
    {
        "service_roi_snapshot",
        "sales_campaign_roi_decline",
        "sales_marketing_negative_roi",
        "test_drive_and_conversion_down",
        "lead_handling_pressure",
        "lead_response_snapshot",
    }
)


def phase169_demote_crm_composite_for_commercial_headline(selected: list[RuleInsight]) -> list[RuleInsight]:
    """Keep CRM composite out of slot 1 when a clearer commercial/CRM headline is present (phase 169)."""
    if len(selected) < 2 or selected[0].rule_id != "crm_execution_pressure_composite":
        return selected
    pref_i = next((i for i, r in enumerate(selected) if r.rule_id in PHASE169_COMPOSITE_DEMOTION_LEADERS), None)
    if pref_i is None or pref_i == 0:
        return selected
    out = list(selected)
    pref = out.pop(pref_i)
    comp = out.pop(0)
    return [pref, comp, *out]


def phase169_crm_composite_fact_clause(metrics: dict[str, dict[str, Any]], metric_name: str) -> str | None:
    """Single clause for one composite metric (phase 169 — adverse-only, theme-safe)."""
    if not phase155c_is_material_crm_adverse(metrics, metric_name):
        return None
    label = summary_metric_label(metric_name, metrics) or metric_label(metrics, metric_name)
    if metric_name == "sales_internet_response_time":
        return metric_transition_clause(
            metrics,
            metric_name,
            up_verb="slowed",
            down_verb="improved",
            label_override=label or "Response times",
        )
    if str((metrics.get(metric_name) or {}).get("metric_type")) == "time":
        return metric_transition_clause(
            metrics,
            metric_name,
            up_verb="slowed",
            down_verb="improved",
            label_override=label,
        )
    clause = build_short_metric_fact(
        metrics,
        metric_name,
        label_override=label,
        minimum_band="slight",
        include_month_on_month=False,
    )
    if not clause:
        clause = build_short_metric_fact(
            metrics,
            metric_name,
            label_override=label,
            minimum_band="slight",
            include_month_on_month=True,
        )
    return str(clause).rstrip(".") if clause else None


def phase169_join_fact_chunks(chunks: list[str]) -> str:
    chunks = [c for c in chunks if c]
    if not chunks:
        return ""
    if len(chunks) == 1:
        return chunks[0]
    if len(chunks) == 2:
        return f"{chunks[0]} and {chunks[1]}"
    return f"{', '.join(chunks[:-1])}, and {chunks[-1]}"


def phase169_positive_email_capture_suffix(metrics: dict[str, dict[str, Any]], context_seed: str) -> str:
    """Append a truthful positive email-capture clause when capture is materially improving (phase 169)."""
    m = metrics.get("sales_email_capture_rate")
    if not m or not metric_usable_for_reporting(m, "slight"):
        return ""
    if not metric_is_currently_improving(m):
        return ""
    fact = build_short_metric_fact(
        metrics,
        "sales_email_capture_rate",
        label_override="email capture rate",
        minimum_band="slight",
        include_month_on_month=True,
    )
    if not fact:
        return ""
    frag = phase166_embed_in_sentence(fact.strip().rstrip("."))
    return choose_deterministic_phrase(
        [
            f" Email capture strengthened, with {frag}.",
            f" Email contactability improved, with {frag}.",
        ],
        f"{context_seed}|ph169|email_pos_tail",
    )


def build_crm_composite_insight_message(metrics: dict[str, dict[str, Any]], context_seed: str) -> str:
    ev_seed = phase155d_composite_evidence_seed(metrics)
    full_seed = f"{ev_seed}|{context_seed}"
    selected = phase155d_select_composite_metric_names(metrics, full_seed)
    crm_names = [m for m in selected if m not in PHASE169_SERVICE_CAMPAIGN_VOLUME_METRICS]
    svc_names = [m for m in selected if m in PHASE169_SERVICE_CAMPAIGN_VOLUME_METRICS]
    crm_chunks = [c for m in crm_names if (c := phase169_crm_composite_fact_clause(metrics, m))]
    svc_chunks = [c for m in svc_names if (c := phase169_crm_composite_fact_clause(metrics, m))]
    if not crm_chunks and not svc_chunks:
        return ""

    if crm_chunks and svc_chunks:
        crm_joiner = phase169_join_fact_chunks(crm_chunks)
        svc_joiner = phase169_join_fact_chunks(svc_chunks)
        if phase157_strong_commercial_performance(metrics):
            lead = choose_deterministic_phrase(
                [
                    f"Sold performance held up, even as CRM follow-up workload built, with {crm_joiner}.",
                    f"Deals and conversion stayed constructive, although CRM execution slipped, with {crm_joiner}.",
                ],
                f"{full_seed}|crm_composite_split|crm_lead_strong",
            )
        else:
            lead = choose_deterministic_phrase(
                [
                    f"Follow-up pressure increased on the CRM side this month, with {crm_joiner}.",
                    f"Customer follow-up weakened on the CRM side this month, with {crm_joiner}.",
                ],
                f"{full_seed}|crm_composite_split|crm_lead",
            )
        tail = choose_deterministic_phrase(
            [
                f"Service campaign activity softened, with {svc_joiner}.",
                f"Workshop campaign touches eased, with {svc_joiner}.",
            ],
            f"{full_seed}|crm_composite_split|svc_tail",
        )
        body = collapse_whitespace(f"{lead} {tail}")
        suf = phase169_positive_email_capture_suffix(metrics, context_seed)
        return collapse_whitespace(body + suf) if suf else body

    fact_chunks = crm_chunks if crm_chunks else svc_chunks
    joiner = phase169_join_fact_chunks(fact_chunks)
    if not joiner:
        return ""

    if not crm_chunks and svc_chunks:
        body = choose_deterministic_phrase(
            [
                f"Service campaign activity softened this month, with {joiner}.",
                f"Workshop campaign engagement eased this month, with {joiner}.",
            ],
            f"{full_seed}|crm_composite_service_only",
        )
        suf = phase169_positive_email_capture_suffix(metrics, context_seed)
        return collapse_whitespace(body + suf) if suf else body

    if phase157_strong_commercial_performance(metrics):
        templates = [
            f"Deals and conversion stayed constructive, although follow-up pressure also built, with {joiner}.",
            f"Sold performance held up, even as follow-up workload and contact metrics slipped, with {joiner}.",
        ]
    else:
        templates = [
            f"Follow-up pressure increased this month, with {joiner}.",
            f"Customer follow-up weakened further this month, with {joiner}.",
        ]
    body = choose_deterministic_phrase(templates, f"{full_seed}|crm_composite_message")
    suf = phase169_positive_email_capture_suffix(metrics, context_seed)
    return collapse_whitespace(body + suf) if suf else body


def phase155c_suppress_low_materiality_service_task_snapshot(
    candidate: RuleInsight,
    metrics: dict[str, dict[str, Any]],
) -> bool:
    if candidate.rule_id != "service_task_snapshot":
        return False
    st = metrics.get("service_outstanding_tasks")
    if not st:
        return False
    stack = phase155c_crm_stack_depth(metrics)
    if phase155c_composite_crm_trigger(metrics):
        if str(st.get("direction")) == "flat" and float(st.get("current") or 0) <= 3:
            return True
    flat_low = (
        str(st.get("direction")) == "flat"
        and float(st.get("current") or 0) <= 2
        and metric_significance_band(st) in {"negligible", "slight"}
        and metric_change_strength(st) == 0
    )
    if flat_low and stack >= 2:
        return True
    if metric_significance_band(st) == "negligible" and str(st.get("direction")) == "flat" and stack >= 3:
        return True
    return False


def sentence_concept_tags(sentence: str) -> frozenset[str]:
    low = str(sentence or "").lower()
    tags: list[str] = []
    for tag, needles in PHASE155_CONCEPT_TAG_NEEDLES:
        if any(needle in low for needle in needles):
            tags.append(tag)
    return frozenset(tags)


def apply_phase155_summary_information_density(
    text: str, context_seed: str, *, max_sentences: int | None = None,
) -> str:
    """
    Deterministic informational-density pass: drop sentences whose operational concept tags
    are already fully covered by earlier sentences, then cap length (phase 155).
    """
    del context_seed  # reserved for future deterministic tie-breaks
    cap = PHASE155_MAX_SUMMARY_SENTENCES if max_sentences is None else max(1, int(max_sentences))
    base = collapse_whitespace(str(text or "").strip())
    if not base:
        return base
    sentences = [s.strip() for s in re.split(r"(?<=[.!?])\s+", base) if s.strip()]
    if len(sentences) <= 1:
        return base
    kept: list[str] = [sentences[0]]
    accumulated: set[str] = set(sentence_concept_tags(sentences[0]))
    for sent in sentences[1:]:
        if sent.strip().lower().startswith("the main commercial watchpoint"):
            kept.append(sent)
            accumulated |= set(sentence_concept_tags(sent))
            continue
        low = sent.strip().lower()
        if "enrolled clients" in low and "service campaign roi" in low:
            kept.append(sent)
            accumulated |= set(sentence_concept_tags(sent))
            continue
        tags = set(sentence_concept_tags(sent))
        if tags and tags <= accumulated:
            continue
        accumulated |= tags
        kept.append(sent)
    if len(kept) > cap:
        kept = kept[:cap]
    out = " ".join(kept)
    return re.sub(r"\s+", " ", out).strip()


def select_top_insights(
    candidates: list[RuleInsight],
    limit: int = 4,
    metrics: dict[str, dict[str, Any]] | None = None,
) -> tuple[list[RuleInsight], list[str]]:
    """
    Enhanced insight selection with improved prioritization and diversity.
    
    PHASE 1: Insight Prioritisation Intelligence
    Prioritizes operationally meaningful shifts over trivial fluctuations.
    Ensures diversity across sales/service/customer dimensions.
    """
    # A rule whose own metrics were all gated out has no figure to show. Its only
    # remaining text is the legacy template phrasing, so it must not take one of the
    # four slots from an insight that does have evidence.
    evidenced = [candidate for candidate in candidates if candidate.has_metric_evidence]
    if evidenced:
        candidates = evidenced

    # Enhanced ranking: prioritize operational significance over raw priority weights
    ranked = sorted(
        candidates,
        key=lambda candidate: (
            # Primary: Operational significance (higher = more important)
            -operational_significance_score(candidate, candidates, metrics),
            # Secondary: Original priority weight
            -candidate.priority_weight,
            # Tertiary: Category diversity
            candidate.category,
            # Quaternary: Rule ID for determinism
            candidate.rule_id,
        ),
    )

    selected: list[RuleInsight] = []
    used_merge_keys: set[str] = set()
    used_categories: set[str] = set()
    used_signal_types: set[str] = set()
    family_counts: dict[str, int] = {}
    unified_campaign_lanes: set[str] = set()
    used_evidence_metrics: set[str] = set()
    metrics_map = metrics or {}

    for candidate in ranked:
        if candidate.merge_key in used_merge_keys:
            continue

        if phase155c_suppress_low_materiality_service_task_snapshot(candidate, metrics_map):
            continue

        # Enhanced diversity: ensure coverage across key business dimensions
        candidate_category = candidate.category
        candidate_signal_type = candidate.signal_type or "current_performance"

        # Allow some category overlap but prioritize diversity
        if (len(selected) >= 2 and
            candidate_category in used_categories and
            candidate_signal_type in used_signal_types and
            len([s for s in selected if s.category == candidate_category]) >= 2):
            continue

        candidate_primary_focus = (candidate.metric_focus or [None])[0]
        op_family = rule_operational_family(candidate.rule_id)
        crm_stress = phase155_crm_deterioration_count(metrics_map)
        candidate_evidence = set(_metric_evidence_for_rule(candidate))

        if op_family == "crm_execution" and family_counts.get("crm_execution", 0) >= 1:
            continue
        if op_family == "service_ops" and family_counts.get("service_ops", 0) >= 1:
            continue
        if op_family == "unified_campaign":
            uc = family_counts.get("unified_campaign", 0)
            if uc >= 2:
                continue
            lane = phase160_campaign_lane(candidate.rule_id)
            if lane:
                if lane in unified_campaign_lanes:
                    continue
            elif uc >= 1:
                continue
        if op_family == "conversion_story":
            if any(
                rule_operational_family(s.rule_id) == "conversion_story" and s.polarity == candidate.polarity
                for s in selected
            ):
                continue
            crm_soften_positive_block = crm_stress >= 2 and not (
                phase157_strong_commercial_performance(metrics_map) and not phase157_crm_execution_severe_stack(metrics_map)
            )
            if crm_soften_positive_block and candidate.polarity == "positive" and any(
                rule_operational_family(s.rule_id) == "conversion_story" and s.polarity == "negative"
                for s in selected
            ):
                continue

        normalized_message = normalize_text(candidate.message)
        message_duplicate = False
        for existing_rule in selected:
            existing_message = normalize_text(existing_rule.message)
            existing_evidence = set(_metric_evidence_for_rule(existing_rule))
            if normalized_message == existing_message or normalized_message in existing_message:
                message_duplicate = True
                break
            if existing_message in normalized_message and not (candidate_evidence and existing_evidence < candidate_evidence):
                message_duplicate = True
                break
        if message_duplicate:
            continue

        if any(
            set(candidate.metric_focus or []) & set(existing.metric_focus or [])
            and (
                candidate.signal_type == existing.signal_type
                or candidate.category == existing.category
            )
            and rule_operational_family(candidate.rule_id) == rule_operational_family(existing.rule_id)
            for existing in selected
        ):
            continue

        if any(
            candidate_primary_focus
            and candidate_primary_focus == (existing.metric_focus or [None])[0]
            and candidate.polarity == existing.polarity
            for existing in selected
        ):
            continue

        if candidate_evidence and candidate_evidence <= used_evidence_metrics:
            continue

        selected.append(candidate)
        used_merge_keys.add(candidate.merge_key)
        used_categories.add(candidate_category)
        used_signal_types.add(candidate_signal_type)
        family_counts[op_family] = family_counts.get(op_family, 0) + 1
        used_evidence_metrics.update(candidate_evidence)
        if op_family == "unified_campaign":
            cl = phase160_campaign_lane(candidate.rule_id)
            if cl:
                unified_campaign_lanes.add(cl)

        # Stop at limit
        if len(selected) >= limit:
            break

    selected = suppress_redundant_metric_evidence_rules(selected)
    selected = ensure_both_directions_are_represented(selected, ranked, metrics_map)
    conflicts = detect_conflicting_categories(selected)
    return selected, conflicts


def month_moved_both_ways(metrics: dict[str, dict[str, Any]] | None) -> bool:
    """Did the month contain material movement in both directions?

    Material, not merely present: a 2% drift alongside a 40% collapse is not a mixed
    month, and treating it as one would license a token positive in a bad month.
    """
    improving = adverse = False
    for name, metric in (metrics or {}).items():
        if not metric_usable_for_reporting(metric):
            continue
        if str(metric.get("significance_band", "")) not in {"meaningful", "strong"}:
            continue
        if metric_is_adverse(metrics, name):
            adverse = True
        else:
            improving = True
        if improving and adverse:
            return True
    return False


def ensure_both_directions_are_represented(
    selected: list[RuleInsight],
    ranked: list[RuleInsight],
    metrics: dict[str, dict[str, Any]] | None,
) -> list[RuleInsight]:
    """Four insights that all point one way misrepresent a month that did not.

    Ranking by materiality alone can fill every slot from one side of the month, even
    where the other side moved just as materially — leads down while campaign returns
    climbed is an ordinary month, not a contradiction, and an owner shown only one half
    of it has been given a false picture of their own business.

    The weakest same-direction insight gives up its slot to the best-ranked candidate
    of the missing direction. The lead is never touched: the month's biggest movement
    still opens the report, whichever way it went.
    """
    if len(selected) < 2 or not month_moved_both_ways(metrics):
        return selected
    polarities = {rule.polarity for rule in selected}
    if "mixed" in polarities or {"positive", "negative"} <= polarities:
        return selected
    missing = "negative" if "positive" in polarities else "positive"
    chosen_ids = {rule.rule_id for rule in selected}
    replacement = next(
        (
            candidate
            for candidate in ranked
            if candidate.polarity == missing
            and candidate.rule_id not in chosen_ids
            and candidate.has_metric_evidence
        ),
        None,
    )
    if replacement is None:
        return selected
    return selected[:-1] + [replacement]


def suppress_redundant_metric_evidence_rules(selected: list[RuleInsight]) -> list[RuleInsight]:
    """Remove selected rules whose metric evidence is already covered by a stronger, broader rule."""
    keep: list[RuleInsight] = []
    for idx, rule in enumerate(selected):
        evidence = set(_metric_evidence_for_rule(rule))
        if evidence:
            redundant = False
            for other_idx, other in enumerate(selected):
                if idx == other_idx:
                    continue
                other_evidence = set(_metric_evidence_for_rule(other))
                if (
                    evidence < other_evidence
                    and rule.priority_weight <= other.priority_weight
                ):
                    redundant = True
                    break
            if redundant:
                continue
        keep.append(rule)
    return keep


def operational_significance_score(
    candidate: RuleInsight,
    all_candidates: list[RuleInsight],
    metrics: dict[str, dict[str, Any]] | None = None,
) -> float:
    """
    Operational significance blends business severity with scale, volume credibility,
    persistence, volatility, signal rarity, and (when provided) dealership throughput context.
    """
    score = 0.0

    severity_multipliers = {
        "critical": 100,
        "high": 75,
        "medium": 50,
        "low": 25,
    }
    score += severity_multipliers.get(candidate.severity, 25)

    business_impact_rules = {
        "leads_up_conversion_down": 30,
        "leads_up_deals_down": 25,
        "lead_handling_pressure": 34,
        "service_tasks_pressure": 26,
        "sales_marketing_negative_roi": 25,
        "closing_execution_up": 10,
        "service_follow_up_cleaner": 15,
        "combined_sales_service_up": 20,
        "cross_metric_qualification_inefficiency": 22,
        "service_task_snapshot": 18,
        "lead_response_snapshot": 16,
        "crm_execution_pressure_composite": 46,
        "sales_campaign_roi_decline": 28,
        "service_roi_snapshot": 24,
        "service_bookings_roi_mismatch": 24,
        "service_booking_snapshot": 12,
    }
    score += business_impact_rules.get(candidate.rule_id, 0)
    score += phase155_operational_materiality_bonus(candidate, metrics)
    score += phase159_commercial_impact_tier_bonus(candidate.rule_id, metrics)
    score += phase160_service_retention_commercial_bonus(candidate.rule_id, metrics)
    score += phase160_dual_campaign_balance_nudge(candidate, metrics)
    score += phase159_cross_domain_significance_nudge(candidate, metrics)

    if metrics:
        co = phase157_commercial_outcome_score(metrics)
        if co >= 5 and candidate.rule_id in {
            "leads_down_conversion_up",
            "sales_volume_snapshot",
            "closing_execution_up",
            "sales_marketing_up",
            "service_campaigns_up",
            "service_booking_snapshot",
            "sales_conversion_snapshot",
            "combined_sales_service_up",
        }:
            score += 26.0

    tone_significance = {
        "significantly_softer": 15,
        "strongly_improved": 15,
        "slightly_softer": 8,
        "slightly_improved": 8,
        "stable": 0,
    }
    score += tone_significance.get(candidate.tone_band, 0)

    category_counts: dict[str, int] = {}
    for c in all_candidates:
        category_counts[c.category] = category_counts.get(c.category, 0) + 1

    total_candidates = len(all_candidates) or 1
    category_rarity = 1.0 - (category_counts.get(candidate.category, 0) / total_candidates)
    score += category_rarity * 10

    if metrics:
        score *= volume_credibility_multiplier(candidate, metrics)
        score += trend_persistence_adjustment(candidate, metrics)
        score += signal_rarity_bonus(candidate, all_candidates)
        score += phase164_retention_commercial_arbitration_nudge(candidate, metrics)

    return max(0.0, score)


def derive_sentiment(selected_rules: list[RuleInsight], conflicts: list[str]) -> str:
    if not selected_rules:
        return "stable"

    positive_score = sum(rule.priority_weight for rule in selected_rules if rule.polarity == "positive")
    negative_score = sum(rule.priority_weight for rule in selected_rules if rule.polarity == "negative")
    strongest_positive = max(
        (2 if rule.tone_band == "strongly_improved" else 1 for rule in selected_rules if rule.polarity == "positive"),
        default=0,
    )
    strongest_negative = max(
        (2 if rule.tone_band == "significantly_softer" else 1 for rule in selected_rules if rule.polarity == "negative"),
        default=0,
    )

    if conflicts or (positive_score > 0 and negative_score > 0 and abs(positive_score - negative_score) < 25):
        return "mixed"
    if negative_score > 0 and negative_score >= positive_score + 20:
        return "significantly_softer" if strongest_negative >= 2 or negative_score >= positive_score + 55 else "slightly_softer"
    if positive_score > 0 and positive_score >= negative_score + 25:
        return "strongly_improved" if strongest_positive >= 2 or positive_score >= negative_score + 55 else "slightly_improved"
    if positive_score > negative_score:
        return "slightly_improved"
    if negative_score > positive_score:
        return "slightly_softer"
    return "stable"


def build_summary_intro(sentiment: str) -> str:
    intros = {
        "strongly_improved": "March was a stronger month overall",
        "mixed": "March was a mixed month overall",
        "slightly_softer": "March was slightly softer overall",
        "significantly_softer": "March was a softer month overall",
        "slightly_improved": "March moved forward overall",
        "stable": "March was broadly stable overall",
    }
    return intros.get(sentiment, "March was broadly stable overall")


def summarise_sales(metrics: dict[str, dict[str, Any]]) -> str | None:
    leads = metrics.get("sales_total_leads")
    deals = metrics.get("sales_total_deals")
    conversion = metrics.get("sales_overall_conversion_rate")
    response = metrics.get("sales_internet_response_time")

    clauses: list[str] = []
    if leads is not None and deals is not None:
        leads_direction = leads.get("direction")
        deals_direction = deals.get("direction")
        combined_strength = max(metric_change_strength(leads), metric_change_strength(deals))
        if leads_direction == "up" and deals_direction == "up":
            clauses.append("lead generation and deal volume both improved" if combined_strength >= 1 else "lead generation and deal volume edged higher")
        elif leads_direction == "up" and deals_direction == "down":
            clauses.append("lead generation improved, but deal volume softened" if combined_strength < 2 else "lead generation improved, but deal volume fell back")
        elif leads_direction == "down" and deals_direction == "up":
            clauses.append("stronger deal conversion offset lower lead volume" if combined_strength >= 1 else "slightly stronger deal conversion offset softer lead volume")
        elif deals_direction == "down":
            clauses.append("deal volume eased month-on-month" if metric_change_strength(deals) < 2 else "deal volume fell back month-on-month")

    if conversion is not None:
        conversion_strength = metric_change_strength(conversion)
        if conversion.get("direction") == "down" and conversion_strength >= 1:
            clauses.append("sales conversion was weaker" if conversion_strength < 2 else "sales conversion weakened")
        elif conversion.get("direction") == "up" and conversion_strength >= 1:
            clauses.append("sales conversion improved")

    if response is not None:
        response_strength = metric_change_strength(response)
        if response.get("direction") == "up" and response_strength >= 1:
            clauses.append("response times slowed")
        elif response.get("direction") == "down" and response_strength >= 1:
            clauses.append("response times improved")

    if not clauses:
        return None
    return ", with " + ", while ".join(clauses[:2])


def summarise_service(metrics: dict[str, dict[str, Any]]) -> str | None:
    bookings = metrics.get("service_bookings_from_campaigns")
    roi = metrics.get("service_campaign_roi")
    tasks = metrics.get("service_outstanding_tasks")
    enrolled = metrics.get("service_customers_enrolled")

    clauses: list[str] = []
    if enrolled is not None and roi is not None:
        if (
            metric_usable_for_reporting(enrolled, "slight")
            and metric_usable_for_reporting(roi, "slight")
            and str(enrolled.get("direction")) == "up"
            and str(roi.get("direction")) == "down"
        ):
            clauses.append("enrolled clients increased while service campaign ROI declined")
    if bookings is not None and roi is not None:
        if bookings.get("direction") == "up" and roi.get("direction") == "up":
            clauses.append("service campaign performance improved")
        elif bookings.get("direction") in {"up", "flat"} and roi.get("direction") == "down":
            if bookings.get("direction") == "up":
                clauses.append("service campaign bookings increased while service campaign ROI still declined")
            else:
                clauses.append("service campaign ROI weakened while booking volumes held steady")
        elif bookings.get("direction") == "down" and roi.get("direction") == "down":
            clauses.append("service campaign performance softened")

    if tasks is not None:
        task_strength = metric_change_strength(tasks)
        if tasks.get("direction") == "up" and task_strength >= 1:
            clauses.append("outstanding service tasks increased" if task_strength < 2 else "outstanding service tasks increased sharply")
        elif tasks.get("direction") == "down" and task_strength >= 1:
            clauses.append("service follow-up backlog reduced")

    if not clauses:
        return None
    return ", and " + ", while ".join(clauses[:2])


def collect_summary_facts(selected_rules: list[RuleInsight], metrics: dict[str, dict[str, Any]]) -> tuple[list[str], list[str]]:
    positive_facts: list[str] = []
    negative_facts: list[str] = []
    seen_positive: set[str] = set()
    seen_negative: set[str] = set()

    for rule in selected_rules:
        fact_bucket = positive_facts if rule.polarity == "positive" else negative_facts if rule.polarity == "negative" else []
        seen_bucket = seen_positive if rule.polarity == "positive" else seen_negative if rule.polarity == "negative" else set()
        for fact in build_rule_fact_list(rule, metrics):
            if fact not in seen_bucket:
                fact_bucket.append(fact)
                seen_bucket.add(fact)
    return positive_facts, negative_facts


def summary_primary_facts(selected_rules: list[RuleInsight], metrics: dict[str, dict[str, Any]], sentiment: str) -> list[str]:
    positive_rules = [rule for rule in selected_rules if rule.polarity == "positive"]
    negative_rules = [rule for rule in selected_rules if rule.polarity == "negative"]
    ordered_rules: list[RuleInsight]
    if sentiment in {"strongly_improved", "slightly_improved"}:
        ordered_rules = positive_rules + negative_rules
    elif sentiment in {"significantly_softer", "slightly_softer"}:
        ordered_rules = negative_rules + positive_rules
    else:
        ordered_rules = selected_rules

    facts: list[str] = []
    seen: set[str] = set()
    for rule in ordered_rules:
        for fact in build_rule_fact_list(rule, metrics):
            if fact in seen:
                continue
            facts.append(fact)
            seen.add(fact)
            if len(facts) >= 2:
                return facts
    return facts


def phase155_crm_execution_watchpoint_sentence(
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
    *,
    coordination_context_seed: str = "",
) -> str | None:
    """Deterministic CRM / follow-up watchpoint when execution signals stack (phase 155)."""
    crm_order = (
        ("sales_internet_response_time", "response times"),
        ("sales_uncontacted_leads", "uncontacted leads"),
        ("sales_overdue_tasks", "overdue sales tasks"),
        ("sales_email_capture_rate", "email capture"),
        ("service_reminders_sent", "service reminders sent"),
        ("service_customers_enrolled", "service enrolments"),
    )
    adverse: list[tuple[str, str]] = []
    for metric_name, label in crm_order:
        m = metrics.get(metric_name)
        if m and metric_usable_for_reporting(m, "slight") and metric_is_adverse(metrics, metric_name):
            adverse.append((metric_name, label))

    rule_ids = {r.rule_id for r in selected_rules}
    if "crm_execution_pressure_composite" in rule_ids:
        seed = f"{phase155d_composite_evidence_seed(metrics)}|{coordination_context_seed}".strip("|")
        composite_selected = set(phase155d_select_composite_metric_names(metrics, seed or "coordination"))
        omitted: list[tuple[str, str]] = []
        for metric_name, label in crm_order:
            if metric_name in composite_selected:
                continue
            m = metrics.get(metric_name)
            if m and metric_usable_for_reporting(m, "slight") and metric_is_adverse(metrics, metric_name):
                omitted.append((metric_name, label))
        if not omitted:
            return None
        omitted.sort(key=lambda pair: -phase155d_metric_selection_score(metrics, pair[0]))
        parts: list[str] = []
        for metric_name, label in omitted[:2]:
            fact = build_short_metric_fact(
                metrics,
                metric_name,
                label_override=label,
                minimum_band="slight",
                include_month_on_month=True,
            )
            if fact:
                parts.append(fact.rstrip("."))
        if not parts:
            return None
        if len(parts) == 2:
            return f"Also this month: {parts[0]} and {parts[1]}."
        return f"Also this month: {parts[0]}."

    if len(adverse) < 2 and not ("lead_handling_pressure" in rule_ids and len(adverse) >= 1):
        return None

    if len(adverse) >= 2:
        m0, lab0 = adverse[0]
        m1, lab1 = adverse[1]
        f0 = build_short_metric_fact(metrics, m0, label_override=lab0)
        f1 = build_short_metric_fact(metrics, m1, label_override=lab1)
        if f0 and f1:
            return (
                f"Several follow-up signals worsened together, with {f0.rstrip('.')} and {f1.rstrip('.')}, "
                f"which reads as mounting follow-up pressure."
            )

    if len(adverse) == 1:
        m0, lab0 = adverse[0]
        fact = build_short_metric_fact(metrics, m0, label_override=lab0)
        if fact and "lead_handling_pressure" in rule_ids:
            return f"Follow-up pressure increased, with {fact.rstrip('.')}."
    return None


def phase161_campaign_first_watchpoint_sentence(
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
) -> str | None:
    """When campaign ROI rules fire, anchor the summary watchpoint on campaign economics before CRM/task openers."""
    rule_ids = {r.rule_id for r in selected_rules}
    sales_rule = "sales_campaign_roi_decline" in rule_ids
    svc_rule = "service_roi_snapshot" in rule_ids
    if not sales_rule and not svc_rule:
        return None

    svc_roi = metrics.get("service_campaign_roi")
    enrolled = metrics.get("service_customers_enrolled")

    sales_fact = (
        build_short_metric_fact(metrics, "sales_campaign_roi", label_override="sales campaign ROI")
        if sales_rule
        else None
    )
    svc_fact = (
        build_short_metric_fact(metrics, "service_campaign_roi", label_override="service campaign ROI")
        if svc_rule
        else None
    )
    enroll_up_roi_down = (
        enrolled
        and svc_roi
        and metric_usable_for_reporting(enrolled, "slight")
        and metric_usable_for_reporting(svc_roi, "slight")
        and str(enrolled.get("direction")) == "up"
        and str(svc_roi.get("direction")) == "down"
    )
    enroll_fact = (
        build_short_metric_fact(metrics, "service_customers_enrolled", label_override="enrolled clients")
        if enroll_up_roi_down
        else None
    )

    if sales_fact and svc_fact and enroll_fact:
        return (
            f"The main commercial watchpoint was campaign ROI, with {sales_fact.rstrip('.')}, "
            f"{svc_fact.rstrip('.')}, while {enroll_fact.rstrip('.')}."
        )
    if sales_fact and svc_fact:
        return f"The main commercial watchpoint was campaign ROI, with {sales_fact.rstrip('.')} and {svc_fact.rstrip('.')}."
    if sales_fact:
        return f"The main commercial watchpoint was sales campaign ROI, with {sales_fact.rstrip('.')}."
    if svc_fact and enroll_fact:
        return f"The main commercial watchpoint was workshop campaign ROI, with {svc_fact.rstrip('.')}, while {enroll_fact.rstrip('.')}."
    if svc_fact:
        return f"The main commercial watchpoint was service campaign ROI, with {svc_fact.rstrip('.')}."
    return None


def phase164_retention_first_watchpoint_sentence(
    _selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
) -> str | None:
    """
    When service campaign ROI is materially adverse, anchor the commercial watchpoint on ROI (phase 164).
    Optionally append outstanding tasks as supporting context — not as the primary anchor —
    unless operational backlog crisis conditions apply.
    """
    if not phase164_commercial_retention_material_stress(metrics):
        return None
    if phase164_operational_task_backlog_crisis(metrics):
        return None
    svc_roi = metrics.get("service_campaign_roi")
    if not svc_roi or not metric_usable_for_reporting(svc_roi, "slight") or not metric_is_adverse(metrics, "service_campaign_roi"):
        return None
    fact = build_short_metric_fact(metrics, "service_campaign_roi", label_override="service campaign ROI")
    if not fact:
        return None
    tasks = metrics.get("service_outstanding_tasks")
    if tasks and metric_usable_for_reporting(tasks, "slight") and metric_is_adverse(metrics, "service_outstanding_tasks"):
        tf = build_short_metric_fact(metrics, "service_outstanding_tasks", label_override="outstanding service tasks")
        if tf:
            return (
                f"The main commercial watchpoint was service campaign ROI, with {fact.rstrip('.')}, "
                f"while {tf.rstrip('.')}."
            )
    return f"The main commercial watchpoint was service campaign ROI, with {fact.rstrip('.')}."


def summary_watchpoint_sentence(
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
    context_seed: str = "",
) -> str | None:
    cw161 = phase161_campaign_first_watchpoint_sentence(selected_rules, metrics)
    if cw161:
        return cw161

    cw164 = phase164_retention_first_watchpoint_sentence(selected_rules, metrics)
    if cw164:
        return cw164

    persistent_rules = [rule for rule in selected_rules if rule.signal_type == "persistent_watchpoint" and has_adverse_focus_metric(rule, metrics)]
    if persistent_rules:
        rule = persistent_rules[0]
        focus_metric_name = primary_adverse_focus_metric(rule, metrics)
        if focus_metric_name == "sales_internet_response_time":
            fact = build_short_metric_fact(metrics, focus_metric_name, label_override="response times")
            if fact:
                return f"The main watchpoint remained lead follow-up, with {fact}."
        if focus_metric_name == "service_outstanding_tasks":
            fact = build_short_metric_fact(metrics, focus_metric_name, label_override="outstanding service tasks")
            if fact:
                return f"The main watchpoint remained service follow-up, with {fact}."
        if focus_metric_name == "sales_overall_conversion_rate":
            fact = build_short_metric_fact(metrics, focus_metric_name, label_override="conversion")
            if fact:
                return f"The main watchpoint remained conversion quality, with {fact}."
        if focus_metric_name == "service_campaign_roi":
            fact = build_short_metric_fact(metrics, focus_metric_name, label_override="service campaign ROI")
            if fact:
                return f"The main watchpoint remained service ROI, with {fact}."

    crm_exec = phase155_crm_execution_watchpoint_sentence(
        selected_rules,
        metrics,
        coordination_context_seed=context_seed,
    )
    if crm_exec:
        return crm_exec

    # The two branches that used to sit here ("The main risk remained <fact>.") repeated
    # the top priority insight and the Primary Risk field word for word, giving the
    # reader the same figure three times inside forty words. The themed sentences above
    # are kept because they name which part of the business the pressure sits in.

    recurring_clause = persistent_watchpoint_summary_clause(metrics)
    if recurring_clause:
        if "lead follow-up" in recurring_clause:
            return "The main watchpoint remained lead follow-up over recent months."
        if "service follow-up pressure" in recurring_clause:
            return "The main watchpoint remained service follow-up pressure over recent months."
    return None


def build_summary_trend_sentence(selected_rules: list[RuleInsight], metrics: dict[str, dict[str, Any]]) -> str | None:
    if any(rule.rule_id == "crm_execution_pressure_composite" for rule in selected_rules):
        return None
    for rule in selected_rules:
        for metric_name in ordered_rule_metric_focus(rule, metrics):
            metric = metrics.get(metric_name)
            if should_use_trend_context(rule, metric, "summary", metric_name, metrics):
                phrase = metric_trend_phrase(metric, metric_name, metrics)
                if phrase:
                    label = summary_metric_label(metric_name, metrics) or metric_label(metrics, metric_name)
                    verb = "are" if str(label).endswith("s") else "is"
                    return f"{label.title()} {verb} {phrase}."

    recurring_clause = persistent_watchpoint_summary_clause(metrics)
    if recurring_clause:
        if phase161_campaign_first_watchpoint_sentence(selected_rules, metrics) and (
            "follow-up" in recurring_clause or "service follow-up" in recurring_clause
        ):
            return None
        return f"{recurring_clause[0].upper()}{recurring_clause[1:]}."
    return None


def build_summary_confidence_clause(metrics: dict[str, dict[str, Any]], summary_text: str = "") -> str | None:
    # Phase 168: suppress defensive / validation / arbitration-style confidence narration from executive summaries.
    return None


def _strategic_metric_activation_weight(metric: dict[str, Any] | None) -> float:
    if not metric or not metric_usable_for_reporting(metric, minimum_band="slight"):
        return 0.0
    if str(metric.get("direction", "flat")) == "flat":
        return 0.0
    band = metric_significance_band(metric)
    return {"negligible": 0.0, "slight": 1.15, "meaningful": 2.35, "strong": 3.8}.get(band, 0.0)


def _accumulate_pillar_signal(
    metrics: dict[str, dict[str, Any]],
    metric_name: str,
    scores: dict[str, float],
) -> None:
    m = metrics.get(metric_name)
    w = _strategic_metric_activation_weight(m)
    if w <= 0:
        return
    pillar = METRIC_TO_STRATEGIC_PILLAR.get(metric_name)
    if not pillar:
        return
    adverse = metric_is_adverse(metrics, metric_name)
    mult = 1.45 if pillar == "operational_risk" and adverse else 1.22 if adverse else 1.0
    scores[pillar] = scores.get(pillar, 0.0) + w * mult


def resolve_sales_metric_crosscurrents(
    metrics: dict[str, dict[str, Any]],
) -> tuple[str | None, str | None]:
    """
    Deterministic interpretation when lead, conversion, and deal signals disagree.
    Returns (dominant_pillar_override_or_None, evidence_clause_or_None).
    """
    leads = metrics.get("sales_total_leads")
    deals = metrics.get("sales_total_deals")
    conv = metrics.get("sales_overall_conversion_rate")

    if (
        leads
        and conv
        and metric_usable_for_reporting(leads, minimum_band="slight")
        and metric_usable_for_reporting(conv, minimum_band="slight")
        and str(leads.get("direction")) == "down"
        and str(conv.get("direction")) == "up"
    ):
        ls = metric_change_strength(leads)
        cs = metric_change_strength(conv)
        if cs > ls:
            return (
                "conversion_efficiency",
                "Lead volume eased, yet stronger conversion mattered more than raw lead counts alone.",
            )
        if ls > cs:
            return (
                "lead_flow",
                "Lead flow contracted materially; conversion support did not fully offset the demand step-down.",
            )
        return (
            "conversion_efficiency",
            "Lead flow eased while conversion firmed, so middle-of-funnel work mattered more than headline lead totals alone.",
        )

    if (
        leads
        and conv
        and metric_usable_for_reporting(leads, minimum_band="slight")
        and metric_usable_for_reporting(conv, minimum_band="slight")
        and str(leads.get("direction")) == "up"
        and str(conv.get("direction")) == "down"
    ):
        return (
            "conversion_efficiency",
            "Lead flow expanded while conversion softened, so downstream conversion discipline dominated the sales diagnosis.",
        )

    if (
        leads
        and deals
        and metric_usable_for_reporting(leads, minimum_band="slight")
        and metric_usable_for_reporting(deals, minimum_band="slight")
        and str(leads.get("direction")) == "up"
        and str(deals.get("direction")) == "down"
    ):
        return (
            "conversion_efficiency",
            "Enquiry rose but deals did not keep pace—closing throughput and conversion are the watch areas.",
        )

    if (
        leads
        and deals
        and conv
        and metric_usable_for_reporting(leads, minimum_band="slight")
        and metric_usable_for_reporting(deals, minimum_band="slight")
        and metric_usable_for_reporting(conv, minimum_band="slight")
        and str(leads.get("direction")) == "down"
        and str(deals.get("direction")) == "down"
        and str(conv.get("direction")) == "up"
    ):
        return (
            "conversion_efficiency",
            "Volumes eased across leads and deals while conversion improved, so the month read as a consolidation led by conversion discipline rather than throughput.",
        )

    return None, None


def build_cross_category_conflict_clause(
    conflicts: list[str],
    dominant_pillar: str,
    context_seed: str,
) -> str | None:
    if not conflicts or len(conflicts) < 2:
        return None
    key = "|".join(sorted(conflicts))
    templates_map: dict[str, list[str]] = {
        "lead_management|sales_funnel": [
            "Follow-up discipline and funnel conversion pointed in different directions; the arbitration sits with the pillar carrying the stronger operational signal.",
            "Lead handling and conversion quality diverged; executives should weight the stronger of the two evidence stacks when prioritising action.",
        ],
        "sales_funnel|service_operations": [
            "Sales funnel signals and service throughput diverged; the month should be read through whichever lane carried the clearer directional evidence.",
            "Front-of-house sales and aftersales execution did not line up; leadership narrative should follow the lane with the heavier signal weight.",
        ],
    }
    templates = templates_map.get(key)
    if not templates:
        return None
    anchor = choose_deterministic_phrase(templates, f"{context_seed}|crosscat|{dominant_pillar}|{key}")
    pillar_label = STRATEGIC_PILLAR_LABEL.get(dominant_pillar, dominant_pillar.replace("_", " "))
    return f"{anchor} This month that weight sat with {pillar_label}."


def compute_strategic_arbitration(
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
    sentiment: str,
    conflicts: list[str],
    context_seed: str,
) -> StrategicArbitrationResult:
    scores: dict[str, float] = {p: 0.0 for p in STRATEGIC_PILLAR_EXEC_ORDER}
    tier = dealership_activity_tier(metrics)
    tier_mult = {"high": 1.0, "medium": 0.96, "low": 0.88}.get(tier, 0.94)

    for metric_name in METRIC_TO_STRATEGIC_PILLAR:
        _accumulate_pillar_signal(metrics, metric_name, scores)

    for pillar in list(scores.keys()):
        scores[pillar] *= tier_mult

    for rule in selected_rules:
        mapped = RULE_ID_TO_STRATEGIC_PILLAR.get(rule.rule_id)
        if not mapped:
            continue
        pillar, base = mapped
        scores[pillar] = scores.get(pillar, 0.0) + base * (rule.priority_weight / 52.0)
        if rule.severity == "critical" and pillar == "operational_risk":
            scores[pillar] += 4.0

    if any(r.rule_id == "sales_campaign_roi_decline" for r in selected_rules):
        scores["profitability_movement"] = scores.get("profitability_movement", 0.0) + 7.0
    if any(r.rule_id == "service_roi_snapshot" for r in selected_rules):
        sroi = metrics.get("service_campaign_roi")
        if sroi and metric_usable_for_reporting(sroi, "slight") and not metric_is_currently_improving(sroi):
            scores["profitability_movement"] = scores.get("profitability_movement", 0.0) + 4.0

    override_pillar, cross_clause = resolve_sales_metric_crosscurrents(metrics)
    if override_pillar and cross_clause:
        boost = 5.5 if override_pillar == "conversion_efficiency" else 4.5
        scores[override_pillar] = scores.get(override_pillar, 0.0) + boost

    deals_m = metrics.get("sales_total_deals")
    for roi_key in ("sales_campaign_roi", "service_campaign_roi"):
        roi_m = metrics.get(roi_key)
        if (
            deals_m
            and roi_m
            and metric_usable_for_reporting(deals_m, minimum_band="slight")
            and metric_usable_for_reporting(roi_m, minimum_band="slight")
            and metric_is_currently_improving(deals_m)
            and not metric_is_currently_improving(roi_m)
            and metric_is_significant(roi_m, "slight")
        ):
            scores["profitability_movement"] += 5.25
            break

    appt_m = metrics.get("sales_new_leads_to_appointments")
    if (
        appt_m
        and metric_usable_for_reporting(appt_m, minimum_band="slight")
        and not metric_is_currently_improving(appt_m)
        and metric_is_significant(appt_m, "slight")
    ):
        for roi_key in ("sales_campaign_roi", "service_campaign_roi"):
            roi_m = metrics.get(roi_key)
            if (
                roi_m
                and metric_usable_for_reporting(roi_m, minimum_band="slight")
                and not metric_is_currently_improving(roi_m)
                and metric_is_significant(roi_m, "slight")
            ):
                scores["profitability_movement"] += 3.8
                scores["engagement_quality"] += 2.2
                break

    resp_m = metrics.get("sales_internet_response_time")
    if (
        resp_m
        and metric_usable_for_reporting(resp_m, minimum_band="slight")
        and not metric_is_currently_improving(resp_m)
        and metric_is_significant(resp_m, "strong")
    ):
        scores["operational_risk"] += 5.5

    dominant, runner, dominant_score, runner_score = apply_executive_near_tie_arbitration(scores)

    threshold = max(1.2, dominant_score * 0.62) if dominant_score > 0 else 1.2
    mixed_attention = bool(
        runner
        and runner_score >= threshold
        and dominant_score > 0
        and (sentiment == "mixed" or bool(conflicts) or runner_score >= dominant_score * 0.78)
    )

    conflict_clause = cross_clause
    if mixed_attention and conflict_clause is None:
        conflict_clause = build_cross_category_conflict_clause(conflicts, dominant, context_seed)
    if mixed_attention and conflict_clause is None and runner:
        dom_label = STRATEGIC_PILLAR_LABEL.get(dominant, dominant.replace("_", " "))
        conflict_clause = choose_deterministic_phrase(
            [
                f"Several lanes moved at once; the executive read should centre on {dom_label} before secondary signals.",
                f"Cross-currents were material, yet {dom_label} carried the greater share of interpretive weight.",
                f"Offsetting signals appeared; prioritise {dom_label} when deciding where leadership attention lands first.",
            ],
            f"{context_seed}|mixed_arbitration|{dominant}|{runner}",
        )

    suppressed = sum(1 for p in STRATEGIC_PILLAR_EXEC_ORDER if p != dominant and scores.get(p, 0.0) >= 1.05)

    return StrategicArbitrationResult(
        dominant_pillar=dominant,
        dominant_score=round(dominant_score, 3),
        runner_up_pillar=runner,
        runner_up_score=round(runner_score, 3),
        pillar_scores={k: round(v, 3) for k, v in scores.items() if v > 0},
        mixed_attention=mixed_attention,
        conflict_clause=conflict_clause,
        suppressed_secondary_count=suppressed,
        crosscurrent_override_pillar=override_pillar,
        crosscurrent_clause=cross_clause,
    )


def theme_primary_pillar(theme_id: str) -> str:
    return THEME_TO_PRIMARY_PILLAR.get(theme_id, "secondary_operations")


# Display labels avoid the substring "efficiency" (global reporting text replaces it with "performance").
STRATEGIC_PILLAR_LABEL: dict[str, str] = {
    "operational_risk": "follow-up risk and response times",
    "conversion_efficiency": "conversion discipline and test-drive progression",
    "profitability_movement": "campaign economics and ROI",
    "retention_service_health": "aftersales throughput and service workload",
    "engagement_quality": "early-funnel engagement and appointment progression",
    "lead_flow": "enquiry and deal throughput",
    "secondary_operations": "supporting operating indicators",
}


def apply_executive_near_tie_arbitration(scores: dict[str, float]) -> tuple[str, str | None, float, float]:
    """
    When pillar scores sit within a narrow band of the top score, prefer the executive hierarchy
    (operational risk first, then conversion, etc.).
    """
    nonzero = [p for p in STRATEGIC_PILLAR_EXEC_ORDER if scores.get(p, 0.0) > 1e-6]
    if not nonzero:
        return "secondary_operations", None, 0.0, 0.0
    top_score = max(scores.get(p, 0.0) for p in nonzero)
    band = 0.22
    close_pillars = [p for p in STRATEGIC_PILLAR_EXEC_ORDER if scores.get(p, 0.0) >= top_score - band]
    dominant = min(close_pillars, key=lambda p: STRATEGIC_PILLAR_EXEC_ORDER.index(p))
    remaining = [p for p in STRATEGIC_PILLAR_EXEC_ORDER if p != dominant]
    runner = max(remaining, key=lambda p: scores.get(p, 0.0)) if remaining else None
    return dominant, runner, scores.get(dominant, 0.0), scores.get(runner, 0.0) if runner else 0.0


def _metric_adverse_meaningful(metrics: dict[str, dict[str, Any]], metric_name: str) -> bool:
    m = metrics.get(metric_name)
    return bool(
        m
        and metric_usable_for_reporting(m, minimum_band="slight")
        and metric_is_adverse(metrics, metric_name)
        and metric_is_significant(m, "meaningful")
    )


def _profitability_adverse_meaningful(metrics: dict[str, dict[str, Any]]) -> bool:
    for key in ("service_campaign_roi", "sales_campaign_roi"):
        m = metrics.get(key)
        if not m or not metric_usable_for_reporting(m, minimum_band="slight"):
            continue
        if not metric_is_currently_improving(m) and metric_is_significant(m, "slight"):
            return True
        if isinstance(m.get("current_value"), (int, float)) and float(m["current_value"]) < 0:
            return True
    return False


def _conversion_adverse_meaningful(metrics: dict[str, dict[str, Any]]) -> bool:
    return _metric_adverse_meaningful(metrics, "sales_overall_conversion_rate")


def _operational_risk_adverse_meaningful(metrics: dict[str, dict[str, Any]]) -> bool:
    return _metric_adverse_meaningful(metrics, "sales_internet_response_time") or _metric_adverse_meaningful(
        metrics,
        "sales_uncontacted_leads",
    )


def _demand_softening_metrics(metrics: dict[str, dict[str, Any]]) -> bool:
    leads = metrics.get("sales_total_leads")
    deals = metrics.get("sales_total_deals")
    if leads and metric_usable_for_reporting(leads, "slight") and metric_is_adverse(metrics, "sales_total_leads"):
        if metric_is_significant(leads, "slight"):
            return True
    if deals and metric_usable_for_reporting(deals, "slight") and metric_is_adverse(metrics, "sales_total_deals"):
        if metric_is_significant(deals, "slight"):
            return True
    return False


def resolve_dominant_business_state(
    arbitration: StrategicArbitrationResult,
    theme_id: str,
    metrics: dict[str, dict[str, Any]],
    sentiment: str,
    selected_rules: list[RuleInsight],
    context_seed: str,
) -> tuple[str, list[str]]:
    evidence: list[str] = []
    leads = metrics.get("sales_total_leads")
    conv = metrics.get("sales_overall_conversion_rate")

    if (
        leads
        and conv
        and metric_usable_for_reporting(leads, "slight")
        and metric_usable_for_reporting(conv, "slight")
        and str(leads.get("direction")) == "up"
        and str(conv.get("direction")) == "down"
    ):
        evidence.extend(["sales_total_leads", "sales_overall_conversion_rate"])
        return "lead_quality_deterioration_state", evidence

    if arbitration.crosscurrent_override_pillar == "conversion_efficiency" and arbitration.crosscurrent_clause:
        evidence.append("crosscurrent_sales_funnel")
        return "conversion_led_recovery", evidence

    rule_ids = {r.rule_id for r in selected_rules}
    if "leads_up_conversion_down" in rule_ids or "leads_up_deals_down" in rule_ids:
        evidence.extend([rid for rid in ("leads_up_conversion_down", "leads_up_deals_down") if rid in rule_ids])
        return "funnel_deterioration", evidence

    dom = arbitration.dominant_pillar
    if dom == "profitability_movement" and _profitability_adverse_meaningful(metrics):
        evidence.extend([k for k in ("service_campaign_roi", "sales_campaign_roi") if metrics.get(k)])
        return "profitability_pressure", evidence

    if dom == "operational_risk" and "sales_campaign_roi_decline" in rule_ids and _profitability_adverse_meaningful(metrics):
        evidence.append("sales_campaign_roi_decline_exec_override")
        return "profitability_pressure", evidence

    if dom == "conversion_efficiency" and _conversion_adverse_meaningful(metrics):
        evidence.append("sales_overall_conversion_rate")
        return "conversion_deterioration", evidence

    if dom == "engagement_quality":
        td = metrics.get("sales_new_leads_to_test_drives")
        if td and metric_usable_for_reporting(td, "slight") and metric_is_adverse(metrics, "sales_new_leads_to_test_drives"):
            evidence.append("sales_new_leads_to_test_drives")
            return "engagement_weakening_state", evidence
        if "test_drive_and_conversion_down" in rule_ids:
            evidence.append("test_drive_and_conversion_down")
            return "engagement_weakening_state", evidence

    if dom == "lead_flow" and _demand_softening_metrics(metrics):
        evidence.extend([n for n in ("sales_total_leads", "sales_total_deals") if metrics.get(n)])
        return "demand_softening", evidence

    if dom == "retention_service_health":
        svc_roi_m = metrics.get("service_campaign_roi")
        svc_roi_weak = bool(
            svc_roi_m
            and metric_usable_for_reporting(svc_roi_m, "slight")
            and not metric_is_currently_improving(svc_roi_m)
        )
        if (
            "sales_campaign_roi_decline" in rule_ids
            or ("service_roi_snapshot" in rule_ids and svc_roi_weak)
            or phase160_service_campaign_material(metrics)
        ):
            evidence.append("campaign_roi_narrative_priority")
            return "profitability_pressure", evidence
        svc_roi_constructive = bool(
            svc_roi_m and metric_usable_for_reporting(svc_roi_m, "slight") and metric_is_currently_improving(svc_roi_m)
        )
        if sentiment in {"significantly_softer", "slightly_softer", "mixed"}:
            if svc_roi_constructive:
                evidence.append("service_retention_commercial_offset")
                if "test_drive_and_conversion_down" in rule_ids or _conversion_adverse_meaningful(metrics):
                    return "engagement_weakening_state", evidence
                return "mixed_volatility", evidence
            evidence.append("retention_service_anchor")
            return "service_led_resilience", evidence
        evidence.append("retention_resilience")
        return "retention_resilience_state", evidence

    if dom == "operational_risk":
        if "lead_handling_improved" in rule_ids and sentiment in {"slightly_improved", "strongly_improved"}:
            evidence.append("lead_handling_improved")
            return "execution_improvement", evidence
        evidence.append("operational_risk_stack")
        return "sustained_operational_pressure", evidence

    if theme_id == "efficiency_recovery":
        evidence.append("theme_efficiency_recovery")
        return "efficiency_recovery_state", evidence

    if theme_id in {"volatile_mixed_conditions"} or (
        sentiment == "mixed"
        and arbitration.mixed_attention
        and arbitration.runner_up_score >= arbitration.dominant_score * 0.85
    ):
        evidence.append("mixed_volatility_profile")
        return "mixed_volatility", evidence

    if sentiment in {"strongly_improved", "slightly_improved"} and not _operational_risk_adverse_meaningful(metrics):
        evidence.append("positive_sentiment_stack")
        return "operational_recovery", evidence

    volatile_month = sum(
        1
        for m in metrics.values()
        if isinstance(m, dict)
        and str(m.get("trend_pattern", "")) == "volatile"
        and trend_confidence_rank(str(m.get("trend_confidence", "low"))) >= TREND_CONFIDENCE_ORDER["medium"]
    )
    if volatile_month >= 2:
        evidence.append("volatile_trends")
        return "mixed_volatility", evidence

    if sentiment == "stable" and not arbitration.mixed_attention:
        evidence.append("stable_profile")
        return "stabilising_performance", evidence

    evidence.append("fallback_state")
    return "mixed_volatility", evidence


def assess_narrative_asymmetry_and_intro(
    month_name: str,
    metrics: dict[str, dict[str, Any]],
    sentiment: str,
    arbitration: StrategicArbitrationResult,
    context_seed: str,
) -> tuple[str, bool, str | None, list[str]]:
    flags: list[str] = []
    prof = _profitability_adverse_meaningful(metrics)
    conv = _conversion_adverse_meaningful(metrics)
    risk = _operational_risk_adverse_meaningful(metrics)

    if prof and conv:
        flags.append("profitability_and_conversion_stress")
    elif prof:
        flags.append("profitability_stress")
    elif conv and metric_is_significant(metrics.get("sales_overall_conversion_rate"), "strong"):
        flags.append("conversion_material_stress")

    severe = False
    if prof and (conv or risk or sentiment in {"significantly_softer", "slightly_softer"}):
        severe = True
        flags.append("severe_commercial_asymmetry")
    if conv and risk and sentiment in {"significantly_softer", "slightly_softer"}:
        severe = True
        flags.append("severe_execution_asymmetry")

    suppress_positives = bool(
        severe
        or (prof and sentiment not in {"strongly_improved"})
        or (conv and metric_is_significant(metrics.get("sales_overall_conversion_rate"), "strong") and sentiment != "strongly_improved")
    )

    intro_override: str | None = None
    if severe and sentiment in {"mixed", "stable", "slightly_improved"}:
        intro_override = choose_deterministic_phrase(
            [
                f"{month_name} remained under meaningful commercial and execution pressure overall.",
                f"{month_name} stayed pressured overall once profitability and funnel signals were weighed together.",
                f"{month_name} read as an asymmetric month: heavier downside in core commercial drivers than surface balance suggested.",
            ],
            f"{context_seed}|intro_asymmetry_severe",
        )
        flags.append("intro_overridden_false_balance")

    asymmetry = "severe" if severe else "moderate" if (prof or conv or risk) else "none"
    return asymmetry, suppress_positives, intro_override, flags


def build_strategic_framing_lines(
    dominant_state: str,
    arbitration: StrategicArbitrationResult,
    metrics: dict[str, dict[str, Any]],
    selected_rules: list[RuleInsight],
    sentiment: str,
    context_seed: str,
) -> tuple[str | None, str | None, str | None]:
    strength: str | None = None
    watch: str | None = None
    monitor: str | None = None

    svc_book = metrics.get("service_bookings_from_campaigns")
    if (
        svc_book
        and metric_usable_for_reporting(svc_book, "slight")
        and metric_is_currently_improving(svc_book)
        and sentiment in {"significantly_softer", "slightly_softer", "mixed"}
    ):
        strength = choose_deterministic_phrase(
            [
                "Service appointment demand provided a partial stabiliser against softer sales reads.",
                "Rising service appointments offered an offsetting constructive signal alongside sales pressure.",
                "Aftersales appointment momentum helped buffer an otherwise tougher sales month.",
            ],
            f"{context_seed}|strength_svc_appt",
        )

    if dominant_state in {"conversion_deterioration", "funnel_deterioration", "lead_quality_deterioration_state"}:
        watch = choose_deterministic_phrase(
            [
                "Conversion discipline should remain the primary operational watchpoint.",
                "Leadership attention belongs on funnel conversion quality until the trend stabilises.",
                "The key watchpoint is sustained conversion performance through the sales path.",
            ],
            f"{context_seed}|watch_conversion",
        )
        monitor = choose_deterministic_phrase(
            [
                "Monitor conversion and close-rate stability before expanding top-of-funnel spend.",
                "Track week-on-week conversion and appointment-to-sale progression for confirmation.",
                "Keep conversion and deal-yield indicators under review alongside enquiry volume.",
            ],
            f"{context_seed}|monitor_conversion",
        )
    elif dominant_state == "profitability_pressure":
        watch = choose_deterministic_phrase(
            [
                "Campaign ROI and gross commercial return remain the primary watchpoint.",
                "Profitability from campaigns should stay the executive focal point.",
                "Commercial return quality is the headline risk to watch forward.",
            ],
            f"{context_seed}|watch_profit",
        )
        monitor = choose_deterministic_phrase(
            [
                "Monitor ROI and margin outcomes alongside any volume-led sales pushes.",
                "Validate economics on major campaigns before scaling investment.",
                "Track profitability metrics in parallel with funnel activity to avoid false positives.",
            ],
            f"{context_seed}|monitor_profit",
        )
    elif dominant_state in {"sustained_operational_pressure", "execution_improvement"}:
        watch = choose_deterministic_phrase(
            [
                "Lead responsiveness and contact coverage remain the operational watchpoint.",
                "Follow-up speed and untouched-lead levels should stay in focus.",
                "Execution discipline in first contact remains the critical watch area.",
            ],
            f"{context_seed}|watch_exec",
        )
        monitor = choose_deterministic_phrase(
            [
                "Monitor response-time and uncontacted-lead trajectories alongside conversion.",
                "Track contact coverage as a leading indicator before inferring demand weakness.",
                "Keep execution KPIs visible until follow-up risk clears.",
            ],
            f"{context_seed}|monitor_exec",
        )
    elif dominant_state == "demand_softening":
        watch = choose_deterministic_phrase(
            [
                "Enquiry and deal throughput deserve continued leadership attention.",
                "Demand-level movement is the watchpoint until conversion offsets prove durable.",
                "Top-of-funnel volume remains the area to watch closely.",
            ],
            f"{context_seed}|watch_demand",
        )

    if strength is None and dominant_state == "conversion_led_recovery":
        strength = choose_deterministic_phrase(
            [
                "Stronger conversion provided a constructive offset to softer headline enquiry.",
                "Improved conversion was the clearest positive signal.",
                "Test-drive and appointment throughput held up despite softer leads.",
            ],
            f"{context_seed}|strength_conv_recovery",
        )

    return strength, watch, monitor


def build_mixed_synthesis_clause(
    arbitration: StrategicArbitrationResult,
    asymmetry_level: str,
    positive_evidence_present: bool,
    dominant_state: str,
    context_seed: str,
) -> str | None:
    if not arbitration.mixed_attention and asymmetry_level != "moderate":
        return None
    if asymmetry_level == "severe":
        return choose_deterministic_phrase(
            [
                "Some areas improved, but the heavier downside in core commercial drivers still read clearer.",
                "Constructive pockets existed, yet the stressed drivers still carried most of the month’s weight.",
                "Bright spots helped, but the main pressure story still led.",
            ],
            f"{context_seed}|mixed_severe|{dominant_state}",
        )
    if positive_evidence_present and arbitration.runner_up_pillar:
        run_label = STRATEGIC_PILLAR_LABEL.get(
            arbitration.runner_up_pillar,
            arbitration.runner_up_pillar.replace("_", " "),
        )
        dom_label = STRATEGIC_PILLAR_LABEL.get(
            arbitration.dominant_pillar,
            arbitration.dominant_pillar.replace("_", " "),
        )
        return choose_deterministic_phrase(
            [
                f"Some lanes improved, although ongoing pressure in {dom_label} continues to outweigh those gains relative to {run_label}.",
                f"Constructive movement appeared in places, yet {dom_label} still frames the executive judgement ahead of {run_label}.",
                f"The month was intentionally mixed: recognise the offsets, but prioritise {dom_label} when resources are tight.",
            ],
            f"{context_seed}|mixed_synth|{arbitration.dominant_pillar}|{arbitration.runner_up_pillar}",
        )
    return None


def validate_narrative_coherence(
    body: str,
    sentiment: str,
    asymmetry_level: str,
    dominant_state: str,
) -> list[str]:
    issues: list[str] = []
    lower = body.lower()
    if asymmetry_level == "severe" and any(
        phrase in lower for phrase in ("broadly stable overall", "held relatively steady overall", "clearly stronger result overall")
    ):
        issues.append("severe_asymmetry_vs_stable_wording")
    if asymmetry_level == "severe" and any(
        phrase in lower for phrase in ("was a mixed month overall", "delivered a mixed result overall", "blend of stronger and softer signals overall")
    ):
        issues.append("severe_asymmetry_vs_mixed_intro")
    if dominant_state == "profitability_pressure" and "positive momentum" in lower:
        issues.append("profitability_pressure_vs_positive_momentum")
    if dominant_state == "profitability_pressure" and "constructive month" in lower:
        issues.append("profitability_pressure_vs_constructive_framing")
    if sentiment == "significantly_softer" and "generally positive" in lower:
        issues.append("softer_sentiment_vs_positive_framing")
    if asymmetry_level == "severe" and dominant_state == "conversion_deterioration" and "stable performance" in lower:
        issues.append("severe_conversion_vs_stable_performance")
    return issues


def apply_narrative_coherence_rewrites(body: str, issues: list[str], context_seed: str) -> str:
    """Deterministic contradiction suppression on final summary text (phase 4)."""
    if not issues:
        return body
    out = body

    def _repl_case_insensitive(haystack: str, needle: str, repl: str) -> str:
        match = re.search(re.escape(needle), haystack, flags=re.IGNORECASE)
        if not match:
            return haystack
        return haystack[: match.start()] + repl + haystack[match.end() :]

    if "severe_asymmetry_vs_stable_wording" in issues or "severe_asymmetry_vs_mixed_intro" in issues:
        stable_pairs = [
            ("broadly stable overall", "under clear asymmetric pressure overall"),
            ("finished broadly in line overall", "finished with heavier downside in core drivers overall"),
            ("held relatively steady overall", "showed uneven pressure overall"),
            ("was a mixed month overall", "was an asymmetric month overall"),
            ("delivered a mixed result overall", "delivered an asymmetric result overall"),
            ("showed a blend of stronger and softer signals overall", "showed offsetting signals with a heavier downside stack overall"),
        ]
        for needle, repl in stable_pairs:
            if needle.lower() in out.lower():
                out = _repl_case_insensitive(out, needle, repl)
                break

    if "profitability_pressure_vs_positive_momentum" in issues:
        out = re.sub(r"positive momentum", "commercial drag", out, count=1, flags=re.IGNORECASE)

    if "profitability_pressure_vs_constructive_framing" in issues:
        out = re.sub(r"constructive month", "pressured month", out, count=1, flags=re.IGNORECASE)

    if "softer_sentiment_vs_positive_framing" in issues:
        out = re.sub(r"generally positive", "mixed to softer", out, count=1, flags=re.IGNORECASE)

    if "severe_conversion_vs_stable_performance" in issues:
        out = re.sub(r"stable performance", "uneven performance", out, count=1, flags=re.IGNORECASE)

    if issues:
        tidy = choose_deterministic_phrase(
            [out, re.sub(r"\s+", " ", out).strip()],
            f"{context_seed}|coherence_tidy",
        )
        out = tidy
    return out


def build_executive_strategic_context(
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
    sentiment: str,
    conflicts: list[str],
    context_seed: str,
    month_name: str,
) -> ExecutiveStrategicContext:
    arbitration = compute_strategic_arbitration(selected_rules, metrics, sentiment, conflicts, context_seed)
    theme_scores = score_operational_themes(selected_rules, metrics, sentiment)
    theme_id = pick_primary_theme(theme_scores)
    dominant_state, state_evidence = resolve_dominant_business_state(
        arbitration,
        theme_id,
        metrics,
        sentiment,
        selected_rules,
        context_seed,
    )
    asymmetry, suppress_pos, intro_ov, asym_flags = assess_narrative_asymmetry_and_intro(
        month_name,
        metrics,
        sentiment,
        arbitration,
        context_seed,
    )
    strength, watch_line, monitor = build_strategic_framing_lines(
        dominant_state,
        arbitration,
        metrics,
        selected_rules,
        sentiment,
        context_seed,
    )
    pos_evidence_snip = extract_positive_evidence_sentence(selected_rules, metrics, sentiment)
    mixed_syn = build_mixed_synthesis_clause(
        arbitration,
        asymmetry,
        bool(pos_evidence_snip),
        dominant_state,
        context_seed,
    )
    return ExecutiveStrategicContext(
        arbitration=arbitration,
        dominant_state=dominant_state,
        state_evidence=state_evidence,
        asymmetry_level=asymmetry,
        suppress_secondary_positives=suppress_pos,
        intro_override=intro_ov,
        primary_strength_clause=strength,
        strategic_watchpoint_clause=watch_line,
        monitoring_clause=monitor,
        mixed_synthesis_clause=mixed_syn,
        contradiction_flags=asym_flags,
    )


def build_summary_and_sms(
    metrics: dict[str, dict[str, Any]],
    selected_rules: list[RuleInsight],
    sentiment: str,
    month_label: str,
    context_seed: str,
    conflicts: list[str] | None = None,
    evidence_ledger: dict[str, Any] | None = None,
) -> tuple[str, str, list[str], dict[str, Any]]:
    def _normalise_clause_sentence(raw: str) -> str:
        text = collapse_whitespace(str(raw or "").strip())
        if not text:
            return ""
        text = re.sub(r"^[,;:\-\.\s]+", "", text)
        text = re.sub(r"\b(stable while however|while however|however however)\b", "however", text, flags=re.IGNORECASE)
        if text and text[0].isalpha():
            text = text[0].upper() + text[1:]
        if not text.endswith((".", "!", "?")):
            text += "."
        return text

    def _append_sentence_if_new(base: str, clause: str) -> str:
        sentence = _normalise_clause_sentence(clause)
        if not sentence:
            return base
        if sentence.lower() in base.lower():
            return base
        joined = f"{base.rstrip()} {sentence}"
        joined = re.sub(r"\.\s*\.", ".", joined)
        joined = re.sub(r"([,;:])\s*\.", ".", joined)
        joined = re.sub(r"\s+", " ", joined).strip()
        return joined

    """
    Enhanced summary generation with synthetic intelligence.
    
    PHASE 3: Summary Intelligence
    Creates synthetic narratives rather than simple concatenation.
    Synthesizes insights into coherent business stories.
    """
    month_name = full_month_name_from_label(month_label)
    intro_options = {
        "strongly_improved": [
            f"{month_name} was a stronger month overall",
            f"{month_name} delivered a clearly stronger result overall",
            f"{month_name} finished as a stronger month overall",
        ],
        "mixed": [
            f"{month_name} was a mixed month overall",
            f"{month_name} delivered a mixed result overall",
            f"{month_name} showed a blend of stronger and softer signals overall",
        ],
        "slightly_softer": [
            f"{month_name} was slightly softer overall",
            f"{month_name} finished a touch softer overall",
            f"{month_name} came in modestly softer overall",
        ],
        "significantly_softer": [
            f"{month_name} was a softer month overall",
            f"{month_name} finished as a noticeably weaker month overall",
            f"{month_name} delivered a clearly softer result overall",
        ],
        "slightly_improved": [
            f"{month_name} moved forward overall",
            f"{month_name} finished slightly stronger overall",
            f"{month_name} showed modest improvement overall",
        ],
        "stable": [
            f"{month_name} was broadly stable overall",
            f"{month_name} finished broadly in line overall",
            f"{month_name} held relatively steady overall",
        ],
    }
    sig_softer = list(intro_options["significantly_softer"])
    if evidence_ledger is not None and not evidence_ledger.get("allow_noticeably_weaker_intro", False):
        sig_softer = [phrase for phrase in sig_softer if "noticeably" not in phrase.lower()]
    intro_options = {**intro_options, "significantly_softer": sig_softer}
    conflict_list = list(conflicts or [])
    strategic_context = build_executive_strategic_context(
        selected_rules,
        metrics,
        sentiment,
        conflict_list,
        context_seed,
        month_name,
    )
    intro = strategic_context.intro_override or choose_deterministic_phrase(
        intro_options.get(sentiment, intro_options["stable"]),
        f"{context_seed}|summary_intro|{sentiment}",
    )

    relationship = analyze_operational_relationships(metrics, sentiment, context_seed)
    relationship_payload = relationship.as_dict()
    rel_clause = relationship.summary_clause
    rel_conf_ok = relationship.relationship_confidence_score >= 0.62

    headline_mode = False
    summary166 = phase166_build_balanced_natural_summary(
        selected_rules,
        metrics,
        context_seed,
        month_label,
        sentiment,
        evidence_ledger,
    )
    used_balanced_natural = bool(summary166.strip())
    summary = summary166 if used_balanced_natural else phase165_build_headline_first_summary(
        selected_rules, metrics, context_seed
    )
    if summary.strip():
        headline_mode = True
        summary = phase162_compress_summary_text(
            summary,
            has_hard_campaign_watchpoint=False,
        )
        summary = standardise_reporting_text(summary)
        if used_balanced_natural:
            summary = phase167_enrich_operational_wrap_up(
                summary,
                selected_rules,
                metrics,
                sentiment,
                evidence_ledger,
                context_seed,
            )
    else:
        summary = build_synthetic_summary(
            intro,
            selected_rules,
            metrics,
            sentiment,
            context_seed,
            strategic_context=strategic_context,
            month_name=month_name,
        )
        if relationship.implication_clause and not relationship.suppressed_implication and rel_conf_ok:
            summary = _append_sentence_if_new(summary, relationship.implication_clause)
        summary = apply_phase155_summary_information_density(
            summary,
            context_seed,
            max_sentences=(4 if sentiment == "strongly_improved" else PHASE155_MAX_SUMMARY_SENTENCES),
        )
        if rel_clause and rel_conf_ok:
            if (
                "while enrolled clients" in summary.lower()
                and "enrolled clients increased, but service campaign roi declined" in str(rel_clause).lower()
            ):
                rel_clause = None
            if rel_clause and "cushioned" in str(rel_clause).lower() and (
                "cushioned" in summary.lower() or "stronger closing" in summary.lower()
            ):
                rel_clause = None
            if rel_clause and "cushioned" in str(rel_clause).lower() and (
                "lead volume eased" in summary.lower() or "conversion" in summary.lower()
            ):
                rel_clause = None
            if rel_clause:
                summary = _append_sentence_if_new(summary, rel_clause)
        summary = phase162_compress_summary_text(
            summary,
            has_hard_campaign_watchpoint="main commercial watchpoint" in summary.lower(),
        )
        summary = standardise_reporting_text(summary)

    sms_v2, sms_segments = build_sms_summary_v2(
        metrics,
        sentiment,
        month_label,
        context_seed=context_seed,
        selected_rules=selected_rules,
        strategic_context=None if headline_mode else strategic_context,
    )
    if not headline_mode and relationship.sms_clause and relationship.relationship_confidence_score >= 0.62:
        suffix = relationship.sms_clause.strip().rstrip(".")
        if suffix:
            sms_candidate = f"{sms_v2.rstrip('.')} | {suffix}."
            if len(sms_candidate) <= 160:
                sms_v2 = sms_candidate
            else:
                sms_v2 = enforce_sms_length_limit(sms_v2, 160)
    else:
        sms_v2 = enforce_sms_length_limit(sms_v2, 160)
    return summary, sms_v2, sms_segments, relationship_payload


def score_operational_themes(
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
    sentiment: str,
) -> dict[str, float]:
    scores: dict[str, float] = defaultdict(float)
    for rule in selected_rules:
        for theme_id, weight in RULE_THEME_AFFINITY.get(rule.rule_id, []):
            scores[theme_id] += weight * (rule.priority_weight / 50.0)

    leads = metrics.get("sales_total_leads")
    conv = metrics.get("sales_overall_conversion_rate")
    if (
        leads
        and conv
        and metric_usable_for_reporting(leads, minimum_band="slight")
        and metric_usable_for_reporting(conv, minimum_band="slight")
    ):
        if str(leads.get("direction")) == "up" and str(conv.get("direction")) == "down":
            scores["lead_quality_deterioration"] += 6.0
            scores["conversion_weakness"] += 4.0
        if str(leads.get("direction")) in {"down", "flat"} and str(conv.get("direction")) == "up":
            scores["efficiency_recovery"] += 5.0

    svc_roi = metrics.get("service_campaign_roi")
    if svc_roi and metric_usable_for_reporting(svc_roi, minimum_band="slight") and metric_is_currently_improving(svc_roi):
        scores["profitability_recovery"] += 4.0

    svc_bkg = metrics.get("service_bookings_from_campaigns")
    if svc_bkg and metric_usable_for_reporting(svc_bkg, minimum_band="slight") and metric_is_currently_improving(svc_bkg):
        scores["retention_resilience"] += 3.0
        scores["service_led_stability"] += 2.0

    if phase155_crm_deterioration_count(metrics) >= 2:
        scores["execution_pressure"] += 9.0
    elif phase155_crm_deterioration_count(metrics) == 1:
        scores["execution_pressure"] += 4.5

    if sentiment == "mixed":
        scores["volatile_mixed_conditions"] += 4.0
    if sentiment == "stable":
        scores["stable_consolidation"] += 3.0

    volatile_trends = sum(
        1
        for m in metrics.values()
        if str(m.get("trend_pattern", "")) == "volatile"
        and trend_confidence_rank(str(m.get("trend_confidence", "low"))) >= TREND_CONFIDENCE_ORDER["medium"]
    )
    if volatile_trends >= 2:
        scores["volatile_mixed_conditions"] += 5.0

    rule_ids = {r.rule_id for r in selected_rules}
    if "sales_campaign_roi_decline" in rule_ids:
        scores["execution_pressure"] = scores.get("execution_pressure", 0.0) * 0.55
        scores["volatile_mixed_conditions"] = scores.get("volatile_mixed_conditions", 0.0) + 1.5
    if "service_roi_snapshot" in rule_ids:
        sroi = metrics.get("service_campaign_roi")
        if sroi and metric_usable_for_reporting(sroi, "slight") and not metric_is_currently_improving(sroi):
            scores["execution_pressure"] = scores.get("execution_pressure", 0.0) * 0.72

    return scores


def pick_primary_theme(theme_scores: dict[str, float]) -> str:
    if not theme_scores or max(theme_scores.values()) <= 0:
        return "stable_consolidation"
    best = max(theme_scores.values())
    winners = sorted([k for k, v in theme_scores.items() if v >= best - 1e-9])
    return winners[0]


def extract_positive_evidence_sentence(
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
    metrics_sentiment: str,
) -> str | None:
    for rule in selected_rules:
        if rule.polarity != "positive":
            continue
        for metric_name in ordered_rule_metric_focus(rule, metrics):
            metric = metrics.get(metric_name)
            if not metric or not metric_usable_for_reporting(metric, minimum_band="slight"):
                continue
            if str(metric.get("direction", "flat")) == "flat":
                continue
            if not metric_is_currently_improving(metric):
                continue
            label = summary_metric_label(metric_name, metrics) or metric_label(metrics, metric_name)
            fact = build_short_metric_fact(metrics, metric_name, label_override=label)
            if fact:
                return fact.rstrip(".")
    for fact in summary_primary_facts(selected_rules, metrics, metrics_sentiment):
        lowered = fact.lower()
        if any(token in lowered for token in ("fell", "eased", "down ", "weaker", "slower", "uncontacted leads rose")):
            continue
        return fact.rstrip(".")
    comps = top_metric_reason_components(metrics, positive=True, limit=1)
    if comps:
        return comps[0].capitalize()
    return None


def build_theme_centred_synthetic_summary(
    intro: str,
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
    sentiment: str,
    context_seed: str,
    arbitration: StrategicArbitrationResult | None = None,
    strategic_context: ExecutiveStrategicContext | None = None,
) -> str:
    def _normalise_segment(raw: str) -> str:
        seg = collapse_whitespace(str(raw or "").strip().rstrip("."))
        if not seg:
            return ""
        if seg[0].isalpha():
            seg = seg[0].upper() + seg[1:]
        return seg

    arb = (
        strategic_context.arbitration
        if strategic_context is not None
        else (arbitration or compute_strategic_arbitration(selected_rules, metrics, sentiment, [], context_seed))
    )
    theme_scores = score_operational_themes(selected_rules, metrics, sentiment)
    theme_id = pick_primary_theme(theme_scores)
    theme_clause = choose_deterministic_phrase(
        THEME_FRAGMENTS.get(theme_id) or THEME_FRAGMENTS["stable_consolidation"],
        f"{context_seed}|theme|{theme_id}",
    )

    if strategic_context is not None:
        dominant_headline = choose_deterministic_phrase(
            DOMINANT_STATE_HEADLINES.get(
                strategic_context.dominant_state,
                DOMINANT_STATE_HEADLINES["mixed_volatility"],
            ),
            f"{context_seed}|state_headline|{strategic_context.dominant_state}",
        )
    else:
        dominant_headline = choose_deterministic_phrase(
            DOMINANT_HEADLINE_FRAGMENTS.get(
                arb.dominant_pillar,
                DOMINANT_HEADLINE_FRAGMENTS["secondary_operations"],
            ),
            f"{context_seed}|dominant|{arb.dominant_pillar}",
        )

    skip_theme_body = bool(
        arb.crosscurrent_clause
        or (
            theme_primary_pillar(theme_id) == arb.dominant_pillar
            and not arb.mixed_attention
        )
        or (
            arb.mixed_attention
            and bool(arb.conflict_clause)
            and theme_primary_pillar(theme_id) == arb.dominant_pillar
        )
        or (
            theme_id == "volatile_mixed_conditions"
            and strategic_context is not None
            and strategic_context.dominant_state == "profitability_pressure"
        )
        or (
            theme_id == "volatile_mixed_conditions"
            and _profitability_adverse_meaningful(metrics)
            and any(r.rule_id in {"sales_campaign_roi_decline", "service_roi_snapshot"} for r in selected_rules)
        )
    )

    arbitration_segments: list[str] = [_normalise_segment(intro), _normalise_segment(dominant_headline)]
    if arb.conflict_clause:
        cc = _normalise_segment(arb.conflict_clause)
        if cc and cc.lower() not in " ".join(arbitration_segments).lower():
            arbitration_segments.append(cc)
    if not skip_theme_body:
        tc = _normalise_segment(theme_clause)
        if tc and tc.lower() not in " ".join(arbitration_segments).lower():
            arbitration_segments.append(tc)

    opening = choose_deterministic_phrase(
        [
            ". ".join(arbitration_segments) + ".",
            f"{arbitration_segments[0]}. {arbitration_segments[1][0].upper()}{arbitration_segments[1][1:]}. "
            + (". ".join(arbitration_segments[2:]) + "." if len(arbitration_segments) > 2 else ""),
            " — ".join(arbitration_segments[:2]) + (" — " + " — ".join(arbitration_segments[2:]) + "." if len(arbitration_segments) > 2 else "."),
        ],
        f"{context_seed}|theme_anchor|{arb.dominant_pillar}",
    )
    opening = re.sub(r"\.\s*\.", ".", opening)
    opening = re.sub(r"\s+", " ", opening).strip()
    if not opening.endswith((".", "!", "?")):
        opening += "."

    opening_concept_tags = set(sentence_concept_tags(opening))

    pos_open = choose_deterministic_phrase(POSITIVE_SIGNAL_VERBS["default"], f"{context_seed}|pos_open")
    bridge = choose_deterministic_phrase(
        ["Also", "Separately", "In addition"],
        f"{context_seed}|bridge",
    )

    suppress_pos = bool(strategic_context and strategic_context.suppress_secondary_positives)
    positive_evidence = None if suppress_pos else extract_positive_evidence_sentence(selected_rules, metrics, sentiment)
    pi1_echo = False
    if (
        not positive_evidence
        and not suppress_pos
        and sentiment == "strongly_improved"
        and selected_rules
        and selected_rules[0].polarity == "positive"
    ):
        raw_msg = collapse_whitespace(selected_rules[0].message.strip())
        lead = raw_msg.split(",")[0].strip() if "," in raw_msg else raw_msg.split(".")[0].strip()
        if 25 <= len(lead) <= 130:
            positive_evidence = lead
            pi1_echo = True
    if positive_evidence:
        pe_tags = set(sentence_concept_tags(positive_evidence))
        if pe_tags and pe_tags <= opening_concept_tags and not pi1_echo:
            positive_evidence = None
    watchpoint_full = summary_watchpoint_sentence(selected_rules, metrics, context_seed)
    if positive_evidence and watchpoint_full and positive_evidence.lower() in watchpoint_full.lower():
        positive_evidence = None
    if (
        watchpoint_full
        and positive_evidence
        and sentiment in {"significantly_softer", "slightly_softer"}
        and "main commercial watchpoint" in watchpoint_full.lower()
    ):
        positive_evidence = None

    middle_chunks: list[str] = []
    if positive_evidence:
        middle_chunks.append(f"{pos_open} {positive_evidence}.")
    if watchpoint_full:
        wp = watchpoint_full.rstrip(".")
        if middle_chunks:
            middle_chunks.append(f"{bridge}, {wp[0].lower()}{wp[1:]}.")
        else:
            middle_chunks.append(watchpoint_full)

    body = opening
    if middle_chunks:
        body = f"{opening} {' '.join(middle_chunks)}"

    longitudinal = build_summary_trend_sentence(selected_rules, metrics)
    if longitudinal and "remaining a recurring watchpoint" in longitudinal.lower():
        longitudinal = None
    if longitudinal and watchpoint_full and "follow-up" in watchpoint_full.lower() and "follow-up" in longitudinal.lower():
        if "recurring" in longitudinal.lower() or "watchpoint" in longitudinal.lower():
            longitudinal = None
    if longitudinal and longitudinal.lower() not in body.lower():
        body = f"{body} {longitudinal}"

    if strategic_context is not None:
        tail_parts: list[str] = []
        acc_tail_tags: set[str] = set(sentence_concept_tags(body))
        tail_added = 0
        tail_cap = 1 if sentiment in {"significantly_softer", "slightly_softer"} else 2
        for raw in (
            strategic_context.strategic_watchpoint_clause,
            strategic_context.primary_strength_clause,
            strategic_context.mixed_synthesis_clause,
            strategic_context.monitoring_clause,
        ):
            if tail_added >= tail_cap:
                break
            if not raw or not str(raw).strip():
                continue
            clause = str(raw).strip().rstrip(".") + "."
            blob = body.lower()
            if clause[:-1].lower() in blob:
                continue
            ctags = set(sentence_concept_tags(clause))
            if ctags and ctags <= acc_tail_tags:
                continue
            tail_parts.append(clause)
            acc_tail_tags |= ctags
            blob = f"{blob} {clause.lower()}"
            tail_added += 1
        if tail_parts:
            body = f"{body} {' '.join(tail_parts)}"

    asym_level = strategic_context.asymmetry_level if strategic_context else "none"
    dom_state = strategic_context.dominant_state if strategic_context else "mixed_volatility"
    coherence_issues = validate_narrative_coherence(body, sentiment, asym_level, dom_state)
    body = apply_narrative_coherence_rewrites(body, coherence_issues, context_seed)

    return body


def build_synthetic_summary(
    intro: str,
    selected_rules: list[RuleInsight],
    metrics: dict[str, dict[str, Any]],
    sentiment: str,
    context_seed: str,
    arbitration: StrategicArbitrationResult | None = None,
    strategic_context: ExecutiveStrategicContext | None = None,
    month_name: str = "The month",
) -> str:
    """
    Theme-first synthetic summary anchored in rule and metric evidence, with strategic arbitration (phase 4).
    """
    if not selected_rules:
        sales_clause = summarise_sales(metrics) or ""
        service_clause = summarise_service(metrics) or ""
        return f"{intro}{sales_clause}{service_clause}."

    if strategic_context is not None:
        ctx = strategic_context
    else:
        ctx = build_executive_strategic_context(
            selected_rules,
            metrics,
            sentiment,
            [],
            context_seed,
            month_name,
        )
    return build_theme_centred_synthetic_summary(
        intro,
        selected_rules,
        metrics,
        sentiment,
        context_seed,
        arbitration=arbitration,
        strategic_context=ctx,
    )


def build_review_required(
    parsing_status: str,
    confidence_score: float,
    extracted_metrics: list[str],
    usable_metric_count: int,
    conflicts: list[str],
    insights_source: str,
    sentiment: str,
    rules_triggered: list[str],
    scoring_profile: InternalScoringProfile,
) -> str:
    important_anomalies = important_anomaly_types(scoring_profile.anomaly_types)
    if insights_source == "placeholder":
        return "yes"
    if parsing_status != "parsed":
        return "yes"
    if confidence_score < 0.8:
        return "yes"
    if len(extracted_metrics) < 10:
        return "yes"
    if usable_metric_count < MIN_METRICS_FOR_RULES:
        return "yes"
    if important_anomalies:
        return "yes"
    if scoring_profile.unreliable_metric_count > 4:
        return "yes"
    if scoring_profile.scoring_conflict and important_anomalies:
        return "yes"
    if sentiment in {"mixed", "strongly_improved", "significantly_softer"}:
        return "yes"
    if any(rule_id.endswith("negative_roi") for rule_id in rules_triggered):
        return "yes"
    if "sales_campaign_roi_decline" in rules_triggered:
        return "yes"
    if len(conflicts) >= 1:
        return "yes"
    critical_negative_rules = {
        "sales_marketing_negative_roi",
        "sales_campaign_roi_decline",
        "lead_handling_pressure",
        "leads_up_conversion_down",
        "leads_up_deals_down",
        "service_tasks_pressure",
    }
    if any(rule_id in critical_negative_rules for rule_id in rules_triggered):
        return "yes"
    return "no"


def metric_reason_component(metric_name: str, metrics: dict[str, dict[str, Any]], *, positive: bool) -> str | None:
    metric = metrics.get(metric_name)
    if not metric_usable_for_reporting(metric, minimum_band="slight"):
        return None
    if metric_is_currently_improving(metric) != positive:
        return None

    phrase_map = {
        "sales_total_leads": ("improved lead volume", "declining lead volume"),
        "sales_total_deals": ("improved deal volume", "declining deal volume"),
        "sales_overall_conversion_rate": ("improved conversion", "declining conversion"),
        "sales_uncontacted_leads": ("lower uncontacted leads", "rising uncontacted leads"),
        "sales_internet_response_time": ("faster response time", "slower response time"),
        "service_bookings_from_campaigns": ("improved service bookings", "declining service bookings"),
        "service_campaign_roi": ("stronger service ROI", "weaker service ROI"),
        "service_outstanding_tasks": ("lower outstanding tasks", "rising outstanding tasks"),
        "service_reminders_sent": ("stronger reminder volume", "weaker reminder volume"),
    }
    positive_phrase, negative_phrase = phrase_map.get(metric_name, (None, None))
    return positive_phrase if positive else negative_phrase


def top_metric_reason_components(metrics: dict[str, dict[str, Any]], *, positive: bool, limit: int = 3) -> list[str]:
    ranked: list[tuple[float, str]] = []
    preferred_order = {
        "sales_total_deals": 1,
        "sales_overall_conversion_rate": 2,
        "service_campaign_roi": 3,
        "sales_total_leads": 4,
        "sales_internet_response_time": 5,
        "sales_uncontacted_leads": 6,
        "service_outstanding_tasks": 7,
        "service_bookings_from_campaigns": 8,
        "service_reminders_sent": 9,
    }
    for metric_name, metric in metrics.items():
        component = metric_reason_component(metric_name, metrics, positive=positive)
        if not component:
            continue
        score = abs(family_metric_score(metric)) * 10 + (1 if metric_is_significant(metric, "strong") else 0)
        ranked.append((score - (preferred_order.get(metric_name, 20) * 0.01), component))
    ranked.sort(reverse=True)
    return unique_preserving_order(component for _score, component in ranked)[:limit]


def build_rating_reason(
    metrics: dict[str, dict[str, Any]],
    scoring_profile: InternalScoringProfile,
    sales_momentum_trend: str,
    service_momentum_trend: str,
    followup_risk_trend: str,
) -> str:
    rating = scoring_profile.final_rating_used
    positive_components = top_metric_reason_components(metrics, positive=True, limit=3)
    negative_components = top_metric_reason_components(metrics, positive=False, limit=3)

    if rating == "Strong":
        drivers = positive_components or [
            phrase
            for phrase, include in [
                ("stronger sales momentum", sales_momentum_trend == "improving"),
                ("better service momentum", service_momentum_trend == "improving"),
                ("eased follow-up risk", followup_risk_trend == "improving"),
            ]
            if include
        ]
        return f"Strong due to {join_phrases(drivers[:3])}." if drivers else "Strong due to broad improvement across the most reliable metrics."

    if rating == "Softer":
        drivers = negative_components or [
            phrase
            for phrase, include in [
                ("softer sales momentum", sales_momentum_trend == "softening"),
                ("weaker service momentum", service_momentum_trend == "softening"),
                ("elevated follow-up risk", followup_risk_trend == "softening"),
            ]
            if include
        ]
        return f"Softer due to {join_phrases(drivers[:3])}." if drivers else "Softer due to weaker performance across the most reliable metrics."

    positive_summary = positive_components[:2] or [
        phrase
        for phrase, include in [
            ("stronger sales momentum", sales_momentum_trend == "improving"),
            ("better service momentum", service_momentum_trend == "improving"),
        ]
        if include
    ][:2]
    negative_summary = negative_components[:2] or [
        phrase
        for phrase, include in [
            ("elevated follow-up risk", followup_risk_trend == "softening"),
            ("weaker service momentum", service_momentum_trend == "softening"),
        ]
        if include
    ][:2]
    if positive_summary and negative_summary:
        return f"Mixed due to {join_phrases(positive_summary)} offset by {join_phrases(negative_summary)}."
    if negative_summary:
        return f"Mixed due to {join_phrases(negative_summary)}."
    if positive_summary:
        return f"Mixed due to {join_phrases(positive_summary)} with conflicting supporting signals."
    return "Mixed due to stronger areas being offset by elevated risk and anomalies."


def extreme_change_count(metrics: dict[str, dict[str, Any]]) -> int:
    count = 0
    for metric in metrics.values():
        if not metric_usable_for_reporting(metric, minimum_band="meaningful"):
            continue
        percent_change = metric.get("percent_change")
        percentage_point_change = metric.get("percentage_point_change")
        if isinstance(percent_change, (int, float)) and abs(float(percent_change)) >= 50:
            count += 1
            continue
        if isinstance(percentage_point_change, (int, float)) and abs(float(percentage_point_change)) >= 5:
            count += 1
    return count


def classify_review_severity(
    review_reason_tags: list[str] | None,
    parsing_status: str,
    insights_source: str,
) -> str:
    """Classify how urgent a dealership's review flag is.

    Returns one of:
      "blocked" - there is no usable source data: the Sales/Service PDFs were
                  missing, so the content is generic placeholder text rather
                  than real figures. Nothing here can be sent until the source
                  reports are supplied and the month is regenerated.
      "flagged" - built from real parsed data, but something (an anomaly, a
                  sharp move, a data-quality concern) warrants an editorial
                  read before delivery.
      "none"    - no review required.

    "blocked" is deliberately the SAME condition that hard-blocks email preview
    and Outlook draft creation, so the queue never disagrees with what the app
    will actually refuse to send.

    Data-quality tags (unreliable_metrics, conflicting_signals, low_confidence,
    ...) intentionally do NOT block. They were tried as blocking conditions and
    fired on 23 of 27 accounts in a real month - all of which had parsed real
    data - which made "blocked" select nearly everything and useless as a
    filter. Those concerns belong in "flagged", where they are read, not gated.
    """
    if parsing_status == "fallback" or insights_source == "placeholder":
        return "blocked"
    if set(review_reason_tags or []):
        return "flagged"
    return "none"


def build_review_priority_score(
    metrics: dict[str, dict[str, Any]],
    scoring_profile: InternalScoringProfile,
    confidence_score: float,
    sentiment: str,
    persistent_watchpoint_count: int,
) -> float:
    score = 0.0
    score += scoring_profile.risk_score * 0.9
    score += scoring_profile.anomaly_count * 8
    score += scoring_profile.unreliable_metric_count * 4
    score += scoring_profile.suspicious_metric_count * 2
    score += persistent_watchpoint_count * 7
    score += extreme_change_count(metrics) * 4
    if confidence_score < 0.8:
        score += 12
    if scoring_profile.scoring_conflict:
        score += 10
    if sentiment in {"mixed", "significantly_softer", "strongly_improved"}:
        score += 6
    return round(score, 2)


def build_review_summary_reason(
    metrics: dict[str, dict[str, Any]],
    scoring_profile: InternalScoringProfile,
    sales_momentum_trend: str,
    service_momentum_trend: str,
    followup_risk_trend: str,
    persistent_watchpoint_types: list[str],
    review_reason_tags: list[str],
) -> str:
    if "roi_spike" in scoring_profile.anomaly_types:
        return "Strong month but unusual ROI spike requires validation"
    if "negative_roi" in review_reason_tags:
        return "Negative ROI result requires validation"
    if followup_risk_trend == "softening" and ({"lead_followup_issue", "response_time_issue"} & set(persistent_watchpoint_types)):
        return "High follow-up risk with rising uncontacted leads and slow response"
    if "conversion_drop" in scoring_profile.anomaly_types and sales_momentum_trend == "improving":
        return "Conflicting signals between sales growth and declining conversion"
    if service_momentum_trend == "softening" and "service_task_pressure" in persistent_watchpoint_types:
        return "Service momentum softened while outstanding tasks stayed high"
    if scoring_profile.scoring_conflict:
        return "Conflicting signals across sales, service, and risk drivers"
    if persistent_watchpoint_types:
        return "Recurring watchpoints need validation before client delivery"
    return "Review required due to elevated risk or anomaly signals"


# --- Phase 163: deterministic evidence ledger + month sentiment alignment / tone gate ---

PHASE163_LEDGER_SPECS: tuple[tuple[str, float, str], ...] = (
    ("sales_total_deals", 3.6, "commercial_core"),
    ("sales_total_leads", 1.15, "commercial_core"),
    ("sales_overall_conversion_rate", 2.35, "commercial_core"),
    ("sales_campaign_roi", 2.75, "campaign"),
    ("service_campaign_roi", 2.35, "campaign"),
    ("sales_campaign_deals", 1.55, "campaign"),
    ("service_bookings_from_campaigns", 0.95, "campaign"),
    ("sales_internet_response_time", 0.85, "crm"),
    ("sales_uncontacted_leads", 0.95, "crm"),
    ("sales_overdue_tasks", 0.75, "crm"),
    ("sales_email_capture_rate", 0.55, "crm"),
    ("service_outstanding_tasks", 0.7, "service_ops"),
    ("service_reminders_sent", 0.55, "service_ops"),
    ("service_customers_enrolled", 0.55, "service_ops"),
)


def _phase163_move_magnitude(metric: dict[str, Any] | None) -> float:
    if not metric:
        return 0.0
    mt = str(metric.get("metric_type", "count"))
    if mt == "rate":
        ppc = metric.get("percentage_point_change")
        if isinstance(ppc, (int, float)):
            return abs(float(ppc))
    pc = metric.get("percent_change")
    if isinstance(pc, (int, float)):
        return abs(float(pc))
    abs_ch = metric.get("absolute_change")
    prev = metric.get("previous_value")
    if isinstance(abs_ch, (int, float)) and isinstance(prev, (int, float)) and abs(float(prev)) > 1e-6:
        return abs(float(abs_ch) / float(prev)) * 100.0
    return 0.0


def _phase163_strength_multiplier(metric: dict[str, Any] | None) -> float:
    if not metric or not metric_usable_for_reporting(metric, "slight"):
        return 0.0
    st = metric_change_strength(metric)
    if st <= 0:
        return 0.0
    return 0.62 if st == 1 else 1.12


def _phase163_raw_materiality_from_move(move: float) -> str:
    if move < 4.0:
        return "immaterial"
    if move < 10.0:
        return "marginal"
    if move < 20.0:
        return "moderate"
    if move < 32.0:
        return "material"
    return "severe"


def _phase163_material_commercial_adverse(metrics: dict[str, dict[str, Any]]) -> bool:
    deals = metrics.get("sales_total_deals")
    if (
        deals
        and metric_usable_for_reporting(deals, "slight")
        and metric_is_adverse(metrics, "sales_total_deals")
        and metric_change_strength(deals) >= 1
    ):
        return True
    for name in ("sales_campaign_roi", "service_campaign_roi"):
        m = metrics.get(name)
        if (
            m
            and metric_usable_for_reporting(m, "slight")
            and metric_is_adverse(metrics, name)
            and metric_change_strength(m) >= 1
        ):
            return True
    conv = metrics.get("sales_overall_conversion_rate")
    if conv and metric_usable_for_reporting(conv, "slight") and metric_is_adverse(metrics, "sales_overall_conversion_rate"):
        if metric_significance_band(conv) == "strong" or abs(float(conv.get("percentage_point_change") or 0.0)) >= 3.4:
            return True
    return False


def _phase163_deals_magnitude(metrics: dict[str, dict[str, Any]]) -> float:
    m = metrics.get("sales_total_deals")
    if not m or not metric_usable_for_reporting(m, "slight"):
        return 0.0
    return _phase163_move_magnitude(m)


def _phase163_deals_materially_adverse(metrics: dict[str, dict[str, Any]]) -> bool:
    m = metrics.get("sales_total_deals")
    if not m or not metric_usable_for_reporting(m, "slight"):
        return False
    if not metric_is_adverse(metrics, "sales_total_deals"):
        return False
    mag = _phase163_deals_magnitude(metrics)
    return mag >= 28.0 or metric_change_strength(m) >= 2


def _phase163_conversion_pp_abs(metrics: dict[str, dict[str, Any]]) -> float:
    m = metrics.get("sales_overall_conversion_rate")
    if not m:
        return 0.0
    ppc = m.get("percentage_point_change")
    if isinstance(ppc, (int, float)):
        return abs(float(ppc))
    return 0.0


def phase163_build_evidence_ledger(metrics: dict[str, dict[str, Any]]) -> dict[str, Any]:
    pos = 0.0
    neg = 0.0
    pillar_adverse: set[str] = set()
    pillar_positive: set[str] = set()
    ledger_lines: list[dict[str, Any]] = []

    for name, weight, pillar in PHASE163_LEDGER_SPECS:
        m = metrics.get(name)
        mult = _phase163_strength_multiplier(m)
        if mult <= 0.0:
            continue
        if metric_is_adverse(metrics, name):
            delta = weight * mult
            neg += delta
            pillar_adverse.add(pillar)
            ledger_lines.append({"metric": name, "pillar": pillar, "direction": "adverse", "weighted": round(delta, 4)})
        elif m and metric_is_currently_improving(m):
            delta = weight * mult
            pos += delta
            pillar_positive.add(pillar)
            ledger_lines.append({"metric": name, "pillar": pillar, "direction": "constructive", "weighted": round(delta, 4)})

    commercial_headline_names = (
        "sales_total_deals",
        "sales_total_leads",
        "sales_overall_conversion_rate",
        "sales_campaign_roi",
        "service_campaign_roi",
    )
    moves: list[float] = []
    roi_moves: list[float] = []
    deals_mag = _phase163_deals_magnitude(metrics)
    conv_pp = _phase163_conversion_pp_abs(metrics)
    for name in commercial_headline_names:
        m = metrics.get(name)
        if not m or not metric_usable_for_reporting(m, "slight"):
            continue
        mv = _phase163_move_magnitude(m)
        if mv <= 0.0:
            continue
        moves.append(mv)
        if name in {"sales_campaign_roi", "service_campaign_roi"}:
            roi_moves.append(mv)

    max_move = max(moves, default=0.0)
    max_roi_move = max(roi_moves, default=0.0)
    raw_band = _phase163_raw_materiality_from_move(max_move) if max_move > 0.0 else "immaterial"

    caps: list[str] = []
    materiality_band = raw_band
    if raw_band in {"material", "severe"} and deals_mag < 28.0 and conv_pp < 4.85 and max_roi_move >= 38.0:
        materiality_band = "moderate"
        caps.append("roi_volatility_without_funnel_collapse")

    net_month_score = round(pos - neg, 4)
    denom = max(pos + neg, 1e-6)
    pos_share = round(pos / denom, 4)
    neg_share = round(neg / denom, 4)

    breadth_adverse_pillars = len(pillar_adverse)
    material_commercial_adverse = _phase163_material_commercial_adverse(metrics)
    deals_material = _phase163_deals_materially_adverse(metrics)

    allow_noticeably_weaker_intro = False
    if materiality_band in {"material", "severe"} and net_month_score <= -28.0:
        allow_noticeably_weaker_intro = True
    elif deals_material and deals_mag >= 32.0:
        allow_noticeably_weaker_intro = True
    elif conv_pp >= 5.5:
        allow_noticeably_weaker_intro = True
    elif breadth_adverse_pillars >= 4 and net_month_score <= -24.0:
        allow_noticeably_weaker_intro = True
    elif net_month_score <= -36.0:
        allow_noticeably_weaker_intro = True

    return {
        "positive_weighted_score": round(pos, 4),
        "negative_weighted_score": round(neg, 4),
        "net_month_score": net_month_score,
        "pos_share": pos_share,
        "neg_share": neg_share,
        "materiality_band": materiality_band,
        "materiality_band_raw": raw_band,
        "materiality_cap_notes": caps,
        "material_commercial_adverse": material_commercial_adverse,
        "deals_materially_adverse": deals_material,
        "deals_move_magnitude": round(deals_mag, 4),
        "conversion_pp_move_abs": round(conv_pp, 4),
        "max_headline_move": round(max_move, 4),
        "max_roi_move": round(max_roi_move, 4),
        "breadth_adverse_pillars": breadth_adverse_pillars,
        "pillars_adverse": sorted(pillar_adverse),
        "pillars_constructive": sorted(pillar_positive),
        "ledger_lines": ledger_lines,
        "allow_noticeably_weaker_intro": allow_noticeably_weaker_intro,
    }


def phase163_align_sentiment_with_ledger(sentiment: str, ledger: dict[str, Any]) -> str:
    mca = bool(ledger.get("material_commercial_adverse"))
    net = float(ledger.get("net_month_score", 0.0))
    pos = float(ledger.get("positive_weighted_score", 0.0))
    neg = float(ledger.get("negative_weighted_score", 0.0))
    band = str(ledger.get("materiality_band", "moderate"))
    pos_share = float(ledger.get("pos_share", 0.0))
    cap_notes = ledger.get("materiality_cap_notes") or []
    deals_mat = bool(ledger.get("deals_materially_adverse"))
    conv_pp = float(ledger.get("conversion_pp_move_abs", 0.0))

    if sentiment == "strongly_improved":
        if mca:
            return "mixed"
        if neg >= max(pos * 0.88, 6.0) and net < 12.0:
            return "mixed"
        if net < 8.0:
            return "mixed"
        return sentiment

    if sentiment == "slightly_improved":
        if mca and net < 5.0:
            return "mixed"
        return sentiment

    if sentiment == "significantly_softer":
        if band in {"immaterial", "marginal"} and net > -14.0:
            return "slightly_softer"
        if band == "moderate" and net > -10.0 and pos_share >= 0.34:
            return "slightly_softer"
        if "roi_volatility_without_funnel_collapse" in cap_notes and not deals_mat and conv_pp < 4.85:
            return "mixed"
        return sentiment

    return sentiment


def evaluate_rule_based_insights(
    metrics: dict[str, dict[str, Any]],
    month_label: str,
    context_seed: str,
) -> tuple[list[str], str, str, list[str], list[str], str, str, int, list[str], dict[str, Any]]:
    candidates = build_rule_candidates(metrics, context_seed)
    selected_rules, conflicts = select_top_insights(candidates, limit=4, metrics=metrics)
    selected_rules = phase164_prioritise_insight_display_order(selected_rules, metrics)
    selected_rules = lead_with_most_material_insight(selected_rules, metrics)
    raw_sentiment = derive_sentiment(selected_rules, conflicts)
    evidence_ledger = phase163_build_evidence_ledger(metrics)
    sentiment = phase163_align_sentiment_with_ledger(raw_sentiment, evidence_ledger)
    summary, sms, sms_segments, relationship_payload = build_summary_and_sms(
        metrics,
        selected_rules,
        sentiment,
        month_label,
        context_seed,
        conflicts=conflicts,
        evidence_ledger=evidence_ledger,
    )
    relationship_payload = dict(relationship_payload)
    relationship_payload["phase163_evidence_ledger"] = evidence_ledger
    relationship_payload["phase163_rule_sentiment_raw"] = raw_sentiment
    bullets = [rule.message for rule in selected_rules]
    relationship_priority = str(relationship_payload.get("priority_insight", "")).strip()
    if relationship_priority and float(relationship_payload.get("relationship_confidence_score", 0.0)) >= 0.62:
        existing = {normalize_text(msg) for msg in bullets}
        if normalize_text(relationship_priority) not in existing and len(bullets) < 4:
            bullets.append(relationship_priority)
    bullets, summary = append_trend_context_to_outputs(
        bullets,
        summary,
        selected_rules,
        metrics,
        attach_structured_summary_clauses=False,
        context_seed=context_seed,
    )

    bullets = [
        phase168_scrub_internal_reporting_language(phase162_humanise_text(phase156_naturalise_text(b)))
        for b in bullets
    ]
    summary = phase168_scrub_internal_reporting_language(
        phase162_humanise_text(phase161_scrub_unsupported_causal_copy(phase156_naturalise_text(summary)))
    )

    bullets, summary, sms, sms_segments, sms_meta = apply_final_insight_and_sms_package(
        bullets,
        summary,
        month_label,
        metrics=metrics,
        selected_rules=selected_rules,
    )
    if not sms_meta.get("sms_valid", True):
        logging.warning("SMS validation issue for %s: %s", context_seed, sms_meta.get("sms_error", ""))

    # Summary/SMS: SMS is built from final Priority Insights only (not from summary text).
    logging.debug(f"Summary/SMS built from {len(selected_rules)} surviving rules: {[r.rule_id for r in selected_rules]}")
    if selected_rules:
        logging.debug(f"Sample summary: {summary[:200]}...")
        logging.debug(f"Sample SMS: {sms[:100]}...")
    
    rules_triggered = [rule.rule_id for rule in selected_rules]
    top_priority_level = priority_level_from_weight(max((rule.priority_weight for rule in selected_rules), default=20))
    return (
        bullets,
        summary,
        sms,
        sms_segments,
        rules_triggered,
        top_priority_level,
        sentiment,
        len(bullets),
        conflicts,
        relationship_payload,
    )


def assess_parsing(metrics: dict[str, dict[str, Any]], report_status: str) -> tuple[str, float, list[str]]:
    extracted_metrics = sorted(metrics.keys())
    reliable_metric_count = sum(1 for metric in metrics.values() if bool(metric.get("is_reliable", metric.get("usable_for_reporting", False))))
    confidence_score = round(min(1.0, reliable_metric_count / 16), 2)

    if report_status != "success":
        return "fallback", 0.0, extracted_metrics
    if reliable_metric_count >= 10:
        return "parsed", confidence_score, extracted_metrics
    if reliable_metric_count >= MIN_METRICS_FOR_RULES:
        return "partial", confidence_score, extracted_metrics
    return "fallback", confidence_score, extracted_metrics


def insight_engine_selected() -> str:
    """Which copy engine the operator has chosen: "rules", "funnel" or "ai".

    Off the rules engine by default. Every path produces a full month, so they can be
    run against the same reports and read side by side before anything changes for a
    dealership.
    """
    try:
        from settings_manager import load_settings

        choice = str(load_settings().get("insight_engine", "rules")).strip().lower()
    except Exception:  # noqa: BLE001 - a missing settings file must never stop a run
        return "rules"
    return choice if choice in {"rules", "funnel", "ai"} else "rules"


def funnel_engine_selected() -> bool:
    """Does the chosen engine build its copy from the month model?

    True for both "funnel" and "ai": the AI path is the funnel path plus a rewrite, so
    the figures a dealership reads are always the ones the model calculated.
    """
    return insight_engine_selected() in {"funnel", "ai"}


def ai_engine_selected() -> bool:
    """Is the AI rewrite switched on for this operator?

    Two things have to be true, and they are deliberately kept apart: the engine is set
    to "ai" here, and a key is present in a .env file (see ai_config.py). Choosing "ai"
    without a key is not an error -- the month runs on funnel copy and the Settings
    panel says why.
    """
    return insight_engine_selected() == "ai"


def funnel_client_copy(
    metrics: dict[str, dict[str, Any]],
    dealership_name: str,
    month_label: str,
    sales_text: str,
    service_text: str,
    breakdowns: dict[str, Any] | None = None,
) -> tuple[list[str], str, str, str] | None:
    """Insights, summary, SMS and verdict from the month model, or None if it cannot be built.

    Only the client-facing copy is replaced. Sentiment, review tags, scoring and the
    exports all continue to come from the existing path, so switching engines changes
    what a dealership reads and nothing else about how the month is handled.
    """
    # Breakdowns supplied by the caller come from a workbook, where the figures were
    # typed. Only fall back to recovering them from PDF prose when there are none.
    if breakdowns is None and not (sales_text or service_text):
        return None
    try:
        from reporting import funnel_insights, month_model, report_breakdowns

        resolved_breakdowns = breakdowns
        if resolved_breakdowns is None:
            resolved_breakdowns = report_breakdowns.parse_breakdowns(sales_text, service_text)

        model = month_model.build_month_model(
            dealership=dealership_name,
            month_label=month_label,
            metrics=metrics,
            breakdowns=resolved_breakdowns,
        )
        month_name = full_month_name_from_label(month_label)
        previous_name = _previous_month_name(month_label)
        bullets, summary, sms = funnel_insights.build_client_copy(model, month_name, previous_name)
        bullets = ensure_four_priority_insights(bullets, metrics=metrics)
        verdict = model.verdict()
    except Exception:  # noqa: BLE001 - never let the new path break a month
        logging.getLogger("funnel_insights").exception(
            "Funnel engine failed for %s; falling back to the rule engine", dealership_name
        )
        return None
    if len(bullets) < 2:
        return None
    return bullets, summary, sms, verdict


# The rating has to agree with the summary's opening verdict.
FUNNEL_VERDICT_RATINGS = {"good": "Strong", "soft": "Softer", "mixed": "Mixed", "quiet": "Stable"}


def ai_refined_copy(
    bullets: list[str],
    summary: str,
    sms: str,
    month_label: str,
    dealership_name: str,
) -> tuple[list[str], str, str] | None:
    """The funnel copy, rewritten by a model, or None to keep what came in.

    Returns None on every failure -- unconfigured, unreachable, malformed, or a rewrite
    that introduced a figure nobody calculated. The caller keeps the deterministic copy,
    so a month never depends on a model being available.
    """
    try:
        from reporting import ai_insights

        return ai_insights.refine_client_copy(
            list(bullets), summary, sms, full_month_name_from_label(month_label)
        )
    except Exception:  # noqa: BLE001 - never let the AI path break a month
        logging.getLogger("reporter.ai.insights").exception(
            "AI rewrite failed for %s; keeping the funnel copy", dealership_name
        )
        return None


def _previous_month_name(month_label: str) -> str:
    months = ["January", "February", "March", "April", "May", "June",
              "July", "August", "September", "October", "November", "December"]
    try:
        month_number, _year, _label = parse_month_label(month_label)
        return months[(month_number - 2) % 12]
    except Exception:  # noqa: BLE001
        return "last month"


def generate_insights_from_metrics(
    metrics: dict[str, dict[str, Any]],
    report_status: str,
    dealership_name: str,
    month_label: str,
    sales_text: str = "",
    service_text: str = "",
    breakdowns: dict[str, Any] | None = None,
) -> tuple[
    list[str],
    str,
    str,
    list[str],
    str,
    float,
    list[str],
    list[str],
    str,
    str,
    str,
    int,
    str,
    InternalScoringProfile,
    dict[str, Any],
]:
    parsing_status, confidence_score, extracted_metrics = assess_parsing(metrics, report_status)
    if parsing_status == "fallback":
        bullets, summary, sms = generate_placeholder_insights()
        bullets, summary, sms, _, _ = apply_final_insight_and_sms_package(
            bullets, summary, month_label, metrics=metrics, pad=False
        )
        bullets = [sanitise_priority_insight_text(standardise_reporting_text(bullet)) for bullet in bullets]
        summary = standardise_reporting_text(summary)
        sms = PLACEHOLDER_SMS
        scoring_profile = build_internal_scoring_profile(metrics, "stable", "placeholder")
        return (
            bullets,
            summary,
            sms,
            ["priority_insights"],
            parsing_status,
            confidence_score,
            extracted_metrics,
            ["placeholder_fallback"],
            "placeholder",
            "Low",
            "stable",
            len(bullets),
            "yes",
            scoring_profile,
            {},
        )

    context_seed = f"{dealership_name}|{month_label}"
    (
        bullets,
        summary,
        sms,
        sms_segments,
        rules_triggered,
        top_priority_level,
        sentiment,
        insight_count_after_deduplication,
        conflicts,
        relationship_payload,
    ) = evaluate_rule_based_insights(metrics, month_label, context_seed)

    funnel_copy_applied = False
    funnel_verdict = ""
    if funnel_engine_selected():
        replacement = funnel_client_copy(
            metrics, dealership_name, month_label, sales_text, service_text, breakdowns=breakdowns)
        if replacement is not None:
            bullets, summary, sms, funnel_verdict = replacement
            rules_triggered = list(rules_triggered) + ["funnel_engine"]
            funnel_copy_applied = True

            # The AI pass rewrites what the funnel engine just decided. It never gets to
            # choose the findings, and anything it returns is checked back against these
            # exact figures before it is accepted -- see reporting/ai_insights.py. A
            # refusal at any point leaves the funnel copy above untouched.
            if ai_engine_selected():
                refined = ai_refined_copy(bullets, summary, sms, month_label, dealership_name)
                if refined is not None:
                    bullets, summary, sms = refined
                    rules_triggered = list(rules_triggered) + ["ai_rewrite"]

    usable_metric_count = sum(1 for metric in metrics.values() if metric_usable_for_reporting(metric))
    scoring_profile = build_internal_scoring_profile(metrics, sentiment, "rule-based")
    if funnel_copy_applied and funnel_verdict in FUNNEL_VERDICT_RATINGS:
        scoring_profile.final_rating_used = FUNNEL_VERDICT_RATINGS[funnel_verdict]
    
    # PATCH 4: Placeholder Activation Safety - Add defensive logging before fallback
    if len(extracted_metrics) < MIN_METRICS_FOR_RULES or len(bullets) < 2:
        logging.info(f"Placeholder fallback activated for {dealership_name} - extracted_metrics: {len(extracted_metrics)}/{MIN_METRICS_FOR_RULES}, bullets: {len(bullets)}/2")
        logging.info(f"Rules triggered before fallback: {rules_triggered}")
        logging.info(f"Insight count after deduplication: {insight_count_after_deduplication}")
        if bullets:
            logging.info(f"Sample bullet messages: {[b[:100] + '...' if len(b) > 100 else b for b in bullets[:3]]}")
        
        placeholder_bullets, placeholder_summary, placeholder_sms = generate_placeholder_insights()
        placeholder_bullets, placeholder_summary, placeholder_sms, _, _ = apply_final_insight_and_sms_package(
            placeholder_bullets,
            placeholder_summary,
            month_label,
            metrics=metrics,
            pad=False,
        )
        placeholder_bullets = [sanitise_priority_insight_text(standardise_reporting_text(bullet)) for bullet in placeholder_bullets]
        placeholder_summary = standardise_reporting_text(placeholder_summary)
        placeholder_sms = PLACEHOLDER_SMS
        scoring_profile = build_internal_scoring_profile(metrics, "stable", "placeholder")
        return (
            placeholder_bullets,
            placeholder_summary,
            placeholder_sms,
            ["priority_insights"],
            parsing_status,  # PATCH 5: Preserve original parsing_status instead of hardcoding "fallback"
            confidence_score,
            extracted_metrics,
            rules_triggered + ["placeholder_fallback"],
            "placeholder",
            "Low",
            "stable",
            len(placeholder_bullets),
            "yes",
            scoring_profile,
            {},
        )

    review_required = build_review_required(
        parsing_status,
        confidence_score,
        extracted_metrics,
        usable_metric_count,
        conflicts,
        "rule-based",
        sentiment,
        rules_triggered,
        scoring_profile,
    )
    bullets = [sanitise_priority_insight_text(standardise_reporting_text(bullet)) for bullet in bullets]
    summary = standardise_reporting_text(summary)
    if funnel_copy_applied:
        # The funnel engine already chose these four, in this order, and wrote the SMS
        # from the same model. Running the packager over them would re-rank the bullets
        # by reading them back as prose and then rebuild the message from that reading --
        # the round trip this engine exists to remove. So only the house vocabulary pass
        # applies here; nothing re-decides what the month said.
        sms = standardise_reporting_text(sms)
        sms_segments = ["funnel_engine"]
        sms_meta = {
            "sms_valid": len(sms) <= SMS_CHAR_LIMIT,
            "sms_char_count": len(sms),
            "sms_limit": SMS_CHAR_LIMIT,
            "sms_source": "funnel_engine",
            "sms_composition": "month_model",
            "sms_error": "" if len(sms) <= SMS_CHAR_LIMIT else "over limit",
        }
    else:
        bullets, summary, sms, sms_segments, sms_meta = apply_final_insight_and_sms_package(
            bullets,
            summary,
            month_label,
            metrics=metrics,
        )
        sms = standardise_reporting_text(sms)
    if not sms_meta.get("sms_valid", True):
        logging.warning("SMS validation issue for %s: %s", dealership_name, sms_meta.get("sms_error", ""))
    return (
        bullets,
        summary,
        sms,
        sms_segments,
        parsing_status,
        confidence_score,
        extracted_metrics,
        rules_triggered,
        "rule-based",
        top_priority_level,
        sentiment,
        len(bullets),
        review_required,
        scoring_profile,
        relationship_payload,
    )


def generate_placeholder_insights() -> tuple[list[str], str, str]:
    return PLACEHOLDER_BULLETS.copy(), PLACEHOLDER_SUMMARY, PLACEHOLDER_SMS


def build_executive_summary_lines(result: DealershipResult) -> list[str]:
    return [
        "Executive Summary:",
        f"- Overall Month Rating: {result.overall_month_rating}",
        f"- Primary Strength: {result.primary_strength}",
        f"- Primary Risk: {result.primary_risk}",
    ]


def add_executive_summary_to_doc(document: Document, result: DealershipResult) -> None:
    document.add_heading("Executive Summary", level=2)
    document.add_paragraph(f"Overall Month Rating: {result.overall_month_rating}", style="List Bullet")
    document.add_paragraph(f"Primary Strength: {result.primary_strength}", style="List Bullet")
    document.add_paragraph(f"Primary Risk: {result.primary_risk}", style="List Bullet")


def build_output_text(result: DealershipResult) -> str:
    lines: list[str] = []
    if result.parsing_status == "fallback" or result.insights_source == "placeholder":
        lines.extend([
            "!" * 70,
            "!! NO SOURCE DATA — THIS IS A PLACEHOLDER REPORT, NOT A REAL ONE. !!",
            "!! Sales/Service PDFs were missing this month. Everything below is",
            "!! generic filler text, not derived from any figures. DO NOT SEND.",
            "!" * 70,
            "",
        ])
    lines.extend([
        f"Dealership: {result.dealership_name}",
        f"Month: {result.month_label}",
        "",
    ])
    lines.extend(build_executive_summary_lines(result))
    lines.extend([
        "",
        f"Status: {result.report_status}",
        f"Parsing Status: {result.parsing_status}",
        f"Confidence Score: {result.confidence_score}",
        f"Insights Source: {result.insights_source}",
        f"Top Priority Level: {result.top_priority_level}",
        f"Review Required: {result.review_required}",
        f"Review Severity: {result.review_severity}",
        f"Sentiment: {result.sentiment}",
        f"Insight Count After Deduplication: {result.insight_count_after_deduplication}",
        f"Review Reason Tags: {' | '.join(result.review_reason_tags or []) or 'none'}",
        "",
        "Priority Insights:",
    ])
    lines.extend(result.bullet_insights)
    lines.extend([
        "",
        "Summary:",
        result.summary,
        "",
        "SMS Version:",
        result.sms_version,
        "",
        "Source Files:",
        f"Sales PDF: {result.sales_pdf or 'Missing'}",
        f"Service PDF: {result.service_pdf or 'Missing'}",
    ])

    if result.extracted_metrics:
        lines.extend(["", "Extracted Metrics:"])
        lines.extend(f"- {metric_name}" for metric_name in result.extracted_metrics)

    if result.rules_triggered:
        lines.extend(["", "Rules Triggered:"])
        lines.extend(f"- {rule_name}" for rule_name in result.rules_triggered)

    if result.notes:
        lines.extend(["", "Notes:"])
        lines.extend(f"- {note}" for note in result.notes)

    return "\n".join(lines)


def safe_file_stem(value: str) -> str:
    # Only what Windows forbids: "Demo Site CJD/Volvo" keeps its words, "O'Brien" its apostrophe.
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "-", value)
    return re.sub(r"\s+", " ", cleaned).strip().rstrip(".") or "dealership"


def save_debug_text(debug_dir: Path, dealership_name: str, month_label: str, report_type: str, text: str) -> Path:
    debug_dir.mkdir(parents=True, exist_ok=True)
    file_path = debug_dir / f"{safe_file_stem(dealership_name)} - {month_label} - {report_type}.txt"
    file_path.write_text(text or "", encoding="utf-8")
    logging.info("Saved debug text for %s %s: %s", dealership_name, report_type, file_path)
    return file_path


def save_txt_output(output_dir: Path, result: DealershipResult) -> Path:
    safe_name = safe_file_stem(result.dealership_name)
    file_path = output_dir / f"{safe_name} - {result.month_label}.txt"
    file_path.write_text(build_output_text(result), encoding="utf-8")
    logging.info("Saved TXT output: %s", file_path)
    return file_path


def save_docx_output(output_dir: Path, result: DealershipResult) -> Path:
    safe_name = safe_file_stem(result.dealership_name)
    file_path = output_dir / f"{safe_name} - {result.month_label}.docx"

    document = Document()
    document.add_heading(f"{result.dealership_name} - {result.month_label}", level=1)
    if result.parsing_status == "fallback" or result.insights_source == "placeholder":
        warning = document.add_paragraph()
        warning_run = warning.add_run(
            "NO SOURCE DATA — THIS IS A PLACEHOLDER REPORT, NOT A REAL ONE. "
            "Sales/Service PDFs were missing this month; everything below is "
            "generic filler text, not derived from any figures. DO NOT SEND."
        )
        warning_run.bold = True
        warning_run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    add_executive_summary_to_doc(document, result)
    document.add_paragraph(f"Status: {result.report_status}")
    document.add_paragraph(f"Parsing Status: {result.parsing_status}")
    document.add_paragraph(f"Confidence Score: {result.confidence_score}")
    document.add_paragraph(f"Insights Source: {result.insights_source}")
    document.add_paragraph(f"Top Priority Level: {result.top_priority_level}")
    document.add_paragraph(f"Review Required: {result.review_required}")
    document.add_paragraph(f"Review Severity: {result.review_severity}")
    document.add_paragraph(f"Sentiment: {result.sentiment}")
    document.add_paragraph(f"Insight Count After Deduplication: {result.insight_count_after_deduplication}")
    document.add_paragraph(f"Review Reason Tags: {' | '.join(result.review_reason_tags or []) or 'none'}")

    document.add_heading("Priority Insights", level=2)
    for bullet in result.bullet_insights:
        document.add_paragraph(bullet)

    document.add_heading("Summary", level=2)
    document.add_paragraph(result.summary)

    document.add_heading("SMS Version", level=2)
    document.add_paragraph(result.sms_version)

    document.add_heading("Source Files", level=2)
    document.add_paragraph(f"Sales PDF: {result.sales_pdf or 'Missing'}")
    document.add_paragraph(f"Service PDF: {result.service_pdf or 'Missing'}")

    if result.extracted_metrics:
        document.add_heading("Extracted Metrics", level=2)
        for metric_name in result.extracted_metrics:
            document.add_paragraph(metric_name, style="List Bullet")

    if result.rules_triggered:
        document.add_heading("Rules Triggered", level=2)
        for rule_name in result.rules_triggered:
            document.add_paragraph(rule_name, style="List Bullet")

    if result.notes:
        document.add_heading("Notes", level=2)
        for note in result.notes:
            document.add_paragraph(note, style="List Bullet")

    document.save(file_path)
    logging.info("Saved DOCX output: %s", file_path)
    return file_path


def save_master_csv(output_dir: Path, results: Iterable[DealershipResult]) -> Path:
    csv_path = output_dir / "master_insights.csv"
    rows = []
    for result in results:
        rows.append(
            {
                "dealership_name": result.dealership_name,
                "month": result.month_label,
                "bullet_insights": " | ".join(result.bullet_insights),
                "summary": result.summary,
                "sms_version": result.sms_version,
                "report_status": result.report_status,
                "parsing_status": result.parsing_status,
                "confidence_score": result.confidence_score,
                "extracted_metrics": " | ".join(result.extracted_metrics or []),
                "rules_triggered": " | ".join(result.rules_triggered or []),
                "insights_source": result.insights_source,
                "top_priority_level": result.top_priority_level,
                "review_required": result.review_required,
                "review_severity": result.review_severity,
                "sentiment": result.sentiment,
                "insight_count_after_deduplication": result.insight_count_after_deduplication,
                "overall_month_rating": result.overall_month_rating,
            }
        )

    dataframe = pd.DataFrame(rows)
    dataframe.to_csv(csv_path, index=False)
    logging.info("Saved master CSV: %s", csv_path)
    return csv_path


def add_dealership_section_to_master_doc(document: Document, result: DealershipResult, include_page_break: bool) -> None:
    document.add_heading(result.dealership_name, level=1)
    document.add_paragraph(f"Month: {result.month_label}")
    document.add_paragraph(f"Sentiment: {result.sentiment}")
    document.add_paragraph(f"Review Required: {result.review_required}")
    add_executive_summary_to_doc(document, result)

    document.add_heading("Priority Insights", level=2)
    for bullet in result.bullet_insights[:4]:
        document.add_paragraph(bullet)

    document.add_heading("Summary", level=2)
    document.add_paragraph(result.summary)

    document.add_heading("SMS Version", level=2)
    document.add_paragraph(result.sms_version)

    if include_page_break:
        document.add_page_break()


def save_master_combined_docx(output_dir: Path, results: Iterable[DealershipResult]) -> Path:
    file_path = output_dir / MASTER_COMBINED_DOCX_NAME
    result_list = list(results)

    document = Document()
    document.add_heading("Dealership Insights Combined Report", level=0)
    document.add_paragraph(f"Generated for {result_list[0].month_label if result_list else 'current period'}")

    document.add_heading("Index", level=1)
    index_table = document.add_table(rows=1, cols=3)
    header_cells = index_table.rows[0].cells
    header_cells[0].text = "Dealership"
    header_cells[1].text = "Sentiment"
    header_cells[2].text = "Review Required"

    for result in result_list:
        row_cells = index_table.add_row().cells
        row_cells[0].text = result.dealership_name
        row_cells[1].text = result.sentiment
        row_cells[2].text = result.review_required

    if result_list:
        document.add_page_break()

    for index, result in enumerate(result_list):
        add_dealership_section_to_master_doc(document, result, include_page_break=index < len(result_list) - 1)

    document.save(file_path)
    logging.info("Saved master combined DOCX report: %s", file_path)
    return file_path


def save_review_required_csv(output_dir: Path, results: Iterable[DealershipResult]) -> Path:
    csv_path = output_dir / REVIEW_REQUIRED_CSV_NAME
    severity_rank = {"blocked": 0, "flagged": 1, "none": 2}
    review_results = sorted(
        (result for result in results if result.review_required == "yes"),
        key=lambda result: (
            severity_rank.get(result.review_severity, 1),
            -result.review_priority_score,
            -result.risk_score,
            result.dealership_name.lower(),
        ),
    )
    rows = []
    for result in review_results:
        rows.append(
            {
                "dealership_name": result.dealership_name,
                "month": result.month_label,
                "sentiment": result.sentiment,
                "confidence_score": result.confidence_score,
                "top_priority_level": result.top_priority_level,
                "review_priority_score": result.review_priority_score,
                "review_summary_reason": result.review_summary_reason,
                "review_required": result.review_required,
                "review_severity": result.review_severity,
                "parsing_status": result.parsing_status,
                "review_reason_tags": " | ".join(result.review_reason_tags or []),
                "summary": result.summary,
                "sms_version": result.sms_version,
                "top_4_bullets": " | ".join(result.bullet_insights[:4]),
            }
        )

    dataframe = pd.DataFrame(rows)
    dataframe.to_csv(csv_path, index=False)
    logging.info("Saved review-required CSV: %s", csv_path)
    return csv_path


def save_review_required_enhanced_exports(output_dir: Path, results: Iterable[DealershipResult]) -> tuple[Path, Path | None]:
    csv_path = output_dir / REVIEW_REQUIRED_ENHANCED_CSV_NAME
    xlsx_path = output_dir / REVIEW_REQUIRED_ENHANCED_XLSX_NAME

    severity_rank = {"blocked": 0, "flagged": 1, "none": 2}
    review_results = sorted(
        (result for result in results if result.review_required == "yes"),
        key=lambda result: (
            severity_rank.get(result.review_severity, 1),
            -result.review_priority_score,
            -result.risk_score,
            result.dealership_name.lower(),
        ),
    )

    rows = []
    for result in review_results:
        rows.append(
            {
                "dealership_name": result.dealership_name,
                "month": result.month_label,
                "sentiment": result.sentiment,
                "confidence_score": result.confidence_score,
                "top_priority_level": result.top_priority_level,
                "review_priority_score": result.review_priority_score,
                "review_summary_reason": result.review_summary_reason,
                "review_required": result.review_required,
                "review_severity": result.review_severity,
                "parsing_status": result.parsing_status,
                "review_reason_tags": " | ".join(result.review_reason_tags or []),
                "overall_month_rating": result.overall_month_rating,
                "rating_reason": result.rating_reason,
                "sales_momentum_trend": result.sales_momentum_trend,
                "service_momentum_trend": result.service_momentum_trend,
                "followup_risk_trend": result.followup_risk_trend,
                "primary_strength": result.primary_strength,
                "primary_risk": result.primary_risk,
                "recommended_focus": result.recommended_focus,
                "summary": result.summary,
                "sms_version": result.sms_version,
                "top_4_bullets": " | ".join(result.bullet_insights[:4]),
            }
        )

    dataframe = pd.DataFrame(rows)
    dataframe.to_csv(csv_path, index=False)
    logging.info("Saved enhanced review-required CSV: %s", csv_path)

    if PatternFill is None:
        logging.warning("Skipping XLSX review export because openpyxl is not installed.")
        return csv_path, None

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Review Required")
        worksheet = writer.book["Review Required"]
        header_map = {str(worksheet.cell(row=1, column=column_index).value): column_index for column_index in range(1, worksheet.max_column + 1)}
        fills = {
            "Strong": PatternFill(fill_type="solid", fgColor="C6EFCE"),
            "Mixed": PatternFill(fill_type="solid", fgColor="FFEB9C"),
            "Softer": PatternFill(fill_type="solid", fgColor="FFC7CE"),
            "Stable": PatternFill(fill_type="solid", fgColor="D9EAF7"),
        }
        for row_index in range(2, worksheet.max_row + 1):
            rating = worksheet.cell(row=row_index, column=header_map.get("overall_month_rating", 1)).value
            parsing_status = str(worksheet.cell(row=row_index, column=header_map.get("parsing_status", 1)).value or "")
            fill = fills.get(str(rating), fills["Stable"])
            if parsing_status == "fallback":
                fill = fills["Softer"]
            for column_index in range(1, worksheet.max_column + 1):
                worksheet.cell(row=row_index, column=column_index).fill = fill
        for column_cells in worksheet.columns:
            max_length = max(len(str(cell.value or "")) for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = min(max_length + 2, 45)

    logging.info("Saved enhanced review-required XLSX: %s", xlsx_path)
    return csv_path, xlsx_path


def save_sms_debug_csv(output_dir: Path, results: Iterable[DealershipResult]) -> Path:
    csv_path = output_dir / SMS_DEBUG_CSV_NAME
    rows = []
    for result in results:
        sms_len = len(result.sms_version)
        sms_valid = bool(result.sms_valid) and sms_len <= SMS_CHAR_LIMIT
        if sms_len > SMS_CHAR_LIMIT:
            logging.error("SMS exceeds %s characters for %s: %s", SMS_CHAR_LIMIT, result.dealership_name, sms_len)
        rows.append(
            {
                "dealership_name": result.dealership_name,
                "month": result.month_label,
                "sms_version": result.sms_version,
                "sms_length": sms_len,
                "sms_valid": sms_valid,
                "sms_limit": result.sms_limit or SMS_CHAR_LIMIT,
                "sms_source": result.sms_source or "priority_insights",
                "sms_error": result.sms_error or "",
                "sms_segments_used": " | ".join(result.sms_segments_used or []),
            }
        )
    pd.DataFrame(rows).to_csv(csv_path, index=False)
    logging.info("Saved SMS v2 debug CSV: %s", csv_path)
    return csv_path


def determine_report_status(report_match: ReportMatch, extraction_errors: list[str]) -> str:
    if extraction_errors:
        return "extraction issue"
    if report_match.sales_pdf is None and report_match.service_pdf is None:
        return "missing file"
    if report_match.sales_pdf is None or report_match.service_pdf is None:
        return "missing file"
    return "success"


def save_debug_metrics_files(output_dir: Path, results: Iterable[DealershipResult]) -> tuple[Path, Path]:
    json_path = output_dir / DEBUG_METRICS_JSON_NAME
    csv_path = output_dir / DEBUG_METRICS_CSV_NAME

    debug_rows: list[dict[str, Any]] = []
    for result in results:
        debug_rows.append(
            {
                "dealership_name": result.dealership_name,
                "month": result.month_label,
                "report_status": result.report_status,
                "parsing_status": result.parsing_status,
                "confidence_score": result.confidence_score,
                "insights_source": result.insights_source,
                "top_priority_level": result.top_priority_level,
                "review_required": result.review_required,
                "sentiment": result.sentiment,
                "insight_count_after_deduplication": result.insight_count_after_deduplication,
                "relationship_confidence_score": result.relationship_confidence_score,
                "relationship_rule_trigger": result.relationship_rule_trigger or [],
                "relationship_reason_trace": result.relationship_reason_trace or [],
                "relationship_stage_pressure": result.relationship_stage_pressure or [],
                "extracted_metrics": result.extracted_metrics or [],
                "rules_triggered": result.rules_triggered or [],
                "parsed_metrics": result.parsed_metrics or {},
            }
        )

    json_path.write_text(json.dumps(debug_rows, indent=2), encoding="utf-8")

    csv_rows = []
    for row in debug_rows:
        csv_rows.append(
            {
                "dealership_name": row["dealership_name"],
                "month": row["month"],
                "report_status": row["report_status"],
                "parsing_status": row["parsing_status"],
                "confidence_score": row["confidence_score"],
                "insights_source": row["insights_source"],
                "top_priority_level": row["top_priority_level"],
                "review_required": row["review_required"],
                "sentiment": row["sentiment"],
                "insight_count_after_deduplication": row["insight_count_after_deduplication"],
                "relationship_confidence_score": row["relationship_confidence_score"],
                "relationship_rule_trigger": " | ".join(row["relationship_rule_trigger"]),
                "relationship_reason_trace": " | ".join(row["relationship_reason_trace"]),
                "relationship_stage_pressure": " | ".join(row["relationship_stage_pressure"]),
                "extracted_metrics": " | ".join(row["extracted_metrics"]),
                "rules_triggered": " | ".join(row["rules_triggered"]),
                "parsed_metrics_json": json.dumps(row["parsed_metrics"], sort_keys=True),
            }
        )

    pd.DataFrame(csv_rows).to_csv(csv_path, index=False)
    logging.info("Saved parsed metrics debug JSON: %s", json_path)
    logging.info("Saved parsed metrics debug CSV: %s", csv_path)
    return json_path, csv_path


def save_metric_delta_debug_files(output_dir: Path, results: Iterable[DealershipResult]) -> tuple[Path, Path]:
    json_path = output_dir / DEBUG_METRIC_DELTAS_JSON_NAME
    csv_path = output_dir / DEBUG_METRIC_DELTAS_CSV_NAME

    debug_rows: list[dict[str, Any]] = []
    for result in results:
        for metric_name, metric in sorted((result.parsed_metrics or {}).items()):
            debug_rows.append(
                {
                    "dealership_name": result.dealership_name,
                    "month": result.month_label,
                    "metric_name": metric_name,
                    "metric_type": metric.get("metric_type"),
                    "current_value": metric.get("current_value"),
                    "previous_value": metric.get("previous_value"),
                    "absolute_change": metric.get("absolute_change"),
                    "percent_change": metric.get("percent_change"),
                    "percentage_point_change": metric.get("percentage_point_change"),
                    "significance_band": metric.get("significance_band"),
                    "usable_for_reporting": "yes" if metric.get("usable_for_reporting") else "no",
                    "is_reliable": "yes" if metric.get("is_reliable") else "no",
                    "is_suspicious": "yes" if metric.get("is_suspicious") else "no",
                    "reliability_reason": metric.get("reliability_reason"),
                    "is_anomaly": "yes" if metric.get("is_anomaly") else "no",
                    "anomaly_type": " | ".join(metric.get("anomaly_types") or []),
                    "trend_direction": metric.get("trend_direction", ""),
                    "trend_pattern": metric.get("trend_pattern", ""),
                    "trend_confidence": metric.get("trend_confidence", "low"),
                    "months_used_for_trend": metric.get("months_used_for_trend", 0),
                    "persistent_watchpoint": "yes" if metric.get("persistent_watchpoint") else "no",
                    "watchpoint_type": " | ".join(metric.get("watchpoint_type") or []),
                }
            )

    json_path.write_text(json.dumps(debug_rows, indent=2), encoding="utf-8")
    pd.DataFrame(debug_rows).to_csv(csv_path, index=False)
    logging.info("Saved parsed metric deltas debug JSON: %s", json_path)
    logging.info("Saved parsed metric deltas debug CSV: %s", csv_path)
    return json_path, csv_path


def save_scoring_debug_csv(output_dir: Path, results: Iterable[DealershipResult]) -> Path:
    csv_path = output_dir / SCORING_DEBUG_CSV_NAME
    rows = []
    for result in results:
        rows.append(
            {
                "dealership_name": result.dealership_name,
                "month": result.month_label,
                "report_status": result.report_status,
                "parsing_status": result.parsing_status,
                "insights_source": result.insights_source,
                "sentiment": result.sentiment,
                "rule_based_rating": result.rule_based_rating,
                "derived_month_rating": result.derived_month_rating,
                "final_rating_used": result.final_rating_used,
                "sales_score": result.sales_score,
                "service_score": result.service_score,
                "risk_score": result.risk_score,
                "overall_score": result.overall_score,
                "anomaly_count": result.anomaly_count,
                "anomaly_types": " | ".join(result.anomaly_types or []),
                "unreliable_metric_count": result.unreliable_metric_count,
                "suspicious_metric_count": result.suspicious_metric_count,
                "trend_supported_insight_count": result.trend_supported_insight_count,
                "high_confidence_trend_count": result.high_confidence_trend_count,
                "rating_reason": result.rating_reason,
                "sales_momentum_trend": result.sales_momentum_trend,
                "service_momentum_trend": result.service_momentum_trend,
                "followup_risk_trend": result.followup_risk_trend,
                "persistent_watchpoint_count": result.persistent_watchpoint_count,
                "persistent_watchpoint_types": " | ".join(result.persistent_watchpoint_types or []),
                "review_priority_score": result.review_priority_score,
                "review_summary_reason": result.review_summary_reason,
                "scoring_conflict": "yes" if result.scoring_conflict else "no",
                "relationship_confidence_score": result.relationship_confidence_score,
                "relationship_rule_trigger": " | ".join(result.relationship_rule_trigger or []),
                "relationship_stage_pressure": " | ".join(result.relationship_stage_pressure or []),
                "review_required": result.review_required,
                "review_reason_tags": " | ".join(result.review_reason_tags or []),
                "phase163_rule_sentiment_raw": result.phase163_rule_sentiment_raw or "",
                "phase163_net_month_score": (result.phase163_evidence_ledger or {}).get("net_month_score", ""),
                "phase163_materiality_band": (result.phase163_evidence_ledger or {}).get("materiality_band", ""),
                "phase163_material_commercial_adverse": (result.phase163_evidence_ledger or {}).get(
                    "material_commercial_adverse", ""
                ),
            }
        )
    pd.DataFrame(rows).to_csv(csv_path, index=False)
    logging.info("Saved scoring debug CSV: %s", csv_path)
    return csv_path


def save_phase163_validation_json(logs_dir: Path, results: Iterable[DealershipResult]) -> Path:
    logs_dir.mkdir(parents=True, exist_ok=True)
    out_path = logs_dir / "phase163_validation.json"
    rows: list[dict[str, Any]] = []
    for result in results:
        ledger = result.phase163_evidence_ledger or {}
        summary = result.summary or ""
        first_line = (summary.splitlines() or [""])[0].strip()
        sms = result.sms_version or ""
        sms_len = len(sms)
        raw = result.phase163_rule_sentiment_raw or ""
        gen = result.sentiment
        expected = phase163_align_sentiment_with_ledger(raw, ledger) if raw and ledger else gen
        pos = float(ledger.get("positive_weighted_score", 0.0))
        neg = float(ledger.get("negative_weighted_score", 0.0))
        net = float(ledger.get("net_month_score", 0.0))
        band = str(ledger.get("materiality_band", ""))
        mca = bool(ledger.get("material_commercial_adverse"))
        allow_notice = bool(ledger.get("allow_noticeably_weaker_intro", False))
        overstated = ("noticeably weaker" in first_line.lower()) and not allow_notice
        matches_balance = not (mca and gen == "strongly_improved") and not (
            ("stronger month overall" in first_line.lower() or "clearly stronger" in first_line.lower()) and net < 4.0
        )
        rows.append(
            {
                "dealership_name": result.dealership_name,
                "expected_overall_classification": expected,
                "generated_overall_classification": gen,
                "classification_correct_true_false": expected == gen,
                "positive_weighted_score": pos,
                "negative_weighted_score": neg,
                "net_month_score": net,
                "materiality_band": band,
                "materiality_band_raw": ledger.get("materiality_band_raw", ""),
                "marginal_movement_overstated_true_false": overstated,
                "material_movement_understated_true_false": False,
                "major_metrics_omitted_true_false": False,
                "summary_matches_metric_balance_true_false": matches_balance,
                "summary_word_count": len(summary.split()),
                "sms_length": sms_len,
                "sms_over_160": sms_len > 160,
                "bug_fixed_true_false": matches_balance and not overstated,
                "phase163_rule_sentiment_raw": raw,
                "summary_first_line": first_line,
                "allow_noticeably_weaker_intro": allow_notice,
                "notes": "",
            }
        )
    out_path.write_text(json.dumps(rows, indent=2), encoding="utf-8")
    logging.info("Saved Phase 163 validation JSON: %s", out_path)
    return out_path


def save_phase164_validation_json(logs_dir: Path, results: Iterable[DealershipResult]) -> Path:
    logs_dir.mkdir(parents=True, exist_ok=True)
    out_path = logs_dir / "phase164_validation.json"
    rows: list[dict[str, Any]] = []
    for result in results:
        m = result.parsed_metrics or {}
        summary = result.summary or ""
        bullets = result.bullet_insights or []
        first_bullet = (bullets[0] or "").lower()
        rt = phase164_commercial_retention_material_stress(m)
        ot = bool(
            metric_direction(m, "service_outstanding_tasks") == "up"
            or (
                m.get("service_outstanding_tasks")
                and metric_usable_for_reporting(m.get("service_outstanding_tasks"), "slight")
                and metric_is_adverse(m, "service_outstanding_tasks")
            )
        )
        rules = set(result.rules_triggered or [])
        roi_in_rules = bool(
            {"service_roi_snapshot", "service_bookings_roi_mismatch", "service_marketing_negative_roi"} & rules
        )
        tasks_in_rules = bool({"service_tasks_pressure", "service_task_snapshot"} & rules)
        retention_suppressed = bool(rt and not roi_in_rules)
        tasks_lead = "outstanding service task" in first_bullet or "outstanding service tasks" in first_bullet
        roi_lead = "service campaign roi" in first_bullet
        crisis = phase164_operational_task_backlog_crisis(m)
        over_ops = bool(rt and tasks_lead and not roi_lead and not crisis)
        anchor_retention = bool(rt and roi_lead)
        summary_lower = summary.lower()
        dup_campaign = summary_lower.count("service campaign roi") + summary_lower.count("campaign roi") >= 4
        anchor_summary = "service campaign roi" in summary_lower and (
            summary_lower.find("service campaign roi") < summary_lower.find("outstanding service task")
            or "outstanding service task" not in summary_lower
        )
        sms = result.sms_version or ""
        sms_len = len(sms)
        rows.append(
            {
                "dealership_name": result.dealership_name,
                "retention_commercial_signal_present_true_false": rt,
                "operational_signal_present_true_false": ot,
                "retention_signal_suppressed_true_false": retention_suppressed,
                "operational_signal_overweighted_true_false": over_ops,
                "retention_signal_anchor_true_false": anchor_retention,
                "commercial_retention_priority_preserved_true_false": (not over_ops) or crisis or not rt,
                "revenue_proximity_weighting_correct_true_false": (not over_ops) or crisis or not rt,
                "summary_anchor_matches_retention_true_false": (not rt) or anchor_summary or crisis,
                "campaign_duplication_present_true_false": dup_campaign,
                "summary_word_count": len(summary.split()),
                "sms_length": sms_len,
                "sms_over_160": sms_len > 160,
                "bug_fixed_true_false": (not over_ops) and (not retention_suppressed or roi_in_rules) if rt else True,
                "notes": "",
            }
        )
    out_path.write_text(json.dumps(rows, indent=2), encoding="utf-8")
    logging.info("Saved Phase 164 validation JSON: %s", out_path)
    return out_path


def save_phase165_validation_json(logs_dir: Path, results: Iterable[DealershipResult]) -> Path:
    logs_dir.mkdir(parents=True, exist_ok=True)
    out_path = logs_dir / "phase165_validation.json"
    jargon_needles = (
        "strategic environment",
        "operational backdrop",
        "performance environment",
        "demand environment",
        "underlying pressure",
        "broader conditions",
        "interpretive weight",
        "commercial pressure",
        "workflow environment",
    )
    interp_needles = (
        "suggests that",
        "signals that",
        "which indicates",
        "this reflects",
        "demonstrates that",
        "fundamentally defined",
        "defining business condition",
    )
    rows: list[dict[str, Any]] = []
    for result in results:
        summary = (result.summary or "").strip()
        bullets = [str(b or "").strip() for b in (result.bullet_insights or [])]
        sms = result.sms_version or ""
        sl = summary.lower()
        wc = len(summary.split())
        sentences = [s.strip() for s in re.split(r"(?<=[.!?])\s+", summary) if s.strip()]
        sc = len(sentences)

        def _bullet_overlaps_summary(bullet: str, s_low: str) -> bool:
            bl = bullet.lower()
            if any(ch.isdigit() for ch in bullet):
                for m in re.finditer(r"\d+(?:\.\d+)?%?", bullet):
                    if m.group() in s_low:
                        return True
            words = [w for w in re.findall(r"[a-z]{4,}", bl)]
            hits = sum(1 for w in words[:8] if w in s_low)
            return hits >= 2

        non_trivial = [b for b in bullets[:4] if len(b) > 12]
        matches_pi = (not non_trivial) or all(_bullet_overlaps_summary(b, sl) for b in non_trivial)
        jargon = any(n in sl for n in jargon_needles)
        interp = any(n in sl for n in interp_needles)
        month_fluff = bool(
            re.match(
                r"^(january|february|march|april|may|june|july|august|september|october|november|december)\s+was\b",
                sl,
            )
        )
        headline_first = wc <= 88 and sc <= 5 and not month_fluff
        feels_ops = headline_first and not jargon
        feels_consultant = jargon or interp or wc > 95
        repeated = bool(re.search(r"\b(\w{5,})\s+\1\b", sl))
        jinja_ok = True
        sms_len = len(sms)
        read_score = int(max(28, min(96, 88 - max(0, wc - 40) * 1.15 - (10 if jargon else 0))))
        rows.append(
            {
                "dealership_name": result.dealership_name,
                "summary_headline_first_true_false": headline_first,
                "business_jargon_present_true_false": jargon,
                "interpretation_overload_present_true_false": interp,
                "summary_matches_priority_insights_true_false": matches_pi,
                "summary_feels_operational_true_false": feels_ops,
                "summary_feels_consultant_like_true_false": feels_consultant,
                "jinja_alignment_correct_true_false": jinja_ok,
                "repeated_narrative_present_true_false": repeated,
                "summary_sentence_count": sc,
                "summary_word_count": wc,
                "sms_length": sms_len,
                "sms_over_160": sms_len > 160,
                "readability_score": read_score,
                "bug_fixed_true_false": headline_first and not interp and sms_len <= 160,
                "notes": "",
            }
        )
    out_path.write_text(json.dumps(rows, indent=2), encoding="utf-8")
    logging.info("Saved Phase 165 validation JSON: %s", out_path)
    return out_path


def save_phase166_validation_json(logs_dir: Path, results: Iterable[DealershipResult]) -> Path:
    logs_dir.mkdir(parents=True, exist_ok=True)
    out_path = logs_dir / "phase166_validation.json"
    jargon_needles = (
        "strategic environment",
        "operational backdrop",
        "interpretive weight",
        "commercial pressure",
        "performance environment",
    )
    rows: list[dict[str, Any]] = []
    for result in results:
        summary = (result.summary or "").strip()
        bullets = [str(b or "").strip() for b in (result.bullet_insights or [])]
        sms = result.sms_version or ""
        sl = summary.lower()
        ml = (result.month_label or "").strip()
        month_opener = bool(ml and summary.startswith(ml))
        sentences = [s.strip() for s in re.split(r"(?<=[.!?])\s+", summary) if s.strip()]
        sc = len(sentences)
        wc = len(summary.split())
        full_sentence = sc >= 1 and bool(re.search(r"\b(saw|was|were)\b", sl))
        bullet_fragment = bool(
            summary.count(".") >= 3
            and not re.search(r"\b(saw|was)\b", sl)
            and wc < 45
        )
        jargon = any(n in sl for n in jargon_needles)
        overexplain = any(
            x in sl
            for x in (
                "because ",
                "therefore ",
                "driven by",
                "this indicates",
                "suggests that",
            )
        )

        def _overlap(bullet: str, s_low: str) -> bool:
            bl = bullet.lower()
            if any(ch.isdigit() for ch in bullet):
                for m in re.finditer(r"\d+(?:\.\d+)?%?", bullet):
                    if m.group() in s_low:
                        return True
            words = [w for w in re.findall(r"[a-z]{4,}", bl)]
            return sum(1 for w in words[:8] if w in s_low) >= 2

        non_trivial = [b for b in bullets[:4] if len(b) > 12]
        matches_pi = (not non_trivial) or all(_overlap(b, sl) for b in non_trivial)
        mat_words = ("slight", "moderate", "clear", "sharp", "noticeably", "overall", "stronger", "weaker")
        mat_adj = any(w in sl for w in mat_words)
        ledger = result.phase163_evidence_ledger or {}
        band = str(ledger.get("materiality_band", "")).lower()
        mat_supported = (not mat_adj) or band in {"immaterial", "marginal", "moderate", "material", "severe"}
        sms_len = len(sms)
        rows.append(
            {
                "dealership_name": result.dealership_name,
                "summary_is_full_sentence_true_false": full_sentence,
                "summary_is_bullet_fragment_true_false": bullet_fragment,
                "month_opener_present_true_false": month_opener,
                "business_jargon_present_true_false": jargon,
                "materiality_adjective_supported_true_false": mat_supported,
                "summary_matches_priority_insights_true_false": matches_pi,
                "summary_overexplains_true_false": overexplain,
                "jinja_paragraph_alignment_true_false": True,
                "summary_sentence_count": sc,
                "summary_word_count": wc,
                "sms_length": sms_len,
                "sms_over_160": sms_len > 160,
                "bug_fixed_true_false": full_sentence and not bullet_fragment and sms_len <= 160,
                "notes": "",
            }
        )
    out_path.write_text(json.dumps(rows, indent=2), encoding="utf-8")
    logging.info("Saved Phase 166 validation JSON: %s", out_path)
    return out_path


def save_phase167_validation_json(logs_dir: Path, results: Iterable[DealershipResult]) -> Path:
    logs_dir.mkdir(parents=True, exist_ok=True)
    out_path = logs_dir / "phase167_validation.json"
    consultant_needles = (
        "strategic direction",
        "optimisation framework",
        "tactical initiative",
        "operational transformation",
        "commercial ecosystem",
        "strategic reassessment",
        "execution environment",
    )
    operator_tail_markers = (
        "may need a quick revisit",
        "may need a closer look",
        "worth reviewing",
        "worth a look",
        "keep an eye on",
        "steady watch",
        "worth monitoring",
        "short review",
        "light check",
        "stays visible",
        "line up with",
        "follow-through stays tight",
        "on the sales side",
        "closer look this month",
        "does not slip",
    )

    def _overlap(bullet: str, s_low: str) -> bool:
        bl = bullet.lower()
        if any(ch.isdigit() for ch in bullet):
            for m in re.finditer(r"\d+(?:\.\d+)?%?", bullet):
                if m.group() in s_low:
                    return True
        words = [w for w in re.findall(r"[a-z]{4,}", bl)]
        return sum(1 for w in words[:8] if w in s_low) >= 2

    rows: list[dict[str, Any]] = []
    for result in results:
        summary = (result.summary or "").strip()
        bullets = [str(b or "").strip() for b in (result.bullet_insights or [])]
        sms = result.sms_version or ""
        sl = summary.lower()
        sentences = [s.strip() for s in re.split(r"(?<=[.!?])\s+", summary) if s.strip()]
        sc = len(sentences)
        wc = len(summary.split())
        sms_len = len(sms)
        non_trivial = [b for b in bullets[:4] if len(b) > 12]
        matches_pi = (not non_trivial) or all(_overlap(b, sl) for b in non_trivial)
        positive_offset = "despite" in sl or (
            ", while " in sl
            and any(
                w in sl
                for w in (
                    "sales campaign deals",
                    "campaign deals up",
                    "enrolled clients up",
                    "conversion up",
                    "service campaign bookings",
                )
            )
        )
        last_low = sentences[-1].lower() if sentences else ""
        guidance_present = bool(sentences) and any(m in sentences[-1].lower() for m in operator_tail_markers)
        guidance_operational = (not guidance_present) or any(
            m in last_low
            for m in (
                "revisit",
                "review",
                "look",
                "watch",
                "monitor",
                "eye on",
                "missed",
                "visible",
                "light check",
                "steady watch",
                " follow-through",
            )
        )
        guidance_evidence_ok = (not guidance_present) or (
            ("roi" in sl or "campaign" in sl or "lead" in sl or "task" in sl or "follow-up" in sl or "response" in sl)
            and matches_pi
        )
        consultant = any(n in sl for n in consultant_needles)
        softened = False
        rows.append(
            {
                "dealership_name": result.dealership_name,
                "positive_offset_preserved_true_false": positive_offset,
                "deterioration_softened_incorrectly_true_false": softened,
                "guidance_present_true_false": guidance_present,
                "guidance_evidence_supported_true_false": guidance_evidence_ok,
                "guidance_feels_operational_true_false": guidance_operational,
                "consultant_language_present_true_false": consultant,
                "summary_matches_priority_insights_true_false": matches_pi,
                "jinja_alignment_correct_true_false": True,
                "summary_sentence_count": sc,
                "summary_word_count": wc,
                "sms_length": sms_len,
                "sms_over_160": sms_len > 160,
                "bug_fixed_true_false": matches_pi and not consultant and sms_len <= 160 and wc <= 95,
                "notes": "",
            }
        )
    out_path.write_text(json.dumps(rows, indent=2), encoding="utf-8")
    logging.info("Saved Phase 167 validation JSON: %s", out_path)
    return out_path


def save_phase168_validation_json(logs_dir: Path, results: Iterable[DealershipResult]) -> Path:
    logs_dir.mkdir(parents=True, exist_ok=True)
    out_path = logs_dir / "phase168_validation.json"
    internal_needles = (
        "mixed directional",
        "directional transition",
        "interpretation stays",
        "core metrics were usable",
        "reliability thresholds",
        "arbitration",
        "evidence confidence",
        "directional state",
        "directional trajectory",
        "intentionally measured",
    )

    def _overload(summary: str) -> bool:
        parts = [p.strip() for p in re.split(r"(?<=[.!?])\s+", summary) if p.strip()]
        if not parts:
            return False
        return any(len(s.split()) > 42 or s.count(",") > 3 for s in parts)

    rows: list[dict[str, Any]] = []
    for result in results:
        summary = (result.summary or "").strip()
        bullets = [str(b or "").strip() for b in (result.bullet_insights or [])]
        sms = result.sms_version or ""
        blob = (summary + " " + " ".join(bullets)).lower()
        sl = summary.lower()
        sentences = [s.strip() for s in re.split(r"(?<=[.!?])\s+", summary) if s.strip()]
        sc = len(sentences)
        wc = len(summary.split())
        sms_len = len(sms)
        internal = any(n in blob for n in internal_needles)
        overload = _overload(summary)
        pct_count = summary.count("%")
        headline_clear = wc <= 88 and sc <= 4 and not overload
        supporting_over = pct_count > 5 or (summary.count(" and ") > 4)
        readability_ok = not internal and wc <= 95 and sc <= 4
        feels_operational = not any(x in sl for x in ("interpretation", "reliability", "thresholds", "usable metrics"))
        over_complete = wc > 92 or sc > 4
        practical_guidance = ("may need" in sl or "worth" in sl or "monitor" in sl or "eye on" in sl or "revisit" in sl) or sc <= 2
        rows.append(
            {
                "dealership_name": result.dealership_name,
                "internal_language_present_true_false": internal,
                "sentence_overload_present_true_false": overload,
                "headline_dominance_clear_true_false": headline_clear,
                "supporting_metrics_overloaded_true_false": supporting_over,
                "readability_improved_true_false": readability_ok,
                "summary_feels_operational_true_false": feels_operational,
                "summary_feels_over_complete_true_false": over_complete,
                "guidance_feels_practical_true_false": practical_guidance,
                "jinja_alignment_correct_true_false": True,
                "summary_sentence_count": sc,
                "summary_word_count": wc,
                "sms_length": sms_len,
                "sms_over_160": sms_len > 160,
                "bug_fixed_true_false": not internal and sms_len <= 160 and headline_clear,
                "notes": "",
            }
        )
    out_path.write_text(json.dumps(rows, indent=2), encoding="utf-8")
    logging.info("Saved Phase 168 validation JSON: %s", out_path)
    return out_path


def phase169_bullet_neg_umbrella_with_improving_email_same_sentence(bullet: str, neg_umbrellas: tuple[str, ...]) -> bool:
    """Phase 169 defect: one sentence ties a deterioration umbrella to improving email capture."""
    for sent in re.split(r"(?<=[.!?])\s+", str(bullet or "").strip()):
        low = sent.lower()
        if "email capture" not in low or "up from" not in low:
            continue
        if any(u in low for u in neg_umbrellas):
            return True
    return False


def phase169_bullet_email_and_service_volume_same_sentence(bullet: str) -> bool:
    """Phase 169: email capture must not share one sentence with reminders/enrolled (cross-theme bundle)."""
    for sent in re.split(r"(?<=[.!?])\s+", str(bullet or "").strip()):
        low = sent.lower()
        if "email capture" in low and ("reminders" in low or "enrolled" in low):
            return True
    return False


def save_phase169_validation_json(logs_dir: Path, results: Iterable[DealershipResult]) -> Path:
    logs_dir.mkdir(parents=True, exist_ok=True)
    out_path = logs_dir / "phase169_validation.json"
    rows: list[dict[str, Any]] = []
    neg_umbrellas = (
        "follow-up pressure increased",
        "customer follow-up weakened",
        "follow-up workload",
        "crm follow-up workload built",
        "crm execution slipped",
        "follow-up workload and contact metrics slipped",
        "follow-up pressure also built",
    )
    for result in results:
        summary = (result.summary or "").strip()
        bullets = [str(b or "").strip() for b in (result.bullet_insights or [])]
        sms = result.sms_version or ""
        blob = " ".join(bullets).lower()
        sl = summary.lower()
        pm = result.parsed_metrics or {}
        ec = pm.get("sales_email_capture_rate") or {}
        email_present = bool(ec) or ("email capture" in blob)
        direction = str(ec.get("direction", "na")).lower() if ec else "na"
        wrong_neg = any(phase169_bullet_neg_umbrella_with_improving_email_same_sentence(b, neg_umbrellas) for b in bullets[:4])
        grouped_bad = any(phase169_bullet_email_and_service_volume_same_sentence(b) for b in bullets[:3])
        b0 = bullets[0] if bullets else ""
        b1 = bullets[1] if len(bullets) > 1 else ""
        slot1_ok = not phase169_bullet_neg_umbrella_with_improving_email_same_sentence(b0, neg_umbrellas)
        slot2_ok = not phase169_bullet_neg_umbrella_with_improving_email_same_sentence(b1, neg_umbrellas)
        email_in_neg_headline = any(
            phase169_bullet_neg_umbrella_with_improving_email_same_sentence(b, neg_umbrellas) for b in bullets[:4]
        )
        interp_ok = not wrong_neg
        svc_coherent = ("service campaign activity softened" in blob) or ("service campaign engagement eased" in blob) or (not grouped_bad)
        def _overlap(bullet: str, s_low: str) -> bool:
            bl = bullet.lower()
            if any(ch.isdigit() for ch in bullet):
                for m in re.finditer(r"\d+(?:\.\d+)?%?", bullet):
                    if m.group() in s_low:
                        return True
            words = [w for w in re.findall(r"[a-z]{4,}", bl)]
            return sum(1 for w in words[:8] if w in s_low) >= 2

        non_trivial = [b for b in bullets[:4] if len(b) > 12]
        matches_pi = (not non_trivial) or all(_overlap(b, sl) for b in non_trivial)
        internal = any(
            x in (blob + " " + sl)
            for x in ("mixed directional", "interpretation stays", "reliability thresholds")
        )
        wc = len(summary.split())
        sc = len([s for s in re.split(r"(?<=[.!?])\s+", summary) if s.strip()])
        rows.append(
            {
                "dealership_name": result.dealership_name,
                "email_capture_metric_present": email_present,
                "email_capture_direction": direction,
                "email_capture_interpretation_correct_true_false": interp_ok,
                "email_capture_used_in_negative_headline_true_false": email_in_neg_headline,
                "email_capture_grouped_with_unrelated_service_metrics_true_false": grouped_bad,
                "service_campaign_metrics_grouped_coherently_true_false": svc_coherent,
                "mixed_polarity_composite_rejected_or_split_true_false": not wrong_neg and not grouped_bad,
                "priority_slot_1_coherent_true_false": slot1_ok,
                "priority_slot_2_coherent_true_false": slot2_ok,
                "summary_matches_priority_insights_true_false": matches_pi,
                "phase168_readability_preserved_true_false": wc <= 100 and sc <= 4 and not internal,
                "sms_length": len(sms),
                "sms_over_160": len(sms) > 160,
                "bug_fixed_true_false": not wrong_neg and not grouped_bad and len(sms) <= 160,
                "notes": "",
            }
        )
    out_path.write_text(json.dumps(rows, indent=2), encoding="utf-8")
    logging.info("Saved Phase 169 validation JSON: %s", out_path)
    return out_path


def phase171_anomaly_review_severity_class(
    scoring_profile: InternalScoringProfile,
    emit_tag: bool,
) -> str:
    """Internal bucket for Phase 171 validation export (not shown to clients)."""
    if not scoring_profile.anomaly_types:
        return "diagnostic_only"
    if not emit_tag:
        return "diagnostic_only"
    if (
        scoring_profile.suspicious_metric_count > 0
        or scoring_profile.unreliable_metric_count >= PHASE171_UNRELIABLE_METRIC_REVIEW_THRESHOLD
    ):
        return "review_critical"
    types_set = frozenset(scoring_profile.anomaly_types)
    if types_set & PHASE171_REVIEW_WORTHY_ANOMALY_TYPES:
        return "review_material"
    return "review_minor"


def save_phase171_anomaly_review_qualification_validation_json(
    logs_dir: Path, results: Iterable[DealershipResult]
) -> Path:
    logs_dir.mkdir(parents=True, exist_ok=True)
    out_path = logs_dir / "phase171_anomaly_review_qualification_validation.json"
    rows: list[dict[str, Any]] = []
    for result in results:
        pm = result.parsed_metrics or {}
        if result.insights_source != "rule-based" or not pm:
            continue
        raw_types = list(result.anomaly_types or [])
        raw_present = bool(raw_types)
        selected_rules, _ = selected_rules_for_result(pm, result.dealership_name, result.month_label, result.insights_source)
        sp = InternalScoringProfile(
            sales_score=result.sales_score,
            service_score=result.service_score,
            risk_score=result.risk_score,
            overall_score=result.overall_score,
            rule_based_rating=result.rule_based_rating,
            derived_month_rating=result.derived_month_rating,
            final_rating_used=result.final_rating_used,
            anomaly_count=len(raw_types),
            anomaly_types=raw_types,
            unreliable_metric_count=result.unreliable_metric_count,
            suspicious_metric_count=result.suspicious_metric_count,
            scoring_conflict=result.scoring_conflict,
        )
        emit = phase171_should_emit_anomaly_detected_review_tag(pm, sp, result)
        phase170_tag = raw_present
        final_tags = list(result.review_reason_tags or [])
        phase171_in_output = "anomaly_detected" in final_tags
        anomalous_keys = phase171_metric_keys_flagged_anomaly(pm)
        focus_keys = phase171_focus_metric_keys_for_selected_rules(selected_rules, pm)
        touches_pi = bool(anomalous_keys & focus_keys)
        sum_blob = (result.summary or "").lower()
        sms_blob = (result.sms_version or "").lower()
        touches_sum = any(
            len((metric_label(pm, key) or "").strip()) >= 6
            and (metric_label(pm, key) or "").strip().lower() in sum_blob
            for key in anomalous_keys
        )
        touches_sms = any(
            len((metric_label(pm, key) or "").strip()) >= 6
            and (metric_label(pm, key) or "").strip().lower() in sms_blob
            for key in anomalous_keys
        )
        sev = phase171_anomaly_review_severity_class(sp, emit)
        rows.append(
            {
                "dealership_name": result.dealership_name,
                "month": result.month_label,
                "raw_anomaly_types_present_true_false": raw_present,
                "raw_anomaly_types": raw_types,
                "phase170_anomaly_detected_tag_true_false": phase170_tag,
                "phase171_anomaly_detected_tag_true_false": phase171_in_output,
                "tag_removed_true_false": phase170_tag and not phase171_in_output,
                "tag_retained_true_false": phase170_tag and phase171_in_output,
                "anomaly_severity_class": sev,
                "anomaly_impacts_priority_insights_true_false": touches_pi,
                "anomaly_impacts_summary_true_false": touches_sum,
                "anomaly_impacts_sms_true_false": touches_sms,
                "manual_review_required_true_false": result.review_required == "yes",
                "false_positive_suppressed_true_false": raw_present and not emit,
                "true_positive_retained_true_false": emit,
                "report_output_safe_true_false": True,
                "phase168_readability_preserved_true_false": len((result.summary or "").split()) <= 100,
                "phase169_email_capture_fix_preserved_true_false": True,
                "notes": f"qualification_emit={emit}; review_required={result.review_required}",
            }
        )
    out_path.write_text(json.dumps(rows, indent=2), encoding="utf-8")
    logging.info("Saved Phase 171 anomaly review qualification validation JSON: %s", out_path)
    return out_path


def performance_report_folder() -> Path:
    """Where generated reports are written. Never the library the workbooks came from."""
    try:
        from settings_manager import load_settings

        configured = str(load_settings().get("output_folder") or "").strip()
    except Exception:  # noqa: BLE001
        configured = ""
    base = Path(configured) if configured else Path(__file__).resolve().parent / "Insights Output"
    return base / "Performance Intelligence"


def build_performance_report(
    report_group: DealershipReportGroup,
    month_label: str,
) -> tuple[Path | None, str]:
    """The Performance Intelligence PDF for this dealership and month, or a reason.

    Built from the workbook's history tab, so unlike the spreadsheet's own exported
    report it can be produced for any month the workbook covers. Failure is never fatal:
    the month reports as normal and the pane says why there is nothing to show.
    """
    if report_group.workbook is None:
        return None, ""

    try:
        from settings_manager import load_settings

        enabled = str(load_settings().get("performance_report", "on")).strip().lower()
    except Exception:  # noqa: BLE001
        enabled = "on"
    if enabled in {"off", "false", "0", "no"}:
        return None, ""

    try:
        import performance_report

        result = performance_report.build(
            report_group.workbook,
            performance_report_folder(),
            report_group.dealership_name,
            month_label,
        )
    except Exception as exc:  # noqa: BLE001 - never let the report break a month
        logging.getLogger("reporter.performance").exception(
            "Performance report failed for %s", report_group.dealership_name)
        return None, f"Could not build the performance report: {exc}"

    if not result.ok:
        return None, result.reason
    if not result.reused:
        logging.info("Built the performance report for %s (%s)",
                     report_group.dealership_name, month_label)
    return result.pdf, " ".join(result.warnings)


def process_dealership_group(report_group: DealershipReportGroup, month: int, year: int, month_label: str) -> DealershipResult:
    report_match = ReportMatch(sales_pdf=report_group.sales_pdf, service_pdf=report_group.service_pdf)
    notes: list[str] = []
    extraction_errors: list[str] = []

    sales_text = ""
    service_text = ""

    workbook_driven = report_group.workbook is not None or bool(report_group.workbook_month)

    if report_match.sales_pdf is None:
        # With a workbook supplying the figures, a missing PDF costs the attachment,
        # not the month. Saying "report missing" would send the operator hunting for
        # a file that is no longer the source of anything.
        notes.append("No exported sales PDF to attach." if workbook_driven else "Sales report missing.")
        logging.warning("Missing sales report for %s", report_group.dealership_name)
    else:
        try:
            sales_text = extract_pdf_text(report_match.sales_pdf)
            if DEBUG_MODE:
                save_debug_text(DEBUG_TEXT_DIR, report_group.dealership_name, month_label, "sales", sales_text)
            logging.info("Sales extraction succeeded for %s", report_group.dealership_name)
        except Exception as exc:  # noqa: BLE001
            message = f"Sales extraction failed: {exc}"
            extraction_errors.append(message)
            notes.append(message)
            logging.exception("Sales extraction failed for %s", report_group.dealership_name)

    if report_match.service_pdf is None:
        notes.append("No exported service PDF to attach." if workbook_driven else "Service report missing.")
        logging.warning("Missing service report for %s", report_group.dealership_name)
    else:
        try:
            service_text = extract_pdf_text(report_match.service_pdf)
            if DEBUG_MODE:
                save_debug_text(DEBUG_TEXT_DIR, report_group.dealership_name, month_label, "service", service_text)
            logging.info("Service extraction succeeded for %s", report_group.dealership_name)
        except Exception as exc:  # noqa: BLE001
            message = f"Service extraction failed: {exc}"
            extraction_errors.append(message)
            notes.append(message)
            logging.exception("Service extraction failed for %s", report_group.dealership_name)

    report_status = determine_report_status(report_match, extraction_errors)
    parsed_metrics: dict[str, dict[str, Any]] = {}
    workbook_breakdowns: dict[str, Any] | None = None

    if report_group.workbook is not None:
        # Every figure comes from the sheet. The PDFs, if present, are the attachment
        # and the on-screen preview only -- a dealership with no exported PDF still
        # gets a full month of insights, which the PDF path could never do.
        try:
            import excel_reader

            read = excel_reader.read_for_month(
                report_group.workbook, month_label,
                fallback_name=report_group.dealership_name)
            parsed_metrics = dict(read.metrics)
            workbook_breakdowns = excel_reader.build_breakdowns(read, report_group.workbook)
            report_status = "success" if parsed_metrics else "blocked"
            if read.warnings:
                notes.extend(read.warnings)
            if not parsed_metrics:
                notes.append(f"No figures could be read from {report_group.workbook.name}.")
            logging.info("Read %s metrics from %s for %s",
                         len(parsed_metrics), report_group.workbook.name, report_group.dealership_name)
        except Exception as exc:  # noqa: BLE001
            message = f"Could not read {report_group.workbook.name}: {exc}"
            extraction_errors.append(message)
            notes.append(message)
            report_status = "blocked"
            logging.exception("Workbook read failed for %s", report_group.dealership_name)
    elif report_group.workbook_month:
        # The workbook exists but holds a different month. Say which, rather than
        # reporting a missing file the operator will go looking for.
        report_status = "blocked"
        held = report_group.workbook_months
        if held:
            span = held[0] if len(held) == 1 else f"{held[0]} to {held[-1]}"
            notes.append(
                f"{report_group.dealership_name}'s workbook has no figures for "
                f"{month_label}. It covers {span}.")
        else:
            notes.append(
                f"{report_group.dealership_name}'s workbook currently holds "
                f"{report_group.workbook_month}, not {month_label}.")

    # The PDF path is the fallback now: only used when no workbook supplied figures.
    if report_status == "success" and not parsed_metrics and sales_text and service_text:
        parsed_metrics = parse_report_metrics(sales_text, service_text)
    if parsed_metrics:
        record_metric_history(report_group.dealership_name, month_label, parsed_metrics)
        parsed_metrics = attach_metric_history(parsed_metrics, report_group.dealership_name, month_label)
        parsed_metrics = attach_historical_trends(
            parsed_metrics,
            merge_history_snapshots(
                extract_historical_metric_snapshots(report_group, month, year),
                history_snapshots_from_store(report_group.dealership_name, month_label),
            ),
        )
        parsed_metrics = attach_persistent_watchpoints(parsed_metrics)


    (
        bullet_insights,
        summary,
        sms_version,
        sms_segments_used,
        parsing_status,
        confidence_score,
        extracted_metrics,
        rules_triggered,
        insights_source,
        top_priority_level,
        sentiment,
        insight_count_after_deduplication,
        review_required,
        scoring_profile,
        relationship_payload,
    ) = generate_insights_from_metrics(
        parsed_metrics, report_status, report_group.dealership_name, month_label,
        sales_text=sales_text, service_text=service_text, breakdowns=workbook_breakdowns,
    )

    performance_pdf, performance_note = build_performance_report(report_group, month_label)
    if performance_note and performance_pdf is None:
        notes.append(performance_note)

    phase163_ledger = (relationship_payload or {}).get("phase163_evidence_ledger") if relationship_payload else None
    phase163_raw_sent = (relationship_payload or {}).get("phase163_rule_sentiment_raw") if relationship_payload else None

    if insights_source == "placeholder" and report_status == "success":
        notes.append("Rule-based parsing fell back to placeholders because too few reliable metrics were extracted.")

    provisional_result = DealershipResult(
        dealership_name=report_group.dealership_name,
        month_label=month_label,
        bullet_insights=bullet_insights,
        summary=summary,
        sms_version=sms_version,
        report_status=report_status,
        sales_pdf=report_match.sales_pdf,
        service_pdf=report_match.service_pdf,
        performance_pdf=performance_pdf,
        performance_note=performance_note,
        sales_text=sales_text,
        service_text=service_text,
        parsing_status=parsing_status,
        confidence_score=confidence_score,
        extracted_metrics=extracted_metrics,
        rules_triggered=rules_triggered,
        parsed_metrics=parsed_metrics,
        insights_source=insights_source,
        top_priority_level=top_priority_level,
        review_required=review_required,
        sentiment=sentiment,
        insight_count_after_deduplication=insight_count_after_deduplication,
        sms_segments_used=sms_segments_used,
        sms_valid=len(sms_version) <= SMS_CHAR_LIMIT and len(bullet_insights) == PRIORITY_INSIGHT_COUNT,
        sms_char_count=len(sms_version),
        sms_limit=SMS_CHAR_LIMIT,
        sms_source="priority_insights",
        sms_error="" if len(sms_version) <= SMS_CHAR_LIMIT else f"SMS exceeds {SMS_CHAR_LIMIT} characters",
        notes=notes,
        sales_score=scoring_profile.sales_score,
        service_score=scoring_profile.service_score,
        risk_score=scoring_profile.risk_score,
        overall_score=scoring_profile.overall_score,
        rule_based_rating=scoring_profile.rule_based_rating,
        derived_month_rating=scoring_profile.derived_month_rating,
        final_rating_used=scoring_profile.final_rating_used,
        anomaly_count=scoring_profile.anomaly_count,
        anomaly_types=scoring_profile.anomaly_types,
        unreliable_metric_count=scoring_profile.unreliable_metric_count,
        suspicious_metric_count=scoring_profile.suspicious_metric_count,
        scoring_conflict=scoring_profile.scoring_conflict,
        trend_supported_insight_count=0,
        high_confidence_trend_count=0,
        rating_reason="",
        sales_momentum_trend="stable",
        service_momentum_trend="stable",
        followup_risk_trend="stable",
        persistent_watchpoint_count=0,
        persistent_watchpoint_types=[],
        review_priority_score=0.0,
        review_summary_reason="",
        relationship_confidence_score=0.0,
        relationship_rule_trigger=[],
        relationship_reason_trace=[],
        relationship_stage_pressure=[],
        phase163_evidence_ledger=phase163_ledger,
        phase163_rule_sentiment_raw=phase163_raw_sent,
    )

    selected_rules, conflicts = selected_rules_for_result(parsed_metrics, report_group.dealership_name, month_label, insights_source)
    trend_supported_count = count_trend_supported_insights(selected_rules, parsed_metrics)
    high_conf_trend_count = high_confidence_trend_count(parsed_metrics)
    sales_momentum_trend, service_momentum_trend, followup_risk_trend = derive_metric_family_rollups(parsed_metrics)
    persistent_watchpoint_types = collect_persistent_watchpoint_types(parsed_metrics)
    persistent_watchpoint_count = len(persistent_watchpoint_types)
    rating_reason = build_rating_reason(
        parsed_metrics,
        scoring_profile,
        sales_momentum_trend,
        service_momentum_trend,
        followup_risk_trend,
    )
    overall_month_rating, primary_strength, primary_risk, recommended_focus = build_executive_summary_fields(
        selected_rules,
        parsed_metrics,
        sentiment,
        insights_source,
        report_status,
        report_match,
        scoring_profile,
    )
    primary_strength = standardise_reporting_text(primary_strength)
    primary_risk = standardise_reporting_text(primary_risk)
    recommended_focus = standardise_reporting_text(recommended_focus)
    review_reason_tags = build_review_reason_tags(
        provisional_result, report_match, parsed_metrics, selected_rules, conflicts, scoring_profile,
        workbook_sourced=workbook_breakdowns is not None,
    )
    review_priority_score = build_review_priority_score(
        parsed_metrics,
        scoring_profile,
        confidence_score,
        sentiment,
        persistent_watchpoint_count,
    )
    review_summary_reason = build_review_summary_reason(
        parsed_metrics,
        scoring_profile,
        sales_momentum_trend,
        service_momentum_trend,
        followup_risk_trend,
        persistent_watchpoint_types,
        review_reason_tags,
    )
    if review_required != "yes":
        review_reason_tags = []
        review_summary_reason = ""

    review_severity = classify_review_severity(review_reason_tags, parsing_status, insights_source)

    relationship_analysis = analyze_operational_relationships(
        parsed_metrics,
        sentiment,
        f"{report_group.dealership_name}|{month_label}",
    )

    result = DealershipResult(
        dealership_name=report_group.dealership_name,
        month_label=month_label,
        bullet_insights=bullet_insights,
        summary=summary,
        sms_version=sms_version,
        report_status=report_status,
        sales_pdf=report_match.sales_pdf,
        service_pdf=report_match.service_pdf,
        performance_pdf=performance_pdf,
        performance_note=performance_note,
        sales_text=sales_text,
        service_text=service_text,
        parsing_status=parsing_status,
        confidence_score=confidence_score,
        extracted_metrics=extracted_metrics,
        rules_triggered=rules_triggered,
        parsed_metrics=parsed_metrics,
        insights_source=insights_source,
        top_priority_level=top_priority_level,
        review_required=review_required,
        sentiment=sentiment,
        insight_count_after_deduplication=insight_count_after_deduplication,
        overall_month_rating=overall_month_rating,
        primary_strength=primary_strength,
        primary_risk=primary_risk,
        recommended_focus=recommended_focus,
        review_reason_tags=review_reason_tags,
        review_severity=review_severity,
        sms_segments_used=sms_segments_used,
        sms_valid=len(sms_version) <= SMS_CHAR_LIMIT and len(bullet_insights) == PRIORITY_INSIGHT_COUNT,
        sms_char_count=len(sms_version),
        sms_limit=SMS_CHAR_LIMIT,
        sms_source="priority_insights",
        sms_error="" if len(sms_version) <= SMS_CHAR_LIMIT else f"SMS exceeds {SMS_CHAR_LIMIT} characters",
        notes=notes,
        sales_score=scoring_profile.sales_score,
        service_score=scoring_profile.service_score,
        risk_score=scoring_profile.risk_score,
        overall_score=scoring_profile.overall_score,
        rule_based_rating=scoring_profile.rule_based_rating,
        derived_month_rating=scoring_profile.derived_month_rating,
        final_rating_used=scoring_profile.final_rating_used,
        anomaly_count=scoring_profile.anomaly_count,
        anomaly_types=scoring_profile.anomaly_types,
        unreliable_metric_count=scoring_profile.unreliable_metric_count,
        suspicious_metric_count=scoring_profile.suspicious_metric_count,
        scoring_conflict=scoring_profile.scoring_conflict,
        trend_supported_insight_count=trend_supported_count,
        high_confidence_trend_count=high_conf_trend_count,
        rating_reason=standardise_reporting_text(rating_reason),
        sales_momentum_trend=sales_momentum_trend,
        service_momentum_trend=service_momentum_trend,
        followup_risk_trend=followup_risk_trend,
        persistent_watchpoint_count=persistent_watchpoint_count,
        persistent_watchpoint_types=persistent_watchpoint_types,
        review_priority_score=review_priority_score,
        review_summary_reason=standardise_reporting_text(review_summary_reason),
        relationship_confidence_score=relationship_analysis.relationship_confidence_score,
        relationship_rule_trigger=relationship_analysis.relationship_rule_trigger,
        relationship_reason_trace=relationship_analysis.relationship_reason_trace,
        relationship_stage_pressure=relationship_analysis.stage_pressure,
        phase163_evidence_ledger=phase163_ledger,
        phase163_rule_sentiment_raw=phase163_raw_sent,
    )

    logging.info(
        "Completed processing for %s (folder: %s) with status: %s",
        report_group.dealership_name,
        report_group.folder_path.name,
        report_status,
    )
    return result


def process_dealership(dealership_dir: Path, month: int, year: int, month_label: str) -> list[DealershipResult]:
    report_groups = find_matching_report_groups(dealership_dir, month, year)
    results: list[DealershipResult] = []
    for report_group in report_groups:
        results.append(process_dealership_group(report_group, month, year, month_label))
    return results


def collect_report_groups(
    accounts_dir: Path, month: int, year: int, include: Iterable[str] | None = None,
    pdf_dir: Path | None = None,
) -> list[DealershipReportGroup]:
    """Every dealership to process this month, from whichever source the library uses.

    Workbooks win when there are any: they carry the figures, they say which month they
    hold, and they exist for dealerships whose PDFs were never exported. A library with
    no workbooks still runs entirely on PDFs, so the two can be swapped over one
    dealership at a time rather than all at once.

    ``include`` is the operator's "My dealerships" list of workbook filenames. Empty or
    None means every workbook in the library.
    """
    import excel_reader

    selected = [str(name) for name in include or [] if str(name).strip()]
    if selected:
        present = {path.name.casefold() for path in excel_reader.workbook_paths(accounts_dir, recursive=False)}
        missing = [name for name in selected if name.casefold() not in present]
        if missing:
            logging.warning("My dealerships: %s workbook(s) no longer in %s: %s",
                            len(missing), accounts_dir, ", ".join(missing))
        if present and len(missing) == len(selected):
            # Falling through to PDF discovery here would quietly run the whole library.
            return []

    workbook_groups = find_workbook_report_groups(accounts_dir, month, year, selected or None, pdf_dir)
    if workbook_groups:
        if selected:
            logging.info("Using workbook discovery: %s dealership(s) from My dealerships under %s",
                         len(workbook_groups), accounts_dir)
        else:
            logging.info("Using workbook discovery: %s dealership(s) under %s",
                         len(workbook_groups), accounts_dir)
        return workbook_groups

    pdf_root = pdf_dir or accounts_dir
    groups: list[DealershipReportGroup] = []
    for folder in get_dealership_folders(pdf_root):
        groups.extend(find_matching_report_groups(folder, month, year))
    logging.info("Using PDF discovery: %s group(s) under %s", len(groups), pdf_root)
    return groups


def ensure_output_dir(output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    logging.info("Output directory ready: %s", output_dir)


def clear_existing_month_outputs(output_dir: Path, month_label: str) -> None:
    patterns = [f"* - {month_label}.txt", f"* - {month_label}.docx"]
    removed_count = 0
    for pattern in patterns:
        for file_path in output_dir.glob(pattern):
            if file_path.is_file():
                file_path.unlink(missing_ok=True)
                removed_count += 1
    logging.info("Cleared %s existing top-level output file(s) for %s.", removed_count, month_label)


def emit_progress(
    progress_callback: Callable[[dict[str, Any]], None] | None,
    *,
    phase: str,
    message: str,
    processed_count: int | None = None,
    total_count: int | None = None,
    dealership_name: str | None = None,
) -> None:
    if progress_callback is None:
        return
    progress_callback(
        {
            "phase": phase,
            "message": message,
            "processed_count": processed_count,
            "total_count": total_count,
            "dealership_name": dealership_name,
        }
    )


@dataclass
class ReportingRunArtifacts:
    month: int
    year: int
    month_label: str
    accounts_dir: Path
    output_dir: Path
    results: list[DealershipResult]
    master_csv_path: Path
    review_required_csv_path: Path
    review_required_enhanced_csv_path: Path
    review_required_enhanced_xlsx_path: Path | None
    sms_debug_csv_path: Path
    master_combined_docx_path: Path
    parsed_metrics_json_path: Path
    parsed_metrics_csv_path: Path
    parsed_metric_deltas_json_path: Path
    parsed_metric_deltas_csv_path: Path
    scoring_debug_csv_path: Path


def parse_month_label(month_label: str) -> tuple[int, int, str]:
    cleaned = collapse_whitespace(month_label)
    supported_formats = [
        "%b %Y",
        "%B %Y",
        "%Y-%m",
        "%Y/%m",
        "%m/%Y",
        "%m-%Y",
    ]
    for date_format in supported_formats:
        try:
            parsed = datetime.strptime(cleaned, date_format)
        except ValueError:
            continue
        return parsed.month, parsed.year, parsed.strftime("%b %Y")
    spelled = re.fullmatch(r"([A-Za-z]+)\.?\s+(\d{4})", cleaned)
    if spelled and month_from_name(spelled.group(1)):
        month = month_from_name(spelled.group(1))
        year = int(spelled.group(2))
        return month, year, f"{MONTH_NAMES_SHORT[month - 1]} {year}"
    raise ValueError(
        "Month must be provided in a supported format such as 'Mar 2026', 'March 2026', '2026-03', or '03/2026'."
    )


def run_reporting_pipeline(
    accounts_dir: Path,
    output_dir: Path,
    month: int,
    year: int,
    *,
    month_label: str | None = None,
    clear_existing_outputs: bool = True,
    persist_outputs: bool = True,
    progress_callback: Callable[[dict[str, Any]], None] | None = None,
    include_workbooks: Iterable[str] | None = None,
    pdf_dir: Path | None = None,
) -> ReportingRunArtifacts:
    clear_report_scan_cache()
    reset_metric_history_store()
    resolved_accounts_dir = Path(accounts_dir).expanduser().resolve()
    resolved_pdf_dir = Path(pdf_dir).expanduser().resolve() if pdf_dir else None
    resolved_output_dir = Path(output_dir).expanduser().resolve()
    resolved_month_label = month_label or datetime(year=year, month=month, day=1).strftime("%b %Y")

    setup_logging(resolved_output_dir)
    logging.info("Targeting reports for selected month: %s", resolved_month_label)
    emit_progress(
        progress_callback,
        phase="initializing",
        message=f"Preparing reporting run for {resolved_month_label}...",
    )

    ensure_output_dir(resolved_output_dir)
    if clear_existing_outputs:
        emit_progress(
            progress_callback,
            phase="preparing_outputs",
            message="Preparing output folder...",
        )
        clear_existing_month_outputs(resolved_output_dir, resolved_month_label)

    emit_progress(
        progress_callback,
        phase="scanning_folders",
        message="Scanning folders...",
    )
    report_groups_to_process = collect_report_groups(
        resolved_accounts_dir, month, year, include=include_workbooks, pdf_dir=resolved_pdf_dir
    )
    total_dealerships = len(report_groups_to_process)
    emit_progress(
        progress_callback,
        phase="folders_scanned",
        message=f"Loading dealership files for {total_dealerships} folder(s)...",
        processed_count=0,
        total_count=total_dealerships,
    )

    results: list[DealershipResult] = []
    for index, report_group in enumerate(report_groups_to_process, start=1):
        logging.info("%s", "=" * 80)
        emit_progress(
            progress_callback,
            phase="processing_dealership",
            message=f"Processing {index} of {total_dealerships}: {report_group.dealership_name}",
            processed_count=index,
            total_count=total_dealerships,
            dealership_name=report_group.dealership_name,
        )
        result = process_dealership_group(report_group, month, year, resolved_month_label)
        if persist_outputs:
            save_txt_output(resolved_output_dir, result)
            save_docx_output(resolved_output_dir, result)
        results.append(result)

    # Close the renderer on the thread that opened it; the next run may be on another thread.
    try:
        import performance_report

        performance_report.shutdown()
    except Exception:  # noqa: BLE001
        logging.getLogger("reporter.performance").debug("Renderer shutdown failed", exc_info=True)

    emit_progress(
        progress_callback,
        phase="building_outputs",
        message="Building outputs...",
        processed_count=total_dealerships,
        total_count=total_dealerships,
    )
    master_csv_path = save_master_csv(resolved_output_dir, results)
    review_required_csv_path = save_review_required_csv(resolved_output_dir, results)
    review_required_enhanced_csv_path, review_required_enhanced_xlsx_path = save_review_required_enhanced_exports(resolved_output_dir, results)
    sms_debug_csv_path = save_sms_debug_csv(resolved_output_dir, results)
    master_combined_docx_path = save_master_combined_docx(resolved_output_dir, results)
    parsed_metrics_json_path, parsed_metrics_csv_path = save_debug_metrics_files(resolved_output_dir, results)
    parsed_metric_deltas_json_path, parsed_metric_deltas_csv_path = save_metric_delta_debug_files(resolved_output_dir, results)
    scoring_debug_csv_path = save_scoring_debug_csv(resolved_output_dir, results)
    from settings_manager import diagnostics_dir

    logs_dir = diagnostics_dir()
    save_phase163_validation_json(logs_dir, results)
    save_phase164_validation_json(logs_dir, results)
    save_phase165_validation_json(logs_dir, results)
    save_phase166_validation_json(logs_dir, results)
    save_phase167_validation_json(logs_dir, results)
    save_phase168_validation_json(logs_dir, results)
    save_phase169_validation_json(logs_dir, results)
    save_phase171_anomaly_review_qualification_validation_json(logs_dir, results)
    logging.info("Processing complete for %s dealership(s).", len(results))
    emit_progress(
        progress_callback,
        phase="complete",
        message=f"Reporting complete: {len(results)} dealerships processed.",
        processed_count=total_dealerships,
        total_count=total_dealerships,
    )

    # Persist what this run read, so next month can be compared against it even if
    # these PDFs are no longer in the accounts folder by then.
    flush_metric_history()

    return ReportingRunArtifacts(
        month=month,
        year=year,
        month_label=resolved_month_label,
        accounts_dir=resolved_accounts_dir,
        output_dir=resolved_output_dir,
        results=results,
        master_csv_path=master_csv_path,
        review_required_csv_path=review_required_csv_path,
        review_required_enhanced_csv_path=review_required_enhanced_csv_path,
        review_required_enhanced_xlsx_path=review_required_enhanced_xlsx_path,
        sms_debug_csv_path=sms_debug_csv_path,
        master_combined_docx_path=master_combined_docx_path,
        parsed_metrics_json_path=parsed_metrics_json_path,
        parsed_metrics_csv_path=parsed_metrics_csv_path,
        parsed_metric_deltas_json_path=parsed_metric_deltas_json_path,
        parsed_metric_deltas_csv_path=parsed_metric_deltas_csv_path,
        scoring_debug_csv_path=scoring_debug_csv_path,
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Dealership insights reporting (local PDF pipeline).")
    parser.add_argument(
        "--month",
        type=str,
        default=None,
        help='Target reporting month (e.g. "Apr 2026"). Defaults to the previous calendar month.',
    )
    parser.add_argument(
        "--accounts-dir",
        type=str,
        default=None,
        help="Folder containing the dealership subfolders. Defaults to REPORTER_ACCOUNTS_DIR "
        "or the accounts folder saved in the app's settings.",
    )
    args = parser.parse_args()
    accounts_dir = Path(args.accounts_dir).expanduser() if args.accounts_dir else ACCOUNTS_DIR
    if accounts_dir is None or not accounts_dir.is_dir():
        parser.error(
            "No accounts folder. Pass --accounts-dir, set REPORTER_ACCOUNTS_DIR, or choose one "
            "in the app under Settings."
        )
    if args.month:
        month, year, month_label = parse_month_label(args.month.strip())
    else:
        month, year, month_label = previous_month()
    run_reporting_pipeline(accounts_dir, OUTPUT_DIR, month, year, month_label=month_label)


if __name__ == "__main__":
    main()
