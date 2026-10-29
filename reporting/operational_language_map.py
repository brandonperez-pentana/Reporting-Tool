"""
Phase 156 — Deterministic operational language naturalisation.

Replaces consultant-style / abstract phrasing with dealership-realistic wording.
Same inputs always yield the same outputs (ordered literal replacement, longest first).
"""

from __future__ import annotations

import re
_PHASE156_RAW_REPLACEMENTS: tuple[tuple[str, str], ...] = (
    (
        "Enquiry volume eased, yet stronger conversion shifted interpretive weight toward mid-funnel execution rather than raw demand volume alone.",
        "Lead volume eased, yet stronger conversion mattered more than raw lead counts alone.",
    ),
    (
        "Lead flow eased while conversion firmed, so mid-funnel execution carried more interpretive weight than headline enquiry totals alone.",
        "Lead flow eased while conversion firmed, so middle-of-funnel work mattered more than headline lead totals alone.",
    ),
    (
        "Customer-management deterioration intensified, including ",
        "Customer follow-up weakened further, including ",
    ),
    (
        "CRM execution pressure increased materially, with ",
        "Follow-up pressure increased this month, with ",
    ),
    (
        "CRM execution quality deteriorated across multiple indicators, with ",
        "Several follow-up signals worsened together, with ",
    ),
    (
        "Although enquiry volume softened, stronger conversion performance provided a partial operational offset.",
        "Leads softened, but stronger conversions partly offset that.",
    ),
    (
        "Although enquiry volume softened, stronger conversion efficiency provided a partial operational offset.",
        "Leads softened, but stronger conversions partly offset that.",
    ),
    (
        "The defining business condition was stronger closing discipline outweighing softer demand reads.",
        "Conversions firmed even as leads softened—that framed the month.",
    ),
    (
        "weaker conversion quality relative to available enquiry",
        "weaker conversion for the lead volume you had",
    ),
    (
        "This month centred on conversion improvement as the interpretive anchor.",
        "This month hinged on conversion improving.",
    ),
    (
        "What fundamentally defined the month was softer demand at the enquiry and deal-throughput level.",
        "The main story was softer leads and slower deal throughput.",
    ),
    (
        "The defining business condition was demand softening rather than a single conversion spike.",
        "Leads slowed rather than a one-off conversion blip.",
    ),
    (
        "What fundamentally defined the month was sustained operational pressure in follow-up and responsiveness.",
        "The month stayed heavy on follow-up load and response times.",
    ),
    (
        "This month centred on sustained operational risk in contact speed and pipeline hygiene.",
        "Contact speed and pipeline hygiene stayed the risk focus.",
    ),
    (
        "What fundamentally defined the month was broader funnel deterioration across progression and conversion.",
        "The sales path showed leakage across stages, not just one weak metric.",
    ),
    (
        "The defining business condition was leakage through multiple funnel stages, not an isolated metric.",
        "Several funnel stages leaked at once rather than one isolated weak spot.",
    ),
    (
        "What fundamentally defined the month was deteriorating conversion discipline through the sales funnel.",
        "Conversion discipline slipped through the sales funnel.",
    ),
    (
        "What fundamentally defined the month was mixed, volatile operating conditions without a single clean trend.",
        "The month was mixed and jumpy—no single clean trend.",
    ),
    (
        "The defining business condition was cross-currents across sales and service with uneven month-to-month movement.",
        "Sales and service sent mixed signals week to week.",
    ),
    (
        "What fundamentally defined the month was a cleaner operational trajectory in execution and follow-up.",
        "Day-to-day follow-up and execution looked cleaner than prior months.",
    ),
    (
        "What fundamentally defined the month was recovering sales efficiency against a tighter or smaller enquiry pool.",
        "Sales efficiency improved against a smaller or tighter lead pool.",
    ),
    (
        "The defining business condition was better conversion discipline even as headline demand eased.",
        "Conversion held up better even as headline leads eased.",
    ),
    (
        "What fundamentally defined the month was conversion-led recovery even as headline enquiry eased.",
        "Conversion carried the month even as headline enquiry eased.",
    ),
    (
        "What fundamentally defined the month was lead-quality strain—stronger enquiry flow paired with weaker conversion economics.",
        "More leads came in but conversion economics looked weaker—that mismatch framed the month.",
    ),
    (
        "Lead handling and conversion quality diverged; executives should weight the stronger of the two evidence stacks when prioritising action.",
        "Lead handling and conversion pointed different ways—follow whichever lane had the clearer signal.",
    ),
    (
        "Cross-currents were material, yet ",
        "Signals conflicted, but ",
    ),
    (
        " carried the greater share of interpretive weight.",
        " mattered more in practice.",
    ),
    (
        "Additional CRM strain showed through ",
        "Also this month: ",
    ),
    (
        "Follow-up execution pressure intensified, with ",
        "Follow-up pressure increased, with ",
    ),
    (
        "The executive read hinges on conversion behaviour and mid-funnel execution rather than volume alone.",
        "Read the month on conversion and middle-of-funnel work, not lead count alone.",
    ),
    (
        "headline demand softened while mid-funnel execution defined the month",
        "headline leads softened while conversion work defined the month",
    ),
    (
        "read the month through conversion quality, not only lead counts",
        "read the month on conversion, not only lead counts",
    ),
    (
        "The main watchpoint remained conversion quality, with ",
        "The main watchpoint remained conversion, with ",
    ),
    (
        "partial operational offset",
        "helped offset softer volumes",
    ),
    (
        "interpretive weight toward mid-funnel execution",
        "weight toward conversion and progression work",
    ),
    (
        "mid-funnel execution",
        "conversion and progression work",
    ),
    (
        "Mid-funnel execution",
        "Conversion and progression work",
    ),
    (
        "Conversion and progression work improved",
        "Conversions improved",
    ),
    (
        "Conversion and progression work",
        "Conversion and test-drive work",
    ),
    (
        "Strategic read: mid-funnel strength.",
        "Read: middle of the funnel held.",
    ),
    (
        "Strategic read: closing quality leads.",
        "Read: closing quality carried the month.",
    ),
    (
        "Strategic read: contact discipline first.",
        "Read: nail contact discipline first.",
    ),
    (
        "Dominant condition: demand softening.",
        "Dominant condition: slower leads.",
    ),
    (
        "Dominant condition: sustained execution pressure.",
        "Dominant condition: follow-up still under strain.",
    ),
    (
        "Dominant condition: conversion-led efficiency recovery.",
        "Dominant condition: tighter funnel, softer leads.",
    ),
    (
        "Strategic read: ",
        "Read: ",
    ),
    (
        "What fundamentally defined the month was ",
        "The month came down to ",
    ),
    (
        "interpretive anchor",
        "main thread",
    ),
    (
        "interpretive weight",
        "weight in the story",
    ),
    (
        "CRM execution pressure",
        "Follow-up pressure",
    ),
    (
        "What fundamentally defined the month was profitability and campaign-return pressure.",
        "Campaign economics and ROI were the main pressure this month.",
    ),
    (
        "The defining business condition was commercial return quality, not peripheral activity noise.",
        "Commercial return quality—not side metrics—was the clearer story.",
    ),
    (
        "This month centred on ROI and economics as the headline operational stress.",
        "ROI and economics were the headline pressure this month.",
    ),
    (
        "Sales and service campaign returns both weakened—ROI is the headline pressure.",
        "Sales and workshop campaign ROI both weakened—tighten yield before scaling activity.",
    ),
    (
        "Campaign economics softened on sales and workshop sides this month.",
        "Sales and workshop campaign economics softened this month.",
    ),
)

_PHASE162_RAW_REPLACEMENTS: tuple[tuple[str, str], ...] = (
    (
        "Campaign economics and ROI were the main pressure this month.",
        "Campaign ROI was the main pressure.",
    ),
    (
        "This month hinged on conversion improving.",
        "Conversion picked up.",
    ),
    (
        "Stronger closing partly cushioned the softer lead flow.",
        "Conversion helped offset softer lead flow.",
    ),
    (
        "Lead volume eased, yet stronger conversion mattered more than raw lead counts alone.",
        "Leads eased, but conversion mattered more than lead count.",
    ),
    (
        "Commercial return quality is the headline risk to",
        "Campaign returns are the headline risk to",
    ),
    (
        "Narrative clarity was limited by cross-currents across sales and service.",
        "Sales and workshop reads did not line up cleanly.",
    ),
    (
        "Some areas improved, but the heavier downside in core commercial drivers still read clearer.",
        "Some areas improved, but the downside in core drivers still read clearer.",
    ),
    (
        "Constructive pockets existed, yet the stressed drivers still carried most of the month's weight.",
        "Upside pockets existed, yet pressure drivers still led the month.",
    ),
    (
        "Monitor response-time and uncontacted-lead trajectories alongside conversion.",
        "Keep an eye on response times, untouched leads, and conversion.",
    ),
    (
        "Track contact coverage as a leading indicator before inferring demand weakness.",
        "Watch contact coverage before reading demand weakness.",
    ),
    (
        "operational pressure",
        "workload pressure",
    ),
    (
        "mixed operating conditions",
        "mixed month",
    ),
    (
        "sustained directional improvement across recent months",
        "a sustained lift over recent months",
    ),
    (
        "carrying sustained directional improvement across recent months",
        "holding a sustained lift over recent months",
    ),
    (
        "under sustained pressure across recent months",
        "under pressure across recent months",
    ),
    (
        "This month saw ",
        "",
    ),
    (
        "performance remained under pressure",
        "results stayed under pressure",
    ),
    (
        "momentum softened",
        "momentum eased",
    ),
    (
        "activity remained elevated",
        "activity stayed high",
    ),
    (
        "broader conditions",
        "the wider picture",
    ),
    (
        "mixed performance environment",
        "mixed month",
    ),
    (
        "demand environment",
        "lead environment",
    ),
    (
        "underlying pressure",
        "steady pressure",
    ),
    (
        "commercial execution",
        "day-to-day execution",
    ),
    (
        "retention execution",
        "retention follow-through",
    ),
    (
        "execution softened",
        "execution eased",
    ),
)

PHASE162_ORDERED_REPLACEMENTS: tuple[tuple[str, str], ...] = tuple(
    sorted(_PHASE162_RAW_REPLACEMENTS, key=lambda pair: -len(pair[0]))
)

PHASE156_ORDERED_REPLACEMENTS: tuple[tuple[str, str], ...] = tuple(
    sorted(_PHASE156_RAW_REPLACEMENTS, key=lambda pair: -len(pair[0]))
)


def phase156_naturalise_text(text: str) -> str:
    if not text:
        return text
    out = str(text)
    for old, new in PHASE156_ORDERED_REPLACEMENTS:
        if old in out:
            out = out.replace(old, new)
    return out


def phase162_humanise_text(text: str) -> str:
    """Phase 162 — concision and human cadence (runs after phase 156)."""
    if not text:
        return text
    out = str(text)
    for old, new in PHASE162_ORDERED_REPLACEMENTS:
        if old in out:
            out = out.replace(old, new)
    out = re.sub(r"\s{2,}", " ", out).strip()
    return out
