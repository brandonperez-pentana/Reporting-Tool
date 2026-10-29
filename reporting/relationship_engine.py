from __future__ import annotations

from dataclasses import dataclass
from typing import Any


SIGNIFICANCE_ORDER = {"not_available": 0, "negligible": 1, "slight": 2, "meaningful": 3, "strong": 4}


@dataclass
class RelationshipAnalysis:
    relationship_confidence_score: float
    relationship_rule_trigger: list[str]
    relationship_reason_trace: list[str]
    stage_pressure: list[str]
    priority_insight: str | None
    summary_clause: str | None
    sms_clause: str | None
    implication_clause: str | None
    suppressed_implication: bool
    causal_tier: str = "medium"

    def as_dict(self) -> dict[str, Any]:
        return {
            "relationship_confidence_score": self.relationship_confidence_score,
            "relationship_rule_trigger": list(self.relationship_rule_trigger),
            "relationship_reason_trace": list(self.relationship_reason_trace),
            "stage_pressure": list(self.stage_pressure),
            "priority_insight": self.priority_insight or "",
            "summary_clause": self.summary_clause or "",
            "sms_clause": self.sms_clause or "",
            "implication_clause": self.implication_clause or "",
            "suppressed_implication": self.suppressed_implication,
            "causal_tier": self.causal_tier,
        }


def _metric_usable(metric: dict[str, Any] | None) -> bool:
    if not isinstance(metric, dict):
        return False
    if bool(metric.get("is_suspicious", False)):
        # Mirrors metric_usable_for_reporting: a metric the engine already distrusts
        # must not drive a cross-metric claim either ("conversions firmed" off a
        # conversion rate parsed as 125%).
        return False
    if bool(metric.get("usable_for_reporting", False)):
        return True
    return bool(metric.get("is_reliable", False))


def _metric_direction(metric: dict[str, Any] | None) -> str:
    if not isinstance(metric, dict):
        return "flat"
    return str(metric.get("direction", "flat"))


def _metric_sig(metric: dict[str, Any] | None) -> str:
    if not isinstance(metric, dict):
        return "not_available"
    return str(metric.get("significance_band", "not_available"))


def _sig_at_least(metric: dict[str, Any] | None, threshold: str = "slight") -> bool:
    return SIGNIFICANCE_ORDER.get(_metric_sig(metric), 0) >= SIGNIFICANCE_ORDER.get(threshold, 2)


def _count_usable(metrics: dict[str, dict[str, Any]]) -> int:
    return sum(1 for metric in metrics.values() if _metric_usable(metric))


def analyze_operational_relationships(
    metrics: dict[str, dict[str, Any]],
    sentiment: str,
    context_seed: str = "",
) -> RelationshipAnalysis:
    # context_seed is accepted for deterministic traceability at caller boundaries.
    _ = context_seed
    rule_triggers: list[str] = []
    reason_trace: list[str] = []
    stage_pressure: set[str] = set()
    confidence = 0.0

    leads = metrics.get("sales_total_leads")
    deals = metrics.get("sales_total_deals")
    conv = metrics.get("sales_overall_conversion_rate")
    bookings = metrics.get("service_bookings_from_campaigns")
    svc_roi = metrics.get("service_campaign_roi")
    enrolled = metrics.get("service_customers_enrolled")
    response = metrics.get("sales_internet_response_time")
    untouched = metrics.get("sales_uncontacted_leads")

    usable_count = _count_usable(metrics)
    reason_trace.append(f"usable_metrics={usable_count}")

    # 1) Demand vs conversion efficiency.
    if _metric_usable(leads) and _metric_usable(conv) and _sig_at_least(leads) and _sig_at_least(conv):
        ldir = _metric_direction(leads)
        cdir = _metric_direction(conv)
        if ldir == "down" and cdir == "up":
            rule_triggers.append("lower_volume_higher_efficiency")
            stage_pressure.update({"demand", "conversion"})
            confidence += 0.58
            reason_trace.append("trigger: leads_down + conversion_up")
        elif ldir == "up" and cdir == "down":
            rule_triggers.append("higher_volume_weaker_efficiency")
            stage_pressure.update({"engagement", "conversion"})
            confidence += 0.56
            reason_trace.append("trigger: leads_up + conversion_down")

    # 2) Throughput mismatch: demand up while deals down.
    if _metric_usable(leads) and _metric_usable(deals) and _sig_at_least(leads) and _sig_at_least(deals):
        if _metric_direction(leads) == "up" and _metric_direction(deals) == "down":
            rule_triggers.append("demand_to_output_mismatch")
            stage_pressure.update({"engagement", "conversion"})
            confidence += 0.52
            reason_trace.append("trigger: leads_up + deals_down")
        elif _metric_direction(leads) == "down" and _metric_direction(deals) == "flat":
            rule_triggers.append("stable_output_under_softer_demand")
            stage_pressure.update({"demand", "conversion"})
            confidence += 0.42
            reason_trace.append("trigger: leads_down + deals_flat")

    # 3a) Service campaign activity vs effectiveness (enrolment is activity, not ROI outcome).
    if _metric_usable(enrolled) and _metric_usable(svc_roi) and _sig_at_least(enrolled) and _sig_at_least(svc_roi):
        if _metric_direction(enrolled) == "up" and _metric_direction(svc_roi) == "down":
            rule_triggers.append("service_enrolled_roi_divergence")
            stage_pressure.update({"retention", "profitability"})
            confidence += 0.52
            reason_trace.append("trigger: enrolled_up + service_roi_down")

    # 3) Campaign bookings vs service ROI (effectiveness vs throughput — observational, not causal).
    if _metric_usable(bookings) and _metric_usable(svc_roi) and _sig_at_least(bookings) and _sig_at_least(svc_roi):
        if _metric_direction(bookings) == "up" and _metric_direction(svc_roi) == "down":
            rule_triggers.append("retention_profitability_mismatch")
            stage_pressure.update({"retention", "profitability"})
            confidence += 0.54
            reason_trace.append("trigger: bookings_up + service_roi_down")
        elif _metric_direction(bookings) == "up" and _metric_direction(svc_roi) == "up":
            rule_triggers.append("retention_profitability_alignment")
            stage_pressure.update({"retention", "profitability"})
            confidence += 0.36
            reason_trace.append("trigger: bookings_up + service_roi_up")

    # 4) Execution pressure relationship.
    if _metric_usable(response) and _sig_at_least(response):
        if _metric_direction(response) == "up":
            rule_triggers.append("execution_latency_pressure")
            stage_pressure.add("execution")
            confidence += 0.30
            reason_trace.append("trigger: response_time_up")
    if _metric_usable(untouched) and _sig_at_least(untouched):
        if _metric_direction(untouched) == "up":
            rule_triggers.append("contact_coverage_pressure")
            stage_pressure.add("execution")
            confidence += 0.30
            reason_trace.append("trigger: uncontacted_leads_up")

    if usable_count >= 10:
        confidence += 0.08
        reason_trace.append("coherence_bonus=high_metric_completeness")
    elif usable_count >= 6:
        confidence += 0.04
        reason_trace.append("coherence_bonus=minimum_metric_completeness")
    else:
        reason_trace.append("coherence_penalty=low_metric_completeness")

    contradiction_pair = {
        "lower_volume_higher_efficiency",
        "higher_volume_weaker_efficiency",
    }
    if contradiction_pair.issubset(set(rule_triggers)):
        confidence -= 0.18
        reason_trace.append("coherence_penalty=contradictory_efficiency_signals")

    confidence = max(0.0, min(1.0, round(confidence, 2)))
    reason_trace.append(f"relationship_confidence_score={confidence}")

    min_gate = usable_count >= 6 and confidence >= 0.62 and bool(rule_triggers)
    suppressed_implication = not min_gate
    implication_clause: str | None = None
    causal_tier = "medium"
    if min_gate:
        if "lower_volume_higher_efficiency" in rule_triggers:
            implication_clause = "Conversions improved while demand softened."
            causal_tier = "high"
        elif "higher_volume_weaker_efficiency" in rule_triggers or "demand_to_output_mismatch" in rule_triggers:
            implication_clause = "Conversion pressure showed through later funnel stages."
            causal_tier = "medium"
        elif "retention_profitability_mismatch" in rule_triggers and "service_enrolled_roi_divergence" in rule_triggers:
            implication_clause = None
            suppressed_implication = True
            causal_tier = "low"
        elif "retention_profitability_mismatch" in rule_triggers:
            implication_clause = None
            suppressed_implication = True
            causal_tier = "low"
        elif "service_enrolled_roi_divergence" in rule_triggers:
            implication_clause = None
            suppressed_implication = True
            causal_tier = "low"

    priority_insight: str | None = None
    summary_clause: str | None = None
    sms_clause: str | None = None

    if min_gate:
        if "lower_volume_higher_efficiency" in rule_triggers:
            priority_insight = "Leads eased while conversions firmed, partly offsetting softer demand."
            summary_clause = "Stronger closing partly cushioned the softer lead flow."
            sms_clause = "Softer leads, firmer conversions."
        elif "higher_volume_weaker_efficiency" in rule_triggers:
            priority_insight = "Enquiry activity increased, while conversion efficiency weakened through the funnel."
            summary_clause = "Higher lead counts did not fully show up in conversion."
            sms_clause = "Volume up, conversion softer."
        elif "demand_to_output_mismatch" in rule_triggers:
            priority_insight = "Enquiry rose while deals did not keep pace—treat closing throughput as the watch area."
            summary_clause = "Enquiry growth outpaced delivered deal outcomes across the month."
            sms_clause = "Leads up, deal conversion lagged."
        elif "retention_profitability_mismatch" in rule_triggers and "service_enrolled_roi_divergence" in rule_triggers:
            priority_insight = "Campaign-linked bookings and enrolled clients both increased while service campaign ROI still declined."
            summary_clause = "Campaign-linked bookings and enrolled clients increased while service campaign ROI still declined."
            sms_clause = "Bookings and enrolments up; service ROI down."
        elif "service_enrolled_roi_divergence" in rule_triggers:
            priority_insight = "Enrolled clients increased while service campaign ROI declined."
            summary_clause = "Enrolled clients increased, but service campaign ROI declined."
            sms_clause = "Enrolments up; service campaign ROI down."
        elif "retention_profitability_mismatch" in rule_triggers:
            priority_insight = "Service campaign bookings increased while service campaign ROI still declined."
            summary_clause = "Service campaign bookings increased while service campaign ROI still declined."
            sms_clause = "Bookings up, ROI softer."
        elif "execution_latency_pressure" in rule_triggers or "contact_coverage_pressure" in rule_triggers:
            priority_insight = "Execution friction remained visible in response-time and lead-contact coverage indicators."
            summary_clause = "Lead response and contact coverage stayed under pressure."
            sms_clause = "Execution pressure persists."

    # Conservative guardrail for very positive sentiment with weak relationship coherence.
    if sentiment in {"strongly_improved", "slightly_improved"} and confidence < 0.7:
        implication_clause = None
        suppressed_implication = True
        reason_trace.append("implication_suppressed=positive_sentiment_with_subcritical_confidence")

    return RelationshipAnalysis(
        relationship_confidence_score=confidence,
        relationship_rule_trigger=rule_triggers,
        relationship_reason_trace=reason_trace,
        stage_pressure=sorted(stage_pressure),
        priority_insight=priority_insight,
        summary_clause=summary_clause,
        sms_clause=sms_clause,
        implication_clause=implication_clause,
        suppressed_implication=suppressed_implication,
        causal_tier=causal_tier,
    )
