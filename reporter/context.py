"""Minimal ReportContext contract for template-backed rendering."""

from __future__ import annotations

import re
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

PRIORITY_INSIGHT_COUNT = 4
CONTEXT_SCHEMA_VERSION = "1.0.0-smoke"


@dataclass(frozen=True)
class SourceFiles:
    sales_pdf: str | None
    service_pdf: str | None
    sales_pdf_status: str
    service_pdf_status: str


@dataclass(frozen=True)
class MetricsSummary:
    metric_count: int
    top_movements: list[str]
    anomaly_count: int
    unreliable_count: int


@dataclass(frozen=True)
class PriorityInsight:
    index: int
    text: str
    is_placeholder: bool = False
    theme_ids: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class SmsValidation:
    is_valid: bool
    char_count: int
    char_limit: int
    format_ok: bool
    quantified: bool
    pi_traceable: bool
    semantic_issues: list[str]
    source: str


@dataclass(frozen=True)
class AnomalyFlags:
    anomaly_detected: bool
    conflicting_signals: bool
    unreliable_metrics: bool
    anomaly_count: int
    anomaly_types: list[str]


@dataclass(frozen=True)
class ExecutiveSummary:
    overall_month_rating: str
    primary_strength: str
    primary_risk: str
    recommended_focus: str
    sentiment: str


@dataclass(frozen=True)
class ValidationResults:
    report_status: str
    parsing_status: str
    confidence_score: float
    insights_source: str
    top_priority_level: str
    insight_count_after_deduplication: int
    pipeline_notes: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class OutputPaths:
    txt_path: str | None = None
    docx_path: str | None = None
    output_dir: str | None = None


@dataclass(frozen=True)
class RunMetadata:
    run_id: str
    generated_at: str
    pipeline_version: str
    accounts_folder: str
    context_schema_version: str


@dataclass(frozen=True)
class ReportContext:
    reporting_month: str
    reporting_year: int
    dealership_name: str
    dealership_group: str | None
    account_type: str
    source_files: SourceFiles
    sales_metrics: MetricsSummary | None
    service_metrics: MetricsSummary | None
    priority_insights: tuple[PriorityInsight, ...]
    summary: str
    sms_text: str
    sms_validation: SmsValidation
    executive_summary: ExecutiveSummary
    anomaly_flags: AnomalyFlags
    review_required: bool
    review_reason_tags: list[str]
    review_severity: str
    output_paths: OutputPaths
    validation_results: ValidationResults
    run_metadata: RunMetadata
    extracted_metrics: tuple[str, ...] = ()
    rules_triggered: tuple[str, ...] = ()
    notes: tuple[str, ...] = ()

    def validate_for_txt_render(self) -> None:
        """Fail clearly when required TXT fields are missing or invalid."""
        if not self.dealership_name.strip():
            raise ValueError("dealership_name is required for TXT rendering")
        if not self.reporting_month.strip():
            raise ValueError("reporting_month is required for TXT rendering")
        if self.validation_results.report_status == "success":
            if len(self.priority_insights) != PRIORITY_INSIGHT_COUNT:
                raise ValueError(
                    f"success status requires exactly {PRIORITY_INSIGHT_COUNT} priority insights; "
                    f"got {len(self.priority_insights)}"
                )

    def validate_for_docx_render(self) -> None:
        """Fail clearly when required DOCX fields are missing or invalid."""
        self.validate_for_txt_render()

    def to_template_dict(self) -> dict[str, Any]:
        """Flat-ish dict for Jinja — presentation fields only."""
        review_tags_display = " | ".join(self.review_reason_tags) if self.review_reason_tags else "none"
        insight_texts = [pi.text for pi in self.priority_insights]
        optional_suffix = ""
        if self.extracted_metrics:
            optional_suffix += "\n\nExtracted Metrics:\n" + "\n".join(f"- {m}" for m in self.extracted_metrics)
        if self.rules_triggered:
            optional_suffix += "\n\nRules Triggered:\n" + "\n".join(f"- {r}" for r in self.rules_triggered)
        if self.notes:
            optional_suffix += "\n\nNotes:\n" + "\n".join(f"- {n}" for n in self.notes)
        placeholder_warning_banner = ""
        if self.validation_results.parsing_status == "fallback" or self.validation_results.insights_source == "placeholder":
            placeholder_warning_banner = (
                "!" * 70 + "\n"
                "!! NO SOURCE DATA — THIS IS A PLACEHOLDER REPORT, NOT A REAL ONE. !!\n"
                "!! Sales/Service PDFs were missing this month. Everything below is\n"
                "!! generic filler text, not derived from any figures. DO NOT SEND.\n"
                + "!" * 70 + "\n"
            )
        return {
            "dealership_name": self.dealership_name,
            "reporting_month": self.reporting_month,
            "reporting_year": self.reporting_year,
            "dealership_group": self.dealership_group,
            "account_type": self.account_type,
            "executive_summary": asdict(self.executive_summary),
            "validation_results": asdict(self.validation_results),
            "review_required_display": "yes" if self.review_required else "no",
            "review_reason_tags_display": review_tags_display,
            "review_severity": self.review_severity,
            "placeholder_warning_banner": placeholder_warning_banner,
            "priority_insight_texts": insight_texts,
            "priority_insights_block": "\n".join(insight_texts),
            "summary": self.summary,
            "sms_text": self.sms_text,
            "source_files": asdict(self.source_files),
            "optional_suffix": optional_suffix,
            "extracted_metrics": list(self.extracted_metrics),
            "rules_triggered": list(self.rules_triggered),
            "notes": list(self.notes),
            "anomaly_flags": asdict(self.anomaly_flags),
            "sms_validation": asdict(self.sms_validation),
            "review_required": self.review_required,
            "review_reason_tags": list(self.review_reason_tags),
        }

    def to_docx_template_dict(self) -> dict[str, Any]:
        """Context dict for docxtpl — includes nested structures for loops."""
        base = self.to_template_dict()
        base["priority_insights"] = [
            {"index": pi.index, "text": pi.text, "is_placeholder": pi.is_placeholder}
            for pi in self.priority_insights
        ]
        base["anomaly_detected_display"] = "yes" if self.anomaly_flags.anomaly_detected else "no"
        if base.get("source_files"):
            base["source_files"] = dict(base["source_files"])
            base["source_files"]["sales_pdf"] = base["source_files"].get("sales_pdf") or "Missing"
            base["source_files"]["service_pdf"] = base["source_files"].get("service_pdf") or "Missing"
        return base

    def to_dict(self) -> dict[str, Any]:
        return _json_safe(asdict(self))


def _json_safe(value: Any) -> Any:
    if isinstance(value, dict):
        return {k: _json_safe(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_safe(v) for v in value]
    if isinstance(value, Path):
        return str(value)
    return value


def parse_month_year(month_label: str) -> tuple[int, int, str]:
    cleaned = re.sub(r"\s+", " ", (month_label or "").strip())
    for fmt in ("%b %Y", "%B %Y"):
        try:
            parsed = datetime.strptime(cleaned, fmt)
            return parsed.month, parsed.year, parsed.strftime("%b %Y")
        except ValueError:
            continue
    raise ValueError(f"Unsupported month label: {month_label!r}")


def _tags_include(tags: list[str] | None, needle: str) -> bool:
    return any(needle in (t or "").lower() for t in (tags or []))


def from_dealership_result(
    result: Any,
    *,
    run_metadata: RunMetadata | None = None,
    output_paths: OutputPaths | None = None,
) -> ReportContext:
    """Build ReportContext from an existing DealershipResult without changing business decisions."""
    import dealership_insights_local as core

    _, year, normalised_month = core.parse_month_label(result.month_label)
    tags = list(result.review_reason_tags or [])
    sales_path = str(result.sales_pdf) if result.sales_pdf else None
    service_path = str(result.service_pdf) if result.service_pdf else None

    if result.report_status == "success" and result.insights_source == "placeholder":
        account_type = "placeholder"
    elif result.report_status == "missing file" or (result.sales_pdf is None and result.service_pdf is None):
        account_type = "missing_pdf"
    else:
        account_type = "standard"

    insights = tuple(
        PriorityInsight(index=i + 1, text=text, is_placeholder=text == core.NO_ADDITIONAL_INSIGHT_MARKER)
        for i, text in enumerate(result.bullet_insights or [])
    )

    sms_text = result.sms_version or ""
    char_count = len(sms_text)
    semantic_issues: list[str] = []
    if result.report_status == "success" and sms_text:
        semantic_issues = list(core.validate_sms_semantics(sms_text, result.bullet_insights or []))

    sms_validation = SmsValidation(
        is_valid=bool(result.sms_valid) and char_count <= int(result.sms_limit or core.SMS_CHAR_LIMIT),
        char_count=char_count,
        char_limit=int(result.sms_limit or core.SMS_CHAR_LIMIT),
        format_ok=sms_text.strip().lower().startswith("dealersocket"),
        quantified=bool(re.search(r"\d", sms_text)),
        pi_traceable=not semantic_issues,
        semantic_issues=semantic_issues,
        source=str(result.sms_source or "priority_insights"),
    )

    meta = run_metadata or RunMetadata(
        run_id="smoke",
        generated_at=datetime.now(timezone.utc).isoformat(),
        pipeline_version="reporter-template-smoke",
        accounts_folder="",
        context_schema_version=CONTEXT_SCHEMA_VERSION,
    )

    return ReportContext(
        reporting_month=normalised_month,
        reporting_year=year,
        dealership_name=result.dealership_name,
        dealership_group=None,
        account_type=account_type,
        source_files=SourceFiles(
            sales_pdf=sales_path,
            service_pdf=service_path,
            sales_pdf_status="found" if sales_path else "missing",
            service_pdf_status="found" if service_path else "missing",
        ),
        sales_metrics=None,
        service_metrics=None,
        priority_insights=insights,
        summary=result.summary or "",
        sms_text=sms_text,
        sms_validation=sms_validation,
        executive_summary=ExecutiveSummary(
            overall_month_rating=result.overall_month_rating,
            primary_strength=result.primary_strength,
            primary_risk=result.primary_risk,
            recommended_focus=result.recommended_focus,
            sentiment=result.sentiment,
        ),
        anomaly_flags=AnomalyFlags(
            anomaly_detected=_tags_include(tags, "anomaly_detected"),
            conflicting_signals=_tags_include(tags, "conflicting_signals"),
            unreliable_metrics=_tags_include(tags, "unreliable_metrics"),
            anomaly_count=int(result.anomaly_count or 0),
            anomaly_types=list(result.anomaly_types or []),
        ),
        review_required=str(result.review_required).lower() == "yes",
        review_reason_tags=tags,
        review_severity=getattr(result, "review_severity", None)
        or core.classify_review_severity(tags, result.parsing_status, result.insights_source),
        output_paths=output_paths or OutputPaths(),
        validation_results=ValidationResults(
            report_status=result.report_status,
            parsing_status=result.parsing_status,
            confidence_score=float(result.confidence_score),
            insights_source=result.insights_source,
            top_priority_level=result.top_priority_level,
            insight_count_after_deduplication=int(result.insight_count_after_deduplication),
            pipeline_notes=list(result.notes or [])[:0],
        ),
        run_metadata=meta,
        extracted_metrics=tuple(result.extracted_metrics or ()),
        rules_triggered=tuple(result.rules_triggered or ()),
        notes=tuple(result.notes or ()),
    )


def from_dealership_report_txt(text: str) -> ReportContext:
    """Parse an existing dealership TXT report into ReportContext (read-only reconstruction)."""
    lines = text.splitlines()
    if not lines:
        raise ValueError("empty report text")

    def _value_after(prefix: str, line: str) -> str:
        return line.split(":", 1)[1].strip() if ":" in line else ""

    dealership_name = ""
    reporting_month = ""
    exec_rating = exec_strength = exec_risk = exec_focus = ""
    report_status = parsing_status = insights_source = top_priority = ""
    review_required_str = "no"
    sentiment = ""
    confidence_score = 0.0
    insight_dedup_count = 0
    review_tags: list[str] = []
    summary = ""
    sms_text = ""
    sales_pdf: str | None = None
    service_pdf: str | None = None
    extracted: list[str] = []
    rules: list[str] = []
    notes: list[str] = []
    priority_insights: list[str] = []

    section: str | None = None
    sms_lines: list[str] = []

    for line in lines:
        if line.startswith("Dealership:"):
            dealership_name = _value_after("Dealership", line)
            section = None
            continue
        if line.startswith("Month:"):
            reporting_month = _value_after("Month", line)
            section = None
            continue
        if line == "Executive Summary:":
            section = "executive"
            continue
        if line.startswith("Status:"):
            report_status = _value_after("Status", line)
            section = None
            continue
        if line.startswith("Parsing Status:"):
            parsing_status = _value_after("Parsing Status", line)
            continue
        if line.startswith("Confidence Score:"):
            try:
                confidence_score = float(_value_after("Confidence Score", line))
            except ValueError:
                confidence_score = 0.0
            continue
        if line.startswith("Insights Source:"):
            insights_source = _value_after("Insights Source", line)
            continue
        if line.startswith("Top Priority Level:"):
            top_priority = _value_after("Top Priority Level", line)
            continue
        if line.startswith("Review Required:"):
            review_required_str = _value_after("Review Required", line)
            continue
        if line.startswith("Sentiment:"):
            sentiment = _value_after("Sentiment", line)
            continue
        if line.startswith("Insight Count After Deduplication:"):
            try:
                insight_dedup_count = int(_value_after("Insight Count After Deduplication", line))
            except ValueError:
                insight_dedup_count = 0
            continue
        if line.startswith("Review Reason Tags:"):
            raw = _value_after("Review Reason Tags", line)
            review_tags = [] if raw == "none" else [t.strip() for t in raw.split("|") if t.strip()]
            continue
        if line == "Priority Insights:":
            section = "insights"
            continue
        if line == "Summary:":
            section = "summary"
            summary = ""
            continue
        if line == "SMS Version:":
            section = "sms"
            sms_lines = []
            continue
        if line == "Source Files:":
            if section == "sms":
                sms_text = "\n".join(sms_lines).strip()
            section = "source"
            continue
        if line.startswith("Sales PDF:"):
            val = _value_after("Sales PDF", line)
            sales_pdf = None if val == "Missing" else val
            continue
        if line.startswith("Service PDF:"):
            val = _value_after("Service PDF", line)
            service_pdf = None if val == "Missing" else val
            continue
        if line == "Extracted Metrics:":
            section = "extracted"
            extracted = []
            continue
        if line == "Rules Triggered:":
            section = "rules"
            rules = []
            continue
        if line == "Notes:":
            section = "notes"
            notes = []
            continue

        if section == "executive" and line.startswith("- Overall Month Rating:"):
            exec_rating = _value_after("- Overall Month Rating", line)
        elif section == "executive" and line.startswith("- Primary Strength:"):
            exec_strength = _value_after("- Primary Strength", line)
        elif section == "executive" and line.startswith("- Primary Risk:"):
            exec_risk = _value_after("- Primary Risk", line)
        elif section == "executive" and line.startswith("- Recommended Focus:"):
            exec_focus = _value_after("- Recommended Focus", line)
        elif section == "insights" and line.strip():
            priority_insights.append(line.strip())
        elif section == "summary":
            if line.strip():
                summary = f"{summary}\n{line}".strip() if summary else line
        elif section == "sms":
            sms_lines.append(line)
        elif section == "extracted" and line.startswith("- "):
            extracted.append(line[2:].strip())
        elif section == "rules" and line.startswith("- "):
            rules.append(line[2:].strip())
        elif section == "notes" and line.startswith("- "):
            notes.append(line[2:].strip())

    if section == "sms" and not sms_text:
        sms_text = "\n".join(sms_lines).strip()

    _, year, normalised_month = parse_month_year(reporting_month)
    insights = tuple(PriorityInsight(index=i + 1, text=t) for i, t in enumerate(priority_insights))

    class _MinimalResult:
        pass

    r = _MinimalResult()
    r.dealership_name = dealership_name
    r.month_label = normalised_month
    r.bullet_insights = list(priority_insights)
    r.summary = summary
    r.sms_version = sms_text
    r.report_status = report_status or "success"
    r.sales_pdf = sales_pdf
    r.service_pdf = service_pdf
    r.parsing_status = parsing_status or "parsed"
    r.confidence_score = confidence_score
    r.extracted_metrics = extracted
    r.rules_triggered = rules
    r.insights_source = insights_source or "rule-based"
    r.top_priority_level = top_priority or "Low"
    r.review_required = review_required_str
    r.sentiment = sentiment or "stable"
    r.insight_count_after_deduplication = insight_dedup_count or len(priority_insights)
    r.overall_month_rating = exec_rating or "Stable"
    r.primary_strength = exec_strength or ""
    r.primary_risk = exec_risk or ""
    r.recommended_focus = exec_focus or ""
    r.review_reason_tags = review_tags
    r.sms_valid = len(sms_text) <= 160
    r.sms_limit = 160
    r.sms_source = "priority_insights"
    r.anomaly_count = 0
    r.anomaly_types = []
    r.notes = notes

    ctx = from_dealership_result(r)
    # Preserve parsed lists exactly (from_dealership_result copies notes into both notes and validation)
    return ReportContext(
        reporting_month=ctx.reporting_month,
        reporting_year=ctx.reporting_year,
        dealership_name=ctx.dealership_name,
        dealership_group=ctx.dealership_group,
        account_type=ctx.account_type,
        source_files=ctx.source_files,
        sales_metrics=ctx.sales_metrics,
        service_metrics=ctx.service_metrics,
        priority_insights=insights,
        summary=summary,
        sms_text=sms_text,
        sms_validation=ctx.sms_validation,
        executive_summary=ExecutiveSummary(
            overall_month_rating=exec_rating or "Stable",
            primary_strength=exec_strength or "",
            primary_risk=exec_risk or "",
            recommended_focus=exec_focus or "",
            sentiment=sentiment or "stable",
        ),
        anomaly_flags=ctx.anomaly_flags,
        review_required=ctx.review_required,
        review_reason_tags=review_tags,
        review_severity=ctx.review_severity,
        output_paths=ctx.output_paths,
        validation_results=ValidationResults(
            report_status=report_status or "success",
            parsing_status=parsing_status or "parsed",
            confidence_score=confidence_score,
            insights_source=insights_source or "rule-based",
            top_priority_level=top_priority or "Low",
            insight_count_after_deduplication=insight_dedup_count or len(priority_insights),
        ),
        run_metadata=ctx.run_metadata,
        extracted_metrics=tuple(extracted),
        rules_triggered=tuple(rules),
        notes=tuple(notes),
    )
