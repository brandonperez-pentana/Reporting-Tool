"""Experimental template-smoke helpers — writes only to dedicated smoke output folders."""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from reporter.context import ReportContext, from_dealership_report_txt, from_dealership_result
from reporter.rendering.txt_renderer import (
    normalise_report_text_for_comparison,
    render_dealership_report_txt,
    reports_are_equivalent,
    template_txt_smoke_enabled,
)
from reporter.rendering.docx_renderer import (
    count_insight_occurrences,
    docx_contains_all,
    extract_docx_plain_text,
    render_dealership_report_docx,
    template_docx_smoke_enabled,
)

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SMOKE_OUTPUT_DIR = ROOT / "Insights Output TEMPLATE SMOKE"
DEFAULT_DOCX_SMOKE_OUTPUT_DIR = ROOT / "Insights Output DOCX TEMPLATE SMOKE"

logger = logging.getLogger("reporter.smoke")


@dataclass
class SmokeRenderResult:
    dealership_name: str
    source_path: str | None
    output_path: str
    equivalent: bool
    legacy_text: str
    rendered_text: str
    warnings: list[str] = field(default_factory=list)


@dataclass
class SmokeRunReport:
    smoke_output_dir: str
    results: list[SmokeRenderResult] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    @property
    def all_equivalent(self) -> bool:
        return bool(self.results) and all(r.equivalent for r in self.results) and not self.errors


def require_smoke_enabled() -> None:
    if not template_txt_smoke_enabled():
        raise RuntimeError(
            "Template TXT smoke is disabled. Set REPORTER_TEMPLATE_TXT_SMOKE=1 to enable experimental smoke output."
        )


def smoke_output_dir(path: Path | None = None) -> Path:
    return (path or DEFAULT_SMOKE_OUTPUT_DIR).resolve()


def render_result_smoke(
    result: Any,
    *,
    output_dir: Path | None = None,
    compare_legacy: bool = True,
) -> SmokeRenderResult:
    """Render DealershipResult through template TXT into the smoke output directory."""
    require_smoke_enabled()
    import dealership_insights_local as core

    out_dir = smoke_output_dir(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    legacy = core.build_output_text(result)
    ctx = from_dealership_result(result)
    rendered = render_dealership_report_txt(ctx)

    safe_name = core.safe_file_stem(result.dealership_name)
    out_path = out_dir / f"{safe_name} - {result.month_label}.txt"
    out_path.write_text(rendered + "\n", encoding="utf-8")

    equivalent = reports_are_equivalent(legacy, rendered) if compare_legacy else True
    if not equivalent:
        logger.warning(
            "Template smoke mismatch | dealership=%s | see %s",
            result.dealership_name,
            out_path,
        )

    return SmokeRenderResult(
        dealership_name=result.dealership_name,
        source_path=None,
        output_path=str(out_path),
        equivalent=equivalent,
        legacy_text=legacy,
        rendered_text=rendered,
    )


def render_txt_file_smoke(
    source_txt: Path,
    *,
    output_dir: Path | None = None,
) -> SmokeRenderResult:
    """Read an existing TXT report (read-only), re-render via template, write to smoke dir."""
    require_smoke_enabled()

    out_dir = smoke_output_dir(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    legacy = source_txt.read_text(encoding="utf-8")
    ctx = from_dealership_report_txt(legacy)
    rendered = render_dealership_report_txt(ctx)

    out_path = out_dir / source_txt.name
    out_path.write_text(rendered + "\n", encoding="utf-8")

    equivalent = reports_are_equivalent(legacy, rendered)
    return SmokeRenderResult(
        dealership_name=ctx.dealership_name,
        source_path=str(source_txt),
        output_path=str(out_path),
        equivalent=equivalent,
        legacy_text=legacy,
        rendered_text=rendered,
    )


def run_smoke_on_production_txt_samples(
    source_dir: Path,
    month_glob: str = "* - Jun 2026.txt",
    *,
    output_dir: Path | None = None,
    limit: int | None = None,
) -> SmokeRunReport:
    """Read production TXT files read-only; write template smoke outputs elsewhere."""
    require_smoke_enabled()
    out_dir = smoke_output_dir(output_dir)
    report = SmokeRunReport(smoke_output_dir=str(out_dir))

    files = sorted(source_dir.glob(month_glob), key=lambda p: p.name.lower())
    if limit is not None:
        files = files[:limit]

    for path in files:
        try:
            report.results.append(render_txt_file_smoke(path, output_dir=out_dir))
        except Exception as exc:
            report.errors.append(f"{path.name}: {exc}")

    return report


@dataclass
class DocxSmokeRenderResult:
    dealership_name: str
    source_path: str | None
    output_path: str
    structural_ok: bool
    missing_fragments: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass
class DocxSmokeRunReport:
    smoke_output_dir: str
    results: list[DocxSmokeRenderResult] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    @property
    def all_ok(self) -> bool:
        return bool(self.results) and all(r.structural_ok for r in self.results) and not self.errors


def require_docx_smoke_enabled() -> None:
    if not template_docx_smoke_enabled():
        raise RuntimeError(
            "Template DOCX smoke is disabled. Set REPORTER_TEMPLATE_DOCX_SMOKE=1 to enable experimental smoke output."
        )


def docx_smoke_output_dir(path: Path | None = None) -> Path:
    return (path or DEFAULT_DOCX_SMOKE_OUTPUT_DIR).resolve()


def _structural_docx_fragments(ctx: ReportContext) -> list[str]:
    fragments = [
        ctx.dealership_name,
        ctx.reporting_month,
        str(ctx.reporting_year),
        ctx.summary,
        ctx.sms_text,
        "yes" if ctx.review_required else "no",
        "yes" if ctx.anomaly_flags.anomaly_detected else "no",
    ]
    fragments.extend(pi.text for pi in ctx.priority_insights)
    if ctx.review_reason_tags:
        fragments.extend(ctx.review_reason_tags)
    return [f for f in fragments if f]


def render_docx_result_smoke(
    result: Any,
    *,
    output_dir: Path | None = None,
) -> DocxSmokeRenderResult:
    """Render DealershipResult through template DOCX into the smoke output directory."""
    require_docx_smoke_enabled()
    import dealership_insights_local as core

    out_dir = docx_smoke_output_dir(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    ctx = from_dealership_result(result)
    safe_name = core.safe_file_stem(result.dealership_name)
    out_path = out_dir / f"{safe_name} - {result.month_label}.docx"
    render_dealership_report_docx(ctx, out_path)

    text = extract_docx_plain_text(out_path)
    missing = docx_contains_all(text, _structural_docx_fragments(ctx))
    insight_count = count_insight_occurrences(text, [pi.text for pi in ctx.priority_insights])

    warnings: list[str] = []
    if insight_count != len(ctx.priority_insights):
        warnings.append(f"expected {len(ctx.priority_insights)} insights in DOCX, found {insight_count}")

    return DocxSmokeRenderResult(
        dealership_name=ctx.dealership_name,
        source_path=None,
        output_path=str(out_path),
        structural_ok=not missing and insight_count == len(ctx.priority_insights),
        missing_fragments=missing,
        warnings=warnings,
    )


def render_docx_from_txt_smoke(
    source_txt: Path,
    *,
    output_dir: Path | None = None,
) -> DocxSmokeRenderResult:
    """Read existing TXT report (read-only), render DOCX via template to smoke dir."""
    require_docx_smoke_enabled()

    out_dir = docx_smoke_output_dir(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    ctx = from_dealership_report_txt(source_txt.read_text(encoding="utf-8"))
    out_path = out_dir / (source_txt.stem + ".docx")
    render_dealership_report_docx(ctx, out_path)

    text = extract_docx_plain_text(out_path)
    missing = docx_contains_all(text, _structural_docx_fragments(ctx))
    insight_count = count_insight_occurrences(text, [pi.text for pi in ctx.priority_insights])

    return DocxSmokeRenderResult(
        dealership_name=ctx.dealership_name,
        source_path=str(source_txt),
        output_path=str(out_path),
        structural_ok=not missing and insight_count == len(ctx.priority_insights),
        missing_fragments=missing,
    )


def run_docx_smoke_on_production_txt_samples(
    source_dir: Path,
    month_glob: str = "* - Jun 2026.txt",
    *,
    output_dir: Path | None = None,
    limit: int | None = None,
) -> DocxSmokeRunReport:
    """Read production TXT read-only; write DOCX template smoke outputs elsewhere."""
    require_docx_smoke_enabled()
    out_dir = docx_smoke_output_dir(output_dir)
    report = DocxSmokeRunReport(smoke_output_dir=str(out_dir))

    files = sorted(source_dir.glob(month_glob), key=lambda p: p.name.lower())
    if limit is not None:
        files = files[:limit]

    for path in files:
        try:
            report.results.append(render_docx_from_txt_smoke(path, output_dir=out_dir))
        except Exception as exc:
            report.errors.append(f"{path.name}: {exc}")

    return report
