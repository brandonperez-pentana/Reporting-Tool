"""DOCX renderer for dealership reports using docxtpl — layout only."""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Any

from reporter.context import ReportContext

PACKAGE_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_TEMPLATE_PATH = PACKAGE_ROOT / "templates" / "docx" / "dealership_report_template.docx"


class DocxTemplateDependencyError(ImportError):
    """Raised when docxtpl is not installed."""


def _import_docxtpl():
    try:
        from docxtpl import DocxTemplate
    except ImportError as exc:
        raise DocxTemplateDependencyError(
            "docxtpl is required for DOCX template rendering. Install with: pip install docxtpl"
        ) from exc
    return DocxTemplate


def template_docx_smoke_enabled() -> bool:
    return os.environ.get("REPORTER_TEMPLATE_DOCX_SMOKE", "").strip() in {"1", "true", "yes", "on"}


def _resolve_template_path(template_path: Path | None = None) -> Path:
    path = template_path or DEFAULT_TEMPLATE_PATH
    if not path.is_file():
        raise FileNotFoundError(
            f"DOCX template not found: {path}. Run tools/build_docx_report_template.py to generate it."
        )
    return path


def _prepare_context(context: ReportContext | dict[str, Any]) -> dict[str, Any]:
    if isinstance(context, ReportContext):
        context.validate_for_docx_render()
        return context.to_docx_template_dict()
    required = (
        "dealership_name",
        "reporting_month",
        "reporting_year",
        "priority_insights",
        "sms_text",
        "summary",
    )
    missing = [key for key in required if key not in context]
    if missing:
        raise ValueError(f"Missing required DOCX context fields: {', '.join(missing)}")
    return context


def render_dealership_report_docx(
    context: ReportContext | dict[str, Any],
    output_path: Path | str,
    *,
    template_path: Path | None = None,
) -> Path:
    """Render dealership DOCX from ReportContext. No business logic."""
    DocxTemplate = _import_docxtpl()
    template_ctx = _prepare_context(context)
    tpl_path = _resolve_template_path(template_path)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    doc = DocxTemplate(str(tpl_path))
    doc.render(template_ctx)
    doc.save(str(out))
    return out


def extract_docx_plain_text(docx_path: Path | str) -> str:
    """Extract paragraph text from a DOCX for structural smoke tests."""
    try:
        from docx import Document
    except ImportError as exc:
        raise DocxTemplateDependencyError(
            "python-docx is required to read DOCX output. Install with: pip install python-docx"
        ) from exc

    document = Document(str(docx_path))
    parts: list[str] = []
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if text:
            parts.append(text)
    return "\n".join(parts)


def docx_contains_all(text: str, required_fragments: list[str]) -> list[str]:
    """Return fragments from required_fragments that are missing from extracted DOCX text."""
    haystack = text.replace("\r\n", "\n")
    missing: list[str] = []
    for fragment in required_fragments:
        if fragment and fragment not in haystack:
            missing.append(fragment)
    return missing


def count_insight_occurrences(text: str, insights: list[str]) -> int:
    return sum(1 for insight in insights if insight and insight in text)
