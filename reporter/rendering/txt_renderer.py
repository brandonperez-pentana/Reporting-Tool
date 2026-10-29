"""Jinja TXT renderer for dealership reports — layout only."""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Any

from jinja2 import Environment, StrictUndefined, TemplateError

from reporter.context import ReportContext

PACKAGE_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_TEMPLATE_PATH = PACKAGE_ROOT / "templates" / "txt" / "dealership_report.txt.j2"

_TEMPLATE_ENV = Environment(undefined=StrictUndefined, autoescape=False, keep_trailing_newline=False)


def template_txt_smoke_enabled() -> bool:
    return os.environ.get("REPORTER_TEMPLATE_TXT_SMOKE", "").strip() in {"1", "true", "yes", "on"}


def _load_template(path: Path | None = None) -> str:
    template_path = path or DEFAULT_TEMPLATE_PATH
    if not template_path.is_file():
        raise FileNotFoundError(f"TXT template not found: {template_path}")
    return template_path.read_text(encoding="utf-8")


def _normalise_trailing_newlines(text: str) -> str:
    """Match legacy build_output_text: single trailing newline only when content present."""
    return text.rstrip("\n")


def render_dealership_report_txt(
    context: ReportContext | dict[str, Any],
    *,
    template_path: Path | None = None,
) -> str:
    """Render dealership TXT from ReportContext. No business logic."""
    if isinstance(context, dict):
        template_ctx = context
    else:
        context.validate_for_txt_render()
        template_ctx = context.to_template_dict()

    required = ("dealership_name", "reporting_month", "priority_insight_texts", "sms_text", "summary")
    missing = [key for key in required if key not in template_ctx]
    if missing:
        raise ValueError(f"Missing required TXT context fields: {', '.join(missing)}")

    template_src = _load_template(template_path)
    try:
        rendered = _TEMPLATE_ENV.from_string(template_src).render(**template_ctx)
    except TemplateError as exc:
        raise ValueError(f"Unable to render dealership TXT template: {exc}") from exc

    return _normalise_trailing_newlines(rendered)


def normalise_report_text_for_comparison(text: str) -> str:
    """Normalise line endings and trailing whitespace for equivalence checks."""
    lines = [line.rstrip() for line in text.replace("\r\n", "\n").split("\n")]
    return "\n".join(lines).strip()


def reports_are_equivalent(legacy: str, rendered: str) -> bool:
    return normalise_report_text_for_comparison(legacy) == normalise_report_text_for_comparison(rendered)
