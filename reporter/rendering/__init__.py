"""Output renderers (layout only — no business logic)."""

from reporter.rendering.docx_renderer import render_dealership_report_docx
from reporter.rendering.txt_renderer import render_dealership_report_txt

__all__ = ["render_dealership_report_txt", "render_dealership_report_docx"]
