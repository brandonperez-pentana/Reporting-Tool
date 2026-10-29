"""Structural tests for template-backed DOCX rendering."""

from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path

import dealership_insights_local as core
from reporter.context import from_dealership_result
from reporter.rendering.docx_renderer import (
    DocxTemplateDependencyError,
    count_insight_occurrences,
    docx_contains_all,
    extract_docx_plain_text,
    render_dealership_report_docx,
)
from reporter.smoke import render_docx_from_txt_smoke, render_docx_result_smoke

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = Path(__file__).resolve().parent / "fixtures" / "golden"
DEMO_GOLDEN = FIXTURES / "demo_site_1_jul_2026.txt"


def _sample_result(**overrides: object) -> core.DealershipResult:
    base = core.DealershipResult(
        dealership_name="Example Motors",
        month_label="Jun 2026",
        bullet_insights=[
            "Deal volume increased from 79 to 99 (+25.3%).",
            "Conversion improved from 34% to 44%.",
            "Sales campaign ROI improved from $8,731 to $13,188 (+51%).",
            "Outstanding service tasks were broadly steady at 7.",
        ],
        summary="Jun 2026 showed constructive momentum across volume and conversion.",
        sms_version="DealerSocket June Insights:\ndeal volume up 25% and conversion improved to 44%.",
        report_status="success",
        parsing_status="parsed",
        confidence_score=0.91,
        insights_source="rule-based",
        top_priority_level="High",
        review_required="yes",
        sentiment="positive",
        insight_count_after_deduplication=4,
        overall_month_rating="Constructive",
        primary_strength="Conversion gains",
        primary_risk="Response time",
        recommended_focus="Sustain funnel momentum",
        review_reason_tags=["anomaly_detected", "conflicting_signals"],
        sms_valid=True,
        sms_char_count=72,
        sms_limit=core.SMS_CHAR_LIMIT,
        sms_source="priority_insights",
    )
    for key, value in overrides.items():
        setattr(base, key, value)
    return base


class DocxTemplateRenderingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        template = ROOT / "templates" / "docx" / "dealership_report_template.docx"
        if not template.is_file():
            from tools.build_docx_report_template import build_template

            build_template()

    def test_render_contains_dealership_and_month(self) -> None:
        result = _sample_result()
        ctx = from_dealership_result(result)
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "report.docx"
            render_dealership_report_docx(ctx, out)
            text = extract_docx_plain_text(out)
            self.assertIn("Example Motors", text)
            self.assertIn("Jun 2026", text)
            self.assertIn("2026", text)

    def test_render_contains_four_priority_insights(self) -> None:
        result = _sample_result()
        ctx = from_dealership_result(result)
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "report.docx"
            render_dealership_report_docx(ctx, out)
            text = extract_docx_plain_text(out)
            self.assertEqual(count_insight_occurrences(text, list(result.bullet_insights)), 4)

    def test_sms_text_unchanged_in_docx(self) -> None:
        sms = "DealerSocket June Insights:\nleads rose 12% and ROI improved 8%."
        result = _sample_result(sms_version=sms)
        ctx = from_dealership_result(result)
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "report.docx"
            render_dealership_report_docx(ctx, out)
            text = extract_docx_plain_text(out)
            self.assertIn(sms, text)

    def test_review_and_anomaly_flags_visible(self) -> None:
        result = _sample_result(review_reason_tags=["anomaly_detected"])
        ctx = from_dealership_result(result)
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "report.docx"
            render_dealership_report_docx(ctx, out)
            text = extract_docx_plain_text(out)
            self.assertIn("Review Required: yes", text)
            self.assertIn("Anomaly Detected: yes", text)
            self.assertIn("anomaly_detected", text)

    def test_summary_present_unchanged(self) -> None:
        result = _sample_result()
        ctx = from_dealership_result(result)
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "report.docx"
            render_dealership_report_docx(ctx, out)
            text = extract_docx_plain_text(out)
            self.assertIn(result.summary, text)

    def test_missing_required_context_raises(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "report.docx"
            with self.assertRaises(ValueError):
                render_dealership_report_docx({"dealership_name": "X"}, out)

    def test_success_status_requires_four_insights(self) -> None:
        result = _sample_result(bullet_insights=["only one insight"])
        ctx = from_dealership_result(result)
        with self.assertRaises(ValueError):
            ctx.validate_for_docx_render()

    def test_docx_smoke_writes_only_to_temp_dir(self) -> None:
        os.environ["REPORTER_TEMPLATE_DOCX_SMOKE"] = "1"
        try:
            result = _sample_result()
            with tempfile.TemporaryDirectory() as tmp:
                smoke = render_docx_result_smoke(result, output_dir=Path(tmp))
                self.assertTrue(Path(smoke.output_path).is_file())
                self.assertTrue(smoke.structural_ok)
        finally:
            os.environ.pop("REPORTER_TEMPLATE_DOCX_SMOKE", None)

    def test_docx_smoke_from_golden_txt(self) -> None:
        if not DEMO_GOLDEN.is_file():
            self.skipTest("golden fixture missing")
        os.environ["REPORTER_TEMPLATE_DOCX_SMOKE"] = "1"
        try:
            with tempfile.TemporaryDirectory() as tmp:
                smoke = render_docx_from_txt_smoke(DEMO_GOLDEN, output_dir=Path(tmp))
                self.assertTrue(smoke.structural_ok, smoke.missing_fragments)
        finally:
            os.environ.pop("REPORTER_TEMPLATE_DOCX_SMOKE", None)

    def test_structural_helper_detects_missing_fragment(self) -> None:
        missing = docx_contains_all("hello world", ["hello", "missing"])
        self.assertEqual(missing, ["missing"])


if __name__ == "__main__":
    unittest.main()
