"""Golden/equivalence tests for template-backed TXT rendering."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import dealership_insights_local as core
from reporter.context import from_dealership_result, from_dealership_report_txt
from reporter.rendering.txt_renderer import (
    normalise_report_text_for_comparison,
    render_dealership_report_txt,
    reports_are_equivalent,
)
from reporter.smoke import render_result_smoke, render_txt_file_smoke

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
        review_reason_tags=["anomaly_detected"],
        extracted_metrics=["sales_total_deals", "sales_overall_conversion_rate"],
        rules_triggered=["sales_volume_snapshot"],
        notes=["Sample note for template rendering."],
        sms_valid=True,
        sms_char_count=72,
        sms_limit=core.SMS_CHAR_LIMIT,
        sms_source="priority_insights",
    )
    for key, value in overrides.items():
        setattr(base, key, value)
    return base


class TxtTemplateRenderingTests(unittest.TestCase):
    def test_template_matches_build_output_text(self) -> None:
        result = _sample_result()
        legacy = core.build_output_text(result)
        ctx = from_dealership_result(result)
        rendered = render_dealership_report_txt(ctx)
        self.assertTrue(reports_are_equivalent(legacy, rendered))

    def test_sms_text_not_altered_by_template(self) -> None:
        sms = "DealerSocket June Insights:\nleads rose 12% and ROI improved 8%."
        result = _sample_result(sms_version=sms)
        ctx = from_dealership_result(result)
        rendered = render_dealership_report_txt(ctx)
        self.assertIn(sms, rendered)
        self.assertNotIn(result.summary, sms)

    def test_priority_insight_texts_preserved(self) -> None:
        result = _sample_result()
        ctx = from_dealership_result(result)
        rendered = render_dealership_report_txt(ctx)
        for insight in result.bullet_insights:
            self.assertIn(insight, rendered)

    def test_exactly_four_priority_insights_in_output(self) -> None:
        result = _sample_result()
        rendered = render_dealership_report_txt(from_dealership_result(result))
        block = rendered.split("Priority Insights:", 1)[1].split("Summary:", 1)[0]
        lines = [ln for ln in block.strip().splitlines() if ln.strip()]
        self.assertEqual(len(lines), 4)

    def test_review_and_anomaly_flags_in_output(self) -> None:
        result = _sample_result(review_reason_tags=["anomaly_detected", "conflicting_signals"])
        rendered = render_dealership_report_txt(from_dealership_result(result))
        self.assertIn("Review Required: yes", rendered)
        self.assertIn("anomaly_detected", rendered)
        self.assertIn("conflicting_signals", rendered)

    def test_golden_fixture_round_trip(self) -> None:
        if not DEMO_GOLDEN.is_file():
            self.skipTest("golden fixture missing")
        golden = DEMO_GOLDEN.read_text(encoding="utf-8")
        ctx = from_dealership_report_txt(golden)
        rendered = render_dealership_report_txt(ctx)
        self.assertTrue(reports_are_equivalent(golden, rendered))

    def test_smoke_writes_only_to_temp_dir(self) -> None:
        import os

        os.environ["REPORTER_TEMPLATE_TXT_SMOKE"] = "1"
        try:
            result = _sample_result()
            with tempfile.TemporaryDirectory() as tmp:
                smoke = render_result_smoke(result, output_dir=Path(tmp))
                self.assertTrue(Path(smoke.output_path).is_file())
                self.assertTrue(smoke.equivalent)
                self.assertTrue(reports_are_equivalent(smoke.legacy_text, smoke.rendered_text))
        finally:
            os.environ.pop("REPORTER_TEMPLATE_TXT_SMOKE", None)

    def test_smoke_from_txt_read_only(self) -> None:
        import os

        if not DEMO_GOLDEN.is_file():
            self.skipTest("golden fixture missing")
        os.environ["REPORTER_TEMPLATE_TXT_SMOKE"] = "1"
        try:
            with tempfile.TemporaryDirectory() as tmp:
                smoke = render_txt_file_smoke(DEMO_GOLDEN, output_dir=Path(tmp))
                self.assertTrue(smoke.equivalent)
        finally:
            os.environ.pop("REPORTER_TEMPLATE_TXT_SMOKE", None)

    def test_normalised_comparison_ignores_trailing_newlines(self) -> None:
        a = "Dealership: X\nMonth: Jun 2026\n"
        b = "Dealership: X\nMonth: Jun 2026"
        self.assertEqual(normalise_report_text_for_comparison(a), normalise_report_text_for_comparison(b))


if __name__ == "__main__":
    unittest.main()
