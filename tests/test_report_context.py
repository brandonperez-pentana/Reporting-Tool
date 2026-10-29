"""Tests for ReportContext contract."""

from __future__ import annotations

import json
import unittest

import dealership_insights_local as core
from reporter.context import (
    PRIORITY_INSIGHT_COUNT,
    ReportContext,
    from_dealership_result,
)


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
        review_reason_tags=["anomaly_detected", "slow_response_time"],
        extracted_metrics=["sales_total_deals", "sales_overall_conversion_rate"],
        rules_triggered=["sales_volume_snapshot"],
        sms_valid=True,
        sms_char_count=72,
        sms_limit=core.SMS_CHAR_LIMIT,
        sms_source="priority_insights",
    )
    for key, value in overrides.items():
        setattr(base, key, value)
    return base


class ReportContextTests(unittest.TestCase):
    def test_from_dealership_result_creates_context(self) -> None:
        result = _sample_result()
        ctx = from_dealership_result(result)
        self.assertEqual(ctx.dealership_name, "Example Motors")
        self.assertEqual(ctx.reporting_month, "Jun 2026")
        self.assertEqual(ctx.reporting_year, 2026)
        self.assertEqual(len(ctx.priority_insights), PRIORITY_INSIGHT_COUNT)

    def test_four_priority_insights_enforced_on_success(self) -> None:
        result = _sample_result()
        ctx = from_dealership_result(result)
        ctx.validate_for_txt_render()
        texts = [pi.text for pi in ctx.priority_insights]
        self.assertEqual(texts, result.bullet_insights)

    def test_success_requires_four_insights_for_txt_render(self) -> None:
        result = _sample_result(bullet_insights=["only one"])
        ctx = from_dealership_result(result)
        with self.assertRaises(ValueError):
            ctx.validate_for_txt_render()

    def test_sms_text_unchanged_in_context(self) -> None:
        sms = "DealerSocket June Insights:\nquantified body here."
        result = _sample_result(sms_version=sms)
        ctx = from_dealership_result(result)
        self.assertEqual(ctx.sms_text, sms)

    def test_review_and_anomaly_flags_visible(self) -> None:
        result = _sample_result(review_required="yes", review_reason_tags=["anomaly_detected"])
        ctx = from_dealership_result(result)
        self.assertTrue(ctx.review_required)
        self.assertTrue(ctx.anomaly_flags.anomaly_detected)
        self.assertIn("anomaly_detected", ctx.review_reason_tags)

    def test_to_dict_is_json_serialisable(self) -> None:
        ctx = from_dealership_result(_sample_result())
        payload = ctx.to_dict()
        json.dumps(payload)
        self.assertEqual(payload["dealership_name"], "Example Motors")
        self.assertEqual(len(payload["priority_insights"]), 4)

    def test_template_dict_exposes_review_display(self) -> None:
        ctx = from_dealership_result(_sample_result())
        td = ctx.to_template_dict()
        self.assertEqual(td["review_required_display"], "yes")
        self.assertIn("anomaly_detected", td["review_reason_tags_display"])

    def test_docx_template_dict_includes_insight_objects(self) -> None:
        ctx = from_dealership_result(_sample_result())
        dd = ctx.to_docx_template_dict()
        self.assertEqual(len(dd["priority_insights"]), 4)
        self.assertEqual(dd["priority_insights"][0]["text"], ctx.priority_insights[0].text)
        self.assertIn("anomaly_detected_display", dd)


if __name__ == "__main__":
    unittest.main()
