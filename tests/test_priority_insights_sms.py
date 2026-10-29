"""Tests for four priority insights finalisation and SMS-from-insights generation."""

from __future__ import annotations

import unittest

import dealership_insights_local as core
import email_assembler
import email_builder


FOUR = [
    "Deal volume increased from 79 to 99 (+25.3%).",
    "Conversion improved from 34% to 44%.",
    "Sales campaign ROI improved from $8,731 to $13,188 (+51%).",
    "Outstanding service tasks were broadly steady at 7.",
]


def metric(
    *,
    direction: str = "up",
    percent_change: float = 12.0,
    previous_display: str = "18",
    current_display: str = "20",
    metric_type: str = "volume",
    significance_band: str = "meaningful",
) -> dict[str, object]:
    return {
        "usable_for_reporting": True,
        "is_reliable": True,
        "direction": direction,
        "percent_change": percent_change,
        "significance_band": significance_band,
        "previous_display": previous_display,
        "current_display": current_display,
        "metric_type": metric_type,
    }


DEDUP_METRICS = {
    "sales_total_deals": metric(percent_change=125.0, previous_display="8", current_display="18"),
    "sales_overall_conversion_rate": metric(
        percent_change=4.0,
        previous_display="25%",
        current_display="26%",
        metric_type="rate",
    ),
    "sales_email_capture_rate": metric(
        percent_change=2.8,
        previous_display="72%",
        current_display="74%",
        metric_type="rate",
    ),
    "sales_total_leads": metric(percent_change=18.0, previous_display="100", current_display="118"),
    "sales_campaign_roi": metric(percent_change=21.0, previous_display="$10,000", current_display="$12,100", metric_type="currency"),
    "service_campaign_roi": metric(percent_change=11.0, previous_display="$90,000", current_display="$99,900", metric_type="currency"),
    "sales_uncontacted_leads": metric(
        direction="down",
        percent_change=-16.0,
        previous_display="12",
        current_display="10",
        metric_type="count",
    ),
    "service_outstanding_tasks": metric(
        direction="down",
        percent_change=-10.0,
        previous_display="10",
        current_display="9",
        metric_type="count",
    ),
}


class EnsureFourPriorityInsightsTests(unittest.TestCase):
    def test_zero_candidates_top_up_to_four(self) -> None:
        out = core.ensure_four_priority_insights([])
        self.assertEqual(len(out), core.PRIORITY_INSIGHT_COUNT)
        self.assertTrue(all(item.strip() for item in out))
        self.assertTrue(all(item != core.NO_ADDITIONAL_INSIGHT_MARKER for item in out))

    def test_one_candidate_top_up_to_four(self) -> None:
        out = core.ensure_four_priority_insights(["Deal volume increased from 10 to 20 (+100%)."])
        self.assertEqual(len(out), 4)
        self.assertEqual(out[0], "Deal volume increased from 10 to 20 (+100%).")

    def test_two_and_three_candidates_top_up(self) -> None:
        for n in (2, 3):
            with self.subTest(n=n):
                out = core.ensure_four_priority_insights(FOUR[:n])
                self.assertEqual(len(out), 4)

    def test_five_candidates_trim_to_four(self) -> None:
        extras = FOUR + ["Lead volume eased from 100 to 90 (-10%)."]
        out = core.ensure_four_priority_insights(extras)
        self.assertEqual(len(out), 4)
        self.assertEqual(out, FOUR)

    def test_metric_fallback_used_when_needed(self) -> None:
        metrics = {
            "sales_total_deals": {
                "usable_for_reporting": True,
                "is_reliable": True,
                "direction": "up",
                "percent_change": 12.0,
                "significance_band": "meaningful",
                "current_display": "20",
                "previous_display": "18",
                "metric_type": "volume",
            }
        }
        out = core.ensure_four_priority_insights(["Only one insight."], metrics=metrics)
        self.assertEqual(len(out), 4)
        joined = " ".join(out).lower()
        self.assertIn("only one insight", joined)

    def test_selected_deal_volume_insight_prevents_deal_volume_fallback(self) -> None:
        out = core.ensure_four_priority_insights(
            [
                "Deal volume increased from 8 to 18 (+125%).",
                "Outstanding service tasks were broadly steady at 7.",
            ],
            metrics=DEDUP_METRICS,
        )
        self.assertEqual(len(out), 4)
        self.assertEqual(sum("deal volume" in item.lower() for item in out), 1)
        self.assertNotIn("deal volume up 125", " ".join(out).lower())
        self.assertTrue(all(core.NO_ADDITIONAL_INSIGHT_MARKER not in item for item in out))

    def test_selected_conversion_insight_prevents_conversion_fallback(self) -> None:
        out = core.ensure_four_priority_insights(
            ["Conversion improved from 25% to 26%."],
            metrics=DEDUP_METRICS,
        )
        self.assertEqual(len(out), 4)
        self.assertEqual(sum("conversion" in item.lower() for item in out), 1)
        self.assertNotIn("overall conversion up from 25% to 26%", " ".join(out).lower())

    def test_selected_email_capture_insight_prevents_email_capture_fallback(self) -> None:
        out = core.ensure_four_priority_insights(
            ["Email capture rate improved from 72% to 74%, strengthening follow-up reach."],
            metrics=DEDUP_METRICS,
        )
        self.assertEqual(len(out), 4)
        self.assertEqual(sum("email capture" in item.lower() for item in out), 1)
        self.assertNotIn("email capture rate up from 72% to 74%", " ".join(out).lower())

    def test_deal_volume_duplication_regression_and_sms(self) -> None:
        out = core.ensure_four_priority_insights(
            [
                "Service task carryover stayed steady and should remain visible in daily follow-up.",
                "Deal volume increased from 8 to 18 (+125%).",
            ],
            metrics=DEDUP_METRICS,
        )
        self.assertEqual(len(out), 4)
        self.assertEqual(sum("deal volume" in item.lower() for item in out), 1)
        self.assertTrue(all(item.strip() for item in out))
        self.assertTrue(all(core.NO_ADDITIONAL_INSIGHT_MARKER not in item for item in out))

        sms, segments, meta = core.build_sms_from_priority_insights(out, "Jun 2026")
        self.assertEqual(segments, ["priority_insights"])
        self.assertEqual(meta["sms_source"], "priority_insights")
        self.assertLessEqual(len(sms), core.SMS_CHAR_LIMIT)
        self.assertTrue(meta["sms_valid"])


class SmsFromPriorityInsightsTests(unittest.TestCase):
    def test_sms_uses_required_dealersocket_format(self) -> None:
        sms, segments, meta = core.build_sms_from_priority_insights(FOUR, "Jun 2026")
        self.assertTrue(sms.startswith("DealerSocket June Insights:"))
        self.assertEqual(segments, ["priority_insights"])
        self.assertEqual(meta["sms_source"], "priority_insights")
        self.assertLessEqual(len(sms), core.SMS_CHAR_LIMIT)
        self.assertTrue(meta["sms_valid"])

    def test_sms_uses_priority_insights_only(self) -> None:
        sms, segments, meta = core.build_sms_from_priority_insights(FOUR, "Jun 2026")
        self.assertEqual(segments, ["priority_insights"])
        self.assertEqual(meta["sms_source"], "priority_insights")
        self.assertNotIn("placeholder summary", sms.lower())
        self.assertLessEqual(len(sms), core.SMS_CHAR_LIMIT)
        self.assertTrue(meta["sms_valid"])

    def test_sms_has_no_generic_referral_language(self) -> None:
        sms, _, _ = core.build_sms_from_priority_insights(FOUR, "Jun 2026")
        self.assertFalse(core.validate_sms_no_generic_referral(sms))
        self.assertNotIn("see report", sms.lower())
        self.assertNotIn("see attached", sms.lower())
        self.assertNotIn("4 key insights", sms.lower())

    def test_sms_format_validation_passes(self) -> None:
        sms, _, _ = core.build_sms_from_priority_insights(FOUR, "Jun 2026")
        self.assertEqual(core.validate_sms_format(sms, "Jun 2026"), [])

    def test_sms_priority_traceability(self) -> None:
        sms, _, _ = core.build_sms_from_priority_insights(FOUR, "Jun 2026")
        self.assertEqual(core.validate_sms_priority_traceability(sms, FOUR), [])

    def test_sms_has_quantified_metrics(self) -> None:
        sms, _, meta = core.build_sms_from_priority_insights(FOUR, "Jun 2026")
        self.assertRegex(sms, r"\d")
        self.assertEqual(core.validate_sms_contains_quantified_metrics(sms, FOUR), [])
        self.assertTrue(meta["sms_valid"])

    def test_sms_numbers_traceable_to_priority_insights(self) -> None:
        sms, _, _ = core.build_sms_from_priority_insights(FOUR, "Jun 2026")
        self.assertEqual(core.validate_sms_number_traceability(sms, FOUR), [])

    def test_sms_no_invented_numbers(self) -> None:
        sms, _, _ = core.build_sms_from_priority_insights(FOUR, "Jun 2026")
        self.assertEqual(core.validate_sms_no_invented_numbers(sms, FOUR), [])

    def test_sms_not_qualitative_only_when_numbers_available(self) -> None:
        insights = [
            "Deal volume increased from 8 to 18 (+125%).",
            "Sales campaign ROI fell from $10,000 to $5,000 (-50%).",
            "Service campaign ROI improved from $2,000 to $3,000 (+50%).",
            "Lead volume increased from 100 to 118 (+18%).",
        ]
        sms, _, meta = core.build_sms_from_priority_insights(insights, "Jun 2026")
        self.assertEqual(core.validate_sms_qualitative_only(sms, insights), [])
        self.assertRegex(sms, r"\d")
        self.assertTrue(meta["sms_valid"])

    def test_june_quantified_sms(self) -> None:
        insights = [
            "Service task carryover remained relatively steady, indicating no major shift in service follow-up pressure.",
            "Deal volume increased from 8 to 18 (+125%).",
            "lead volume up 118.8% month on month.",
            "lead-to-test-drive conversion up from 3% to 30% month on month.",
        ]
        sms, _, meta = core.build_sms_from_priority_insights(insights, "Jun 2026")
        self.assertIn("125%", sms)
        self.assertIn("118.8%", sms)
        self.assertNotIn("see report", sms.lower())
        self.assertTrue(meta["sms_valid"])

    def test_sms_includes_month(self) -> None:
        sms, _, _ = core.build_sms_from_priority_insights(FOUR, "Jun 2026")
        self.assertIn("June", sms)

    def test_sms_under_limit_for_long_insights(self) -> None:
        long_insights = [
            "Deals and conversion stayed constructive, although CRM execution slipped, with overdue sales tasks up 82.9%.",
            "Service campaign ROI fell from $509,050 to $442,019 (-13.2%).",
            "Lead volume eased from 767 to 743 (-3.1%), but conversion improved from 26% to 28%.",
            "Email capture strengthened, with email capture rate up from 48% to 56% month on month.",
        ]
        sms, _, meta = core.build_sms_from_priority_insights(long_insights, "Jun 2026")
        self.assertLessEqual(len(sms), core.SMS_CHAR_LIMIT)
        self.assertTrue(meta["sms_valid"])

    def test_deal_volume_and_campaign_deals_keep_distinct_labels(self) -> None:
        insights = [
            "Deal volume increased from 10 to 20 (+100%).",
            "Sales campaign deals fell from 5 to 3 (-40%).",
            "Service campaign ROI improved from $1,000 to $1,500 (+50%).",
            "Response times improved from 10:00 to 5:00 (-50%).",
        ]
        sms, _, _ = core.build_sms_from_priority_insights(insights, "Jun 2026")
        self.assertIn("deal volume", sms.lower())
        self.assertIn("campaign deals", sms.lower())
        self.assertRegex(sms, r"\d")
        self.assertNotRegex(sms.lower(), r"deals up.*deals down|deals down.*deals up")

    def test_conversion_and_test_drive_conversion_keep_distinct_labels(self) -> None:
        insights = [
            "Conversion improved from 34% to 40%.",
            "lead-to-test-drive conversion down from 26% to 24% month on month.",
            "Deal volume increased from 10 to 12 (+20%).",
            "Email capture strengthened, with email capture rate up from 72% to 74% month on month.",
        ]
        sms, _, _ = core.build_sms_from_priority_insights(insights, "Jun 2026")
        self.assertIn("conversion up from 34% to 40%", sms)
        self.assertRegex(sms, r"\d")
        self.assertNotRegex(sms.lower(), r"conversion up.*conversion down|conversion down.*conversion up")

    def test_lead_volume_and_appointment_conversion_keep_distinct_labels(self) -> None:
        insights = [
            "Despite 37.3% lead volume growth, appointment conversion fell 9.1%, indicating qualification challenges.",
            "Sales campaign deals increased from 31 to 44 (+41.9%), with sales campaign ROI improved from $18,944 to $38,232 (+101.8%).",
            "Deal volume increased from 21 to 153 (+628.6%).",
            "Outstanding service tasks rose from 61 to 75 (+22.9%).",
        ]
        sms, _, _ = core.build_sms_from_priority_insights(insights, "Jun 2026")
        self.assertIn("leads rose", sms)
        self.assertIn("appointment conversion softened", sms)
        self.assertNotIn("leads down", sms.lower())

    def test_campaign_deals_ignores_roi_dollar_values_in_same_sentence(self) -> None:
        insights = [
            "Outstanding service tasks rose from 4 to 8 (+100%).",
            "Sales campaign deals increased from 33 to 35 (+6.1%), with sales campaign ROI improved from $24,247 to $32,538 (+34.2%).",
            "The team responded much faster this month, with response time down 53.5% to 18:23.",
            "Deal volume increased from 137 to 168 (+22.6%).",
        ]
        sms, _, _ = core.build_sms_from_priority_insights(insights, "Jun 2026")
        self.assertIn("campaign deals up from 33 to 35", sms)
        self.assertIn("sales ROI improved from $24,247 to $32,538", sms)
        self.assertNotIn("campaign deals up from $24,247", sms)

    def test_sales_roi_and_service_roi_remain_distinct(self) -> None:
        insights = [
            "Sales campaign ROI fell from $10,000 to $5,000 (-50%).",
            "Service campaign ROI improved from $2,000 to $3,000 (+50%).",
            "Deal volume increased from 10 to 12 (+20%).",
            "Email capture strengthened, with email capture rate up from 70% to 72% month on month.",
        ]
        sms, _, _ = core.build_sms_from_priority_insights(insights, "Jun 2026")
        self.assertIn("service ROI improved", sms)
        self.assertIn("sales ROI softened", sms)

    def test_metric_aware_polarity_for_inverse_and_capture_metrics(self) -> None:
        cases = [
            ("Response times improved from 10:00 to 5:00 (-50%).", "positive", "response times improved"),
            ("Uncontacted leads fell from 17 to 16 (-5.9%).", "positive", "uncontacted leads reduced"),
            ("Outstanding service tasks fell from 286 to 267 (-6.6%).", "positive", "service tasks improved"),
            ("Email capture strengthened, with email capture rate up from 48% to 56% month on month.", "positive", "email capture improved"),
            ("Sales campaign ROI fell from $10,000 to $5,000 (-50%).", "negative", "sales ROI softened"),
            ("Service campaign ROI improved from $2,000 to $3,000 (+50%).", "positive", "service ROI improved"),
        ]
        for text, polarity, theme in cases:
            with self.subTest(text=text):
                self.assertEqual(core._infer_insight_polarity(text), polarity)
                self.assertEqual(core._extract_sms_theme(text), theme)

    def test_follow_up_does_not_trigger_up_polarity(self) -> None:
        text = "Customer follow-up coverage remained visible this month."
        self.assertNotEqual(core._infer_insight_polarity(text), "positive")

    def test_semantic_validator_catches_unqualified_contradictions(self) -> None:
        self.assertTrue(core.validate_sms_semantics("June report: conversion up, but conversion down. See report."))
        self.assertTrue(core.validate_sms_semantics("June report: deals up, but deals down. See report."))
        self.assertTrue(core.validate_sms_semantics("June report: leads up, but leads down. See report."))
        self.assertFalse(core.validate_sms_semantics("June report: deal volume improved, but campaign deals softened. See report."))
        self.assertFalse(core.validate_sms_semantics("June report: conversion improved, but test-drive conversion softened. See report."))
        self.assertFalse(core.validate_sms_semantics("June report: leads rose, but appointment conversion softened. See report."))
        self.assertFalse(core.validate_sms_semantics("June report: conversion improved. See report."))

    def test_semantic_validator_catches_truncated_chunk_text(self) -> None:
        self.assertTrue(core.validate_sms_semantics("June report: Lead qualification performance…. See report."))

    def test_selected_rule_duplicate_metric_evidence_is_suppressed(self) -> None:
        rules = [
            core.RuleInsight(
                rule_id="combined_sales_service_up",
                message="Deal volume increased from 70 to 88 (+25.7%) and service campaign ROI improved from $34,442 to $59,276 (+72.1%).",
                category="combined",
                severity="high",
                priority_weight=90,
                polarity="positive",
                tone_band="improved",
                merge_key="combined_sales_service_up",
                supporting_metrics=["sales_total_deals", "service_campaign_roi"],
                metric_focus=["sales_total_deals", "service_campaign_roi"],
            ),
            core.RuleInsight(
                rule_id="service_roi_snapshot",
                message="Service campaign ROI improved from $34,442 to $59,276 (+72.1%).",
                category="service",
                severity="medium",
                priority_weight=80,
                polarity="positive",
                tone_band="improved",
                merge_key="service_roi_snapshot",
                supporting_metrics=["service_campaign_roi"],
                metric_focus=["service_campaign_roi"],
            ),
            core.RuleInsight(
                rule_id="service_task_snapshot",
                message="Outstanding service tasks rose from 107 to 378 (+253.3%).",
                category="service",
                severity="medium",
                priority_weight=70,
                polarity="negative",
                tone_band="softer",
                merge_key="service_task_snapshot",
                supporting_metrics=["service_outstanding_tasks"],
                metric_focus=["service_outstanding_tasks"],
            ),
        ]
        selected, _ = core.select_top_insights(rules, limit=4, metrics=DEDUP_METRICS)
        self.assertEqual(sum(rule.rule_id == "service_roi_snapshot" for rule in selected), 0)
        final = core.ensure_four_priority_insights([rule.message for rule in selected], metrics=DEDUP_METRICS, selected_rules=selected)
        self.assertEqual(len(final), core.PRIORITY_INSIGHT_COUNT)
        self.assertEqual(sum("service campaign roi improved" in item.lower() for item in final), 1)


class EmailInsightValidationTests(unittest.TestCase):
    def test_four_slots_required_in_final_ready_mode(self) -> None:
        with self.assertRaises(email_builder.InsightValidationError):
            email_builder.normalize_priority_insights_to_four(["only one"], final_ready=True)

    def test_placeholder_padding_rejected(self) -> None:
        with self.assertRaises(email_builder.InsightValidationError):
            email_builder.normalize_priority_insights_to_four(
                FOUR[:1] + [email_builder.NO_ADDITIONAL_INSIGHT_MARKER] * 3,
                final_ready=True,
            )

    def test_preview_mode_allows_padding(self) -> None:
        i1, i2, i3, i4 = email_builder.normalize_priority_insights_to_four(["A"], final_ready=False)
        self.assertEqual(i1, "A")
        self.assertEqual(i2, email_builder.INSIGHT_SLOT_PAD)

    def test_validate_four_insight_slots(self) -> None:
        errors = email_builder.validate_four_insight_slots(
            insight_1="A",
            insight_2="B",
            insight_3="",
            insight_4="D",
        )
        self.assertIn("insight_3 is missing", errors)

    def test_sms_length_validation(self) -> None:
        ok, err = email_builder.validate_sms_output("x" * 120)
        self.assertTrue(ok)
        self.assertEqual(err, "")
        ok2, err2 = email_builder.validate_sms_output("x" * 200)
        self.assertFalse(ok2)
        self.assertIn("exceeds", err2)


class EmailBlankSummaryTests(unittest.TestCase):
    def test_email_renders_with_blank_summary(self) -> None:
        assembled = email_assembler.assemble_dealership_email(
            dealership_name="Test Dealer",
            month_label="Jun 2026",
            greeting_name="Pat",
            priority_insights=FOUR,
            summary_resolved="",
            to_recipients="",
            cc_recipients="",
            pdf_attachments=[],
            report_reference={},
            assets_dir=None,
        )
        self.assertIn(FOUR[0], assembled.plain_text_body)
        self.assertIn(FOUR[3], assembled.plain_text_body)
        # The body names the month on its own ("...Reports for June"); the year
        # appears only in the subject line.
        self.assertIn("June", assembled.plain_text_body)
        self.assertEqual(assembled.subject, "Your Monthly DealerSocket Snapshot - June 2026")
        # A blank summary must not leave an orphaned "Summary" heading behind.
        self.assertNotIn("Summary", assembled.plain_text_body)


if __name__ == "__main__":
    unittest.main()
