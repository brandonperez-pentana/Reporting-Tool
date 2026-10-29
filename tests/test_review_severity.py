"""Review severity classification.

"blocked" must mean exactly one thing: there is no usable source data, so
nothing can be sent until the PDFs are supplied and the month is regenerated.
It is deliberately the same condition that hard-blocks email preview and
Outlook draft creation.

Data-quality tags must NOT block. They were tried as blocking conditions and
fired on 23 of 27 accounts in a real month — every one of which had parsed real
data — which made "blocked" select nearly the whole queue and useless as a
filter. These tests pin that separation so it cannot drift back.
"""

from __future__ import annotations

import unittest

import dealership_insights_local as core

QUALITY_TAGS = [
    "unreliable_metrics",
    "conflicting_signals",
    "low_confidence",
    "anomaly_detected",
    "weak_conversion",
    "slow_response_time",
    "persistent_watchpoint",
]


class ReviewSeverityTests(unittest.TestCase):
    def test_missing_source_pdfs_is_blocked(self) -> None:
        self.assertEqual(
            core.classify_review_severity(["missing_reports"], "fallback", "placeholder"),
            "blocked",
        )

    def test_placeholder_insights_is_blocked(self) -> None:
        self.assertEqual(core.classify_review_severity([], "parsed", "placeholder"), "blocked")
        self.assertEqual(core.classify_review_severity([], "fallback", "rule-based"), "blocked")

    def test_quality_tags_on_real_data_are_flagged_not_blocked(self) -> None:
        for tag in QUALITY_TAGS:
            with self.subTest(tag=tag):
                self.assertEqual(
                    core.classify_review_severity([tag], "parsed", "rule-based"),
                    "flagged",
                )

    def test_many_quality_tags_together_still_only_flagged(self) -> None:
        self.assertEqual(
            core.classify_review_severity(QUALITY_TAGS, "parsed", "rule-based"),
            "flagged",
        )

    def test_partial_parse_is_flagged_not_blocked(self) -> None:
        """Partial data is real data; the app does not refuse to send it."""
        self.assertEqual(
            core.classify_review_severity(["unreliable_metrics"], "partial", "rule-based"),
            "flagged",
        )

    def test_clean_run_needs_no_review(self) -> None:
        self.assertEqual(core.classify_review_severity([], "parsed", "rule-based"), "none")
        self.assertEqual(core.classify_review_severity(None, "parsed", "rule-based"), "none")

    def test_blocked_matches_the_send_blocking_rule(self) -> None:
        """Severity must never disagree with what the app refuses to draft/send.

        report_engine blocks preview/draft on:
            parsing_status == "fallback" or insights_source == "placeholder"
        """
        cases = [
            ("fallback", "placeholder"),
            ("fallback", "rule-based"),
            ("parsed", "placeholder"),
            ("partial", "placeholder"),
            ("parsed", "rule-based"),
            ("partial", "rule-based"),
        ]
        for parsing_status, insights_source in cases:
            with self.subTest(parsing_status=parsing_status, insights_source=insights_source):
                send_blocked = parsing_status == "fallback" or insights_source == "placeholder"
                severity = core.classify_review_severity(
                    ["unreliable_metrics"], parsing_status, insights_source
                )
                self.assertEqual(severity == "blocked", send_blocked)


if __name__ == "__main__":
    unittest.main()
