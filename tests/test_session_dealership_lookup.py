"""Dealership names that change between reporting runs.

A folder containing several entities (e.g. "Demo Site 1") is reported per-entity
("Demo Site 1 North", "Demo Site 1 South") when its source PDFs are found, but
collapses to the folder name when they are not. So a name held from a previous
run can vanish, and every lookup must fail with an explanation rather than a
bare KeyError that surfaces to the operator as just 'Demo Site 1'.
"""

from __future__ import annotations

import unittest
from pathlib import Path

import dealership_insights_local as core
import report_engine


def _result(name: str) -> core.DealershipResult:
    return core.DealershipResult(
        dealership_name=name,
        month_label="Jul 2026",
        bullet_insights=["a", "b", "c", "d"],
        summary="s",
        sms_version="sms",
        report_status="success",
        parsing_status="parsed",
        insights_source="rule-based",
        confidence_score=0.9,
        top_priority_level="High",
        review_required="no",
        sentiment="mixed",
        insight_count_after_deduplication=4,
        review_reason_tags=[],
        extracted_metrics=[],
        rules_triggered=[],
        sms_valid=True,
        sms_char_count=3,
        sms_limit=core.SMS_CHAR_LIMIT,
        sms_source="priority_insights",
    )


def _session(names: list[str]) -> report_engine.ReportingSession:
    return report_engine.ReportingSession(
        accounts_folder=Path("."),
        output_folder=Path("."),
        template_path=None,
        recipient_mapping_path=None,
        month=7,
        year=2026,
        month_label="Jul 2026",
        artifacts=None,
        recipient_mapping={},
        dealerships={n: report_engine._build_state_for_result(_result(n), {}) for n in names},
    )


class SessionDealershipLookupTests(unittest.TestCase):
    def setUp(self) -> None:
        self._previous = report_engine.CURRENT_SESSION
        report_engine.CURRENT_SESSION = _session(["Demo Site 1 North", "Demo Site 1 South"])

    def tearDown(self) -> None:
        report_engine.CURRENT_SESSION = self._previous

    def test_has_dealership(self) -> None:
        self.assertTrue(report_engine.has_dealership("Demo Site 1 North"))
        self.assertFalse(report_engine.has_dealership("Demo Site 1"))

    def test_missing_dealership_raises_explained_error(self) -> None:
        with self.assertRaises(report_engine.DealershipNotInSessionError) as ctx:
            report_engine.get_dealership_report_text("Demo Site 1")
        message = str(ctx.exception)
        self.assertIn("Demo Site 1", message)
        self.assertIn("not part of the current reporting run", message)

    def test_error_message_is_not_a_bare_key_repr(self) -> None:
        """KeyError's default __str__ would render only \"'Demo Site 1'\"."""
        message = str(report_engine.DealershipNotInSessionError("Demo Site 1"))
        self.assertNotEqual(message, "'Demo Site 1'")
        self.assertGreater(len(message), len("'Demo Site 1'"))

    def test_build_preview_for_missing_dealership_is_explained(self) -> None:
        with self.assertRaises(report_engine.DealershipNotInSessionError):
            report_engine.build_email_preview("Demo Site 1")

    def test_update_text_for_missing_dealership_is_explained(self) -> None:
        with self.assertRaises(report_engine.DealershipNotInSessionError):
            report_engine.update_dealership_text("Demo Site 1", summary="x")

    def test_still_raises_for_no_session(self) -> None:
        report_engine.CURRENT_SESSION = None
        self.assertFalse(report_engine.has_dealership("Demo Site 1 North"))
        with self.assertRaises(report_engine.SessionNotInitializedError):
            report_engine.get_dealership_report_text("Demo Site 1 North")


if __name__ == "__main__":
    unittest.main()
