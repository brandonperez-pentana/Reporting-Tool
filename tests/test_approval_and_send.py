"""Approval gating and batch send behaviour.

The send path is the only irreversible thing this app does, so the rules it
depends on are pinned here:
  - a report with no source data can never be approved, and so can never reach
    the send queue;
  - one failing email never aborts the rest of the batch;
  - a test send is redirected and never touches a dealership address.
"""

from __future__ import annotations

import unittest
from pathlib import Path
from unittest import mock

import dealership_insights_local as core
import report_engine


def _result(name: str, *, blocked: bool = False) -> core.DealershipResult:
    return core.DealershipResult(
        dealership_name=name,
        month_label="Jul 2026",
        bullet_insights=["a", "b", "c", "d"],
        summary="s",
        sms_version="sms",
        report_status="success",
        parsing_status="fallback" if blocked else "parsed",
        insights_source="placeholder" if blocked else "rule-based",
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


class ApprovalAndSendTests(unittest.TestCase):
    def setUp(self) -> None:
        self._previous = report_engine.CURRENT_SESSION
        dealerships = {}
        for name, blocked in (("Good Motors", False), ("Other Motors", False), ("No Data Motors", True)):
            state = report_engine._build_state_for_result(_result(name, blocked=blocked), {})
            state.to_recipients = f"{name.split()[0].lower()}@example.com"
            state.contact_ready = True
            dealerships[name] = state
        report_engine.CURRENT_SESSION = report_engine.ReportingSession(
            accounts_folder=Path("."), output_folder=Path("."), template_path=None,
            recipient_mapping_path=None, month=7, year=2026, month_label="Jul 2026",
            artifacts=None, recipient_mapping={}, dealerships=dealerships,
        )

    def tearDown(self) -> None:
        report_engine.CURRENT_SESSION = self._previous

    def test_report_without_source_data_cannot_be_approved(self) -> None:
        allowed, reason = report_engine.can_approve("No Data Motors")
        self.assertFalse(allowed)
        self.assertIn("no source data", reason.lower())

    def test_report_with_data_can_be_approved(self) -> None:
        self.assertTrue(report_engine.can_approve("Good Motors")[0])

    def test_send_queue_contains_only_approved(self) -> None:
        self.assertEqual(report_engine.get_send_queue(), [])
        report_engine.mark_approved("Good Motors")
        queue = report_engine.get_send_queue()
        self.assertEqual([row["dealership"] for row in queue], ["Good Motors"])

    def test_ready_to_send_defaults_off(self) -> None:
        """Ticking must be a deliberate act, never a default."""
        report_engine.mark_approved("Good Motors")
        self.assertFalse(report_engine.get_send_queue()[0]["ready_to_send"])
        report_engine.set_ready_to_send("Good Motors", True)
        self.assertTrue(report_engine.get_send_queue()[0]["ready_to_send"])

    def test_one_failure_does_not_abort_the_batch(self) -> None:
        report_engine.mark_approved("Good Motors")
        report_engine.mark_approved("Other Motors")

        def flaky(name, *args, **kwargs):
            if name == "Good Motors":
                raise RuntimeError("Outlook refused")
            return {"to": "x@y.com", "subject": "s", "html_body": "<p/>"}

        with mock.patch.object(report_engine, "create_outlook_draft", side_effect=flaky):
            results = report_engine.send_emails(["Good Motors", "Other Motors"], send=True)

        self.assertEqual(len(results), 2)
        by_name = {r["dealership"]: r for r in results}
        self.assertEqual(by_name["Good Motors"]["status"], "failed")
        self.assertIn("Outlook refused", by_name["Good Motors"]["error"])
        self.assertEqual(by_name["Other Motors"]["status"], "sent")

    def test_unknown_dealership_is_reported_not_raised(self) -> None:
        results = report_engine.send_emails(["Ghost Motors"], send=True)
        self.assertEqual(results[0]["status"], "failed")
        self.assertIn("not part of the current reporting run", results[0]["error"])

    def test_blocked_dealership_cannot_be_sent(self) -> None:
        with self.assertRaises(RuntimeError):
            report_engine.create_outlook_draft("No Data Motors", send=True)

    def test_test_send_redirects_away_from_dealership(self) -> None:
        report_engine.mark_approved("Good Motors")
        captured: dict = {}

        def capture(**kwargs):
            captured.update(kwargs)
            return {"to": kwargs["to_recipients"], "subject": kwargs["subject"], "saved": True, "sent": True}

        with mock.patch.object(report_engine, "build_email_preview"), \
             mock.patch("email_builder.create_outlook_draft", side_effect=capture):
            state = report_engine.CURRENT_SESSION.dealerships["Good Motors"]
            state.working_email = report_engine.WorkingEmailState(
                to="dealer@example.com", cc="cc@example.com", subject="Snapshot",
                body="b", body_text="b", html_body="<p>b</p>", preview_text="b", attachments=[],
                template_type="html", template_path="", template_encoding="", placeholders_found=[],
                replaced_placeholders=[], missing_placeholders=[], contact={}, report_reference={},
                email_context={},
            )
            report_engine.create_outlook_draft(
                "Good Motors", send=True, override_to="me@company.com"
            )

        self.assertEqual(captured["to_recipients"], "me@company.com")
        self.assertEqual(captured["cc_recipients"], "", "CC must be cleared so no dealership is copied")
        self.assertIn("TEST", captured["subject"])
        self.assertFalse(
            report_engine.CURRENT_SESSION.dealerships["Good Motors"].email_sent,
            "a test send must not mark the real email as sent",
        )


try:  # reporter_ui needs tkinter, and is not shipped with the browser-only package.
    import tkinter  # noqa: F401
    import reporter_ui  # noqa: F401
    _HAS_TK = True
except Exception:  # noqa: BLE001
    _HAS_TK = False


@unittest.skipUnless(_HAS_TK, "tkinter unavailable")
class SendQueueSelectionGuardTests(unittest.TestCase):
    """The send queue rebuilds itself whenever it refreshes, and that rebuild
    re-fires <<TreeviewSelect>>. Without a suppression guard a single click
    looped select -> rebuild -> select and froze the window, so the guard and
    the repeat-selection check are pinned here."""

    def test_populate_send_queue_suppresses_selection_events(self) -> None:
        import inspect

        import reporter_ui

        source = inspect.getsource(reporter_ui.ReporterApp._populate_send_queue)
        self.assertIn("_suppress_tree_select = True", source)
        self.assertIn("_suppress_tree_select = False", source)

    def test_select_handler_ignores_suppressed_and_repeat_selections(self) -> None:
        import inspect

        import reporter_ui

        source = inspect.getsource(reporter_ui.ReporterApp._on_send_tree_select)
        self.assertIn("_suppress_tree_select", source)
        self.assertIn("_current_dealership", source)


if __name__ == "__main__":
    unittest.main()


class WorkingEmailHtmlPreservationTests(unittest.TestCase):
    """Selecting a different dealership flushes the current body back through
    update_working_email. When that flush regenerated HTML unconditionally it
    replaced the branded email (banners, button, inline images) with a plain
    conversion — so the preview lost its layout as soon as a second row was
    clicked. An unchanged body must be a no-op."""

    def setUp(self) -> None:
        self._previous = report_engine.CURRENT_SESSION
        state = report_engine._build_state_for_result(_result("Good Motors"), {})
        state.to_recipients = "good@example.com"
        state.contact_ready = True
        state.working_email = report_engine.WorkingEmailState(
            to="good@example.com", cc="", subject="Snapshot",
            body="Hi Sue,\n\nBody text.", body_text="Hi Sue,\n\nBody text.",
            html_body="<html><body><table>BRANDED</table></body></html>",
            preview_text="Hi Sue,\n\nBody text.", attachments=[],
            template_type="html", template_path="", template_encoding="",
            placeholders_found=[], replaced_placeholders=[], missing_placeholders=[],
            contact={}, report_reference={}, email_context={},
        )
        state.email_preview = {
            "html_body": "<html><body><table>BRANDED</table></body></html>",
            "html_body_outlook": "<html><body><table>CID</table></body></html>",
            "assembler_inline_assets": [{"cid": "banner_top", "path": "x.png"}],
        }
        report_engine.CURRENT_SESSION = report_engine.ReportingSession(
            accounts_folder=Path("."), output_folder=Path("."), template_path=None,
            recipient_mapping_path=None, month=7, year=2026, month_label="Jul 2026",
            artifacts=None, recipient_mapping={}, dealerships={"Good Motors": state},
        )

    def tearDown(self) -> None:
        report_engine.CURRENT_SESSION = self._previous

    def _state(self):
        return report_engine.CURRENT_SESSION.dealerships["Good Motors"]

    def test_unchanged_body_preserves_branded_html(self) -> None:
        report_engine.update_working_email("Good Motors", body="Hi Sue,\n\nBody text.")
        state = self._state()
        self.assertIn("BRANDED", state.working_email.html_body)
        self.assertIn("CID", state.email_preview["html_body_outlook"])
        self.assertEqual(len(state.email_preview["assembler_inline_assets"]), 1)

    def test_line_ending_differences_alone_are_not_an_edit(self) -> None:
        report_engine.update_working_email("Good Motors", body="Hi Sue,\r\n\r\nBody text.")
        self.assertIn("BRANDED", self._state().working_email.html_body)

    def test_editing_other_fields_preserves_branded_html(self) -> None:
        report_engine.update_working_email("Good Motors", subject="New subject", cc_recipients="cc@x.com")
        self.assertIn("BRANDED", self._state().working_email.html_body)

    def test_a_real_edit_still_takes_effect(self) -> None:
        report_engine.update_working_email("Good Motors", body="Hi Sue,\n\nCompletely new text.")
        state = self._state()
        self.assertIn("Completely new text.", state.working_email.body_text)
        self.assertNotIn("BRANDED", state.working_email.html_body)


class EmptyFieldFlushGuardTests(WorkingEmailHtmlPreservationTests):
    """A dealership shown on screen before its email finished building leaves
    the subject and body boxes empty. Clicking another row flushed those empty
    boxes back over the freshly built email, so the first dealership in a run
    ended up with a blank preview and no subject. Empty must never overwrite
    content."""

    def test_empty_subject_does_not_wipe_a_built_subject(self) -> None:
        report_engine.update_working_email("Good Motors", subject="")
        self.assertEqual(self._state().working_email.subject, "Snapshot")

    def test_empty_body_does_not_wipe_branded_html(self) -> None:
        report_engine.update_working_email("Good Motors", body="")
        state = self._state()
        self.assertIn("BRANDED", state.working_email.html_body)
        self.assertIn("CID", state.email_preview["html_body_outlook"])

    def test_whitespace_only_flush_is_also_ignored(self) -> None:
        report_engine.update_working_email("Good Motors", subject="   ", body="\n\n")
        state = self._state()
        self.assertEqual(state.working_email.subject, "Snapshot")
        self.assertIn("BRANDED", state.working_email.html_body)


@unittest.skipUnless(_HAS_TK, "tkinter unavailable")
class ApprovalRefreshesDisplayedDealershipTests(unittest.TestCase):
    def test_built_email_reloads_the_pane_when_it_is_on_screen(self) -> None:
        import inspect

        import reporter_ui

        source = inspect.getsource(reporter_ui.ReporterApp._handle_approval_email_built)
        self.assertIn("_current_dealership", source)
        self.assertIn("_update_email_preview_fields", source)
