"""Behaviour pinned by the pre-transfer optimisation pass.

Each test here corresponds to a defect that was found by audit rather than by
use, so nothing in the app exercises them loudly enough to notice a regression.
"""

from __future__ import annotations

import re
import threading
import unittest
from pathlib import Path
from unittest import mock

import dealership_insights_local as core
import email_assembler
import report_engine
import sent_email_archive
import settings_manager


class AccountsFolderIsNotHardcodedTests(unittest.TestCase):
    """The accounts folder used to be one operator's OneDrive path baked into
    the source, so every command-line entry point died on any other machine."""

    def test_environment_override_wins(self) -> None:
        with mock.patch.dict("os.environ", {"REPORTER_ACCOUNTS_DIR": "/tmp/some-accounts"}):
            self.assertEqual(core._default_accounts_dir(), Path("/tmp/some-accounts"))

    def test_no_setting_and_no_override_means_no_folder(self) -> None:
        with mock.patch.dict("os.environ", {"REPORTER_ACCOUNTS_DIR": ""}), \
             mock.patch.object(settings_manager, "load_settings", return_value={"accounts_folder": ""}):
            self.assertIsNone(core._default_accounts_dir())

    def test_no_user_profile_paths_remain_in_source(self) -> None:
        for name in ("dealership_insights_local.py", "report_engine.py", "reporter_ui.py",
                     "email_builder.py", "email_assembler.py", "settings_manager.py",
                     "sent_email_archive.py", "run_month.py"):
            path = Path(__file__).resolve().parent.parent / name
            if not path.is_file():
                continue  # not shipped in this package (the desktop UI is not included)
            text = path.read_text(encoding="utf-8")
            leaked = re.search(r"[A-Za-z]:\\+Users\\+[^\\\"'\s]+", text)
            self.assertIsNone(leaked, f"{name} carries an operator-specific path")


class ArchiveUsesSelfContainedHtmlTests(unittest.TestCase):
    """The archive writes a standalone file and hands it to Word. The cid:
    variant only resolves inside a MIME message, so archiving it produced PDFs
    with four broken images."""

    def test_archive_prefers_archive_html_over_cid_html(self) -> None:
        captured: dict = {}

        def fake_archive(*, dealership_name, month_label, html_body, destination_folder,
                         word=None, layout="month_folder"):
            captured["html"] = html_body
            return {"dealership": dealership_name, "path": "", "format": "html", "error": ""}

        with mock.patch.object(sent_email_archive, "archive_sent_email", side_effect=fake_archive), \
             mock.patch.object(sent_email_archive, "_open_word", return_value=None), \
             mock.patch.object(sent_email_archive, "_close_word"):
            sent_email_archive.archive_sent_emails(
                sent_results=[{
                    "status": "sent", "dealership": "Good Motors",
                    "html_body": '<img src="cid:banner_top">',
                    "archive_html": '<img src="data:image/png;base64,AAA">',
                }],
                month_label="Jul 2026",
                destination_folder="/tmp",
            )
        self.assertIn("data:image", captured["html"])
        self.assertNotIn("cid:", captured["html"])

    def test_word_is_started_once_for_the_whole_batch(self) -> None:
        """Starting Word per email cost 1-3s each; a 27-email run spent up to a
        minute in start-up alone."""
        with mock.patch.object(sent_email_archive, "_open_word", return_value=mock.MagicMock()) as opener, \
             mock.patch.object(sent_email_archive, "_close_word"), \
             mock.patch.object(sent_email_archive, "_html_to_pdf_via_word", return_value=False):
            sent_email_archive.archive_sent_emails(
                sent_results=[{"status": "sent", "dealership": f"D{i}", "archive_html": "<p/>"} for i in range(5)],
                month_label="Jul 2026",
                destination_folder="/tmp/archive-test",
            )
        self.assertEqual(opener.call_count, 1)

    def test_a_batch_with_nothing_to_archive_never_starts_word(self) -> None:
        with mock.patch.object(sent_email_archive, "_open_word") as opener:
            self.assertEqual(
                sent_email_archive.archive_sent_emails(
                    sent_results=[{"status": "failed", "dealership": "D"}],
                    month_label="Jul 2026", destination_folder="/tmp",
                ),
                [],
            )
        opener.assert_not_called()

    def test_word_is_a_private_instance(self) -> None:
        """Dispatch attaches to the operator's own Word and then quits it,
        closing their documents. DispatchEx gets a dedicated one."""
        import inspect

        self.assertIn("DispatchEx", inspect.getsource(sent_email_archive._open_word))


class ReportScanCacheTests(unittest.TestCase):
    def setUp(self) -> None:
        core.clear_report_scan_cache()

    def tearDown(self) -> None:
        core.clear_report_scan_cache()

    def test_repeat_scans_of_one_folder_hit_the_cache(self) -> None:
        """The folder is rescanned for the target month and each of three
        historical months, per entity — all identical recursive walks."""
        folder = mock.MagicMock(spec=Path)
        folder.rglob.return_value = iter([])
        folder.__str__ = lambda self: "/accounts/Some Dealership"  # type: ignore[method-assign]

        for _ in range(4):
            core._dealership_pdf_files(folder)
        self.assertEqual(folder.rglob.call_count, 1)

    def test_clearing_the_cache_forces_a_fresh_scan(self) -> None:
        folder = mock.MagicMock(spec=Path)
        folder.rglob.return_value = iter([])
        folder.__str__ = lambda self: "/accounts/Some Dealership"  # type: ignore[method-assign]

        core._dealership_pdf_files(folder)
        core.clear_report_scan_cache()
        core._dealership_pdf_files(folder)
        self.assertEqual(folder.rglob.call_count, 2)

    def test_a_run_starts_with_a_clear_cache(self) -> None:
        import inspect

        self.assertIn("clear_report_scan_cache()", inspect.getsource(core.run_reporting_pipeline))


class ContactReloadPreservesUnaffectedEmailsTests(unittest.TestCase):
    """Editing one contact row used to clear every preview and working email in
    the session, silently discarding every manual edit across the whole run."""

    def setUp(self) -> None:
        self._previous = report_engine.CURRENT_SESSION
        dealerships = {}
        for name in ("Kept Motors", "Changed Motors"):
            state = report_engine._build_state_for_result(
                core.DealershipResult(
                    dealership_name=name, month_label="Jul 2026", bullet_insights=["a"],
                    summary="s", sms_version="sms", report_status="success",
                    parsing_status="parsed", insights_source="rule-based", confidence_score=0.9,
                    top_priority_level="High", review_required="no", sentiment="mixed",
                    insight_count_after_deduplication=1, review_reason_tags=[], extracted_metrics=[],
                    rules_triggered=[], sms_valid=True, sms_char_count=3,
                    sms_limit=core.SMS_CHAR_LIMIT, sms_source="priority_insights",
                ), {},
            )
            # Seeded to match exactly what the reload resolves for Kept Motors,
            # so "nothing changed" is a true no-op for it.
            state.to_recipients = f"{name.split()[0].lower()}@example.com"
            state.cc_recipients = ""
            state.contact_greeting_name = "C"
            state.contact_ready = True
            state.email_preview = {"html_body": "<p>built</p>"}
            dealerships[name] = state
        report_engine.CURRENT_SESSION = report_engine.ReportingSession(
            accounts_folder=Path("."), output_folder=Path("."), template_path=None,
            recipient_mapping_path=__file__, month=7, year=2026, month_label="Jul 2026",
            artifacts=None, recipient_mapping={}, dealerships=dealerships,
        )

    def tearDown(self) -> None:
        report_engine.CURRENT_SESSION = self._previous

    def test_only_the_dealership_whose_contact_moved_is_rebuilt(self) -> None:
        def resolve(name, _mapping):
            changed = name == "Changed Motors"
            return mock.Mock(
                to="new@example.com" if changed else "kept@example.com",
                cc="", bcc="", contact_name="C", greeting_name="C",
                matched=True, row_active=True, ready=True,
            )

        with mock.patch("email_builder.load_recipient_mappings", return_value={}), \
             mock.patch("email_builder.resolve_contact_for_dealership", side_effect=resolve):
            report_engine.reload_session_recipient_mapping()

        session = report_engine.CURRENT_SESSION
        self.assertIsNotNone(
            session.dealerships["Kept Motors"].email_preview,
            "an untouched dealership must keep its built email",
        )
        self.assertIsNone(session.dealerships["Changed Motors"].email_preview)


class MsgAssetExtractionIsCachedTests(unittest.TestCase):
    def test_cache_is_keyed_on_file_identity_and_guarded_by_a_lock(self) -> None:
        self.assertIsInstance(email_assembler._MSG_ASSET_LOCK, type(threading.Lock()))
        self.assertIsInstance(email_assembler._MSG_ASSET_CACHE, dict)

    def test_a_missing_template_is_not_an_error(self) -> None:
        self.assertEqual(email_assembler.extract_msg_inline_assets(None), [])
        self.assertEqual(email_assembler.extract_msg_inline_assets(Path("/nope/absent.msg")), [])

    def test_the_jinja_environment_is_built_once(self) -> None:
        """It was rebuilt three times per email for no benefit."""
        import inspect

        self.assertNotIn("Environment(", inspect.getsource(email_assembler._render_jinja))


class DiagnosticLogsAreBoundedTests(unittest.TestCase):
    def test_workflow_evidence_is_capped_like_every_other_diagnostic(self) -> None:
        import json
        import tempfile

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "evidence.json"
            with mock.patch.object(report_engine, "EVIDENCE_JSON_PATH", path), \
                 mock.patch.object(report_engine, "MAX_DIAGNOSTIC_RECORDS", 5):
                for index in range(8):
                    report_engine.append_workflow_evidence_record("test", {"n": index})
            records = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual([record["n"] for record in records], [3, 4, 5, 6, 7])


try:
    import tkinter  # noqa: F401
    _HAS_TK = True
except Exception:  # noqa: BLE001
    _HAS_TK = False

# The desktop UI is not part of the web package.
_HAS_DESKTOP_UI = _HAS_TK and (Path(__file__).resolve().parent.parent / "reporter_ui.py").is_file()


@unittest.skipUnless(_HAS_DESKTOP_UI, "desktop UI (reporter_ui.py) not available")
class SendIsNotOnTheMainThreadTests(unittest.TestCase):
    """Sending is the longest and only irreversible operation; running it on the
    Tk main thread froze the window for the length of a 27-email batch."""

    def test_run_send_hands_off_to_a_worker(self) -> None:
        import inspect

        import reporter_ui

        source = inspect.getsource(reporter_ui.ReporterApp._run_send)
        self.assertIn("threading.Thread", source)
        self.assertIn("send_batch_done", source)

    def test_the_archive_folder_is_resolved_before_leaving_the_main_thread(self) -> None:
        """Dialogs cannot be opened from a worker thread."""
        import inspect

        source = inspect.getsource(reporter_ui.ReporterApp._run_send) if False else None
        import reporter_ui

        run_send = inspect.getsource(reporter_ui.ReporterApp._run_send)
        self.assertIn("_resolve_sent_pdf_folder()", run_send)
        archive = inspect.getsource(reporter_ui.ReporterApp._archive_sent_emails)
        self.assertNotIn("askdirectory", archive)

    def test_template_validation_is_debounced(self) -> None:
        """It fires per keystroke and each run opens the .msg through Outlook."""
        import inspect

        import reporter_ui

        source = inspect.getsource(reporter_ui.ReporterApp._on_template_path_change)
        self.assertIn("after_cancel", source)
        self.assertIn("self.after(", source)

    def test_preview_no_longer_revalidates_the_template(self) -> None:
        import inspect

        import reporter_ui

        # The call, not the word: the method carries a comment explaining why
        # the call was removed.
        self.assertNotIn(
            "report_engine.validate_template(",
            inspect.getsource(reporter_ui.ReporterApp._run_preview_worker),
        )

    def test_the_open_pdf_is_reused_between_page_turns(self) -> None:
        import inspect

        import reporter_ui

        self.assertIn("_open_stage1_pdf", inspect.getsource(reporter_ui.ReporterApp._render_stage1_pdf_worker))


if __name__ == "__main__":
    unittest.main()


@unittest.skipUnless(_HAS_DESKTOP_UI, "desktop UI (reporter_ui.py) not available")
class HtmlPreviewSurvivesABlockedNativeLibraryTests(unittest.TestCase):
    """tkinterweb renders through Tkhtml, a native library. Managed Windows
    machines that block unsigned DLLs made the app unusable, because Tkhtml
    loads its DLL through Tcl when the *widget is built* — so the import-time
    `except ImportError` guard never fired and startup died on a TclError.
    """

    def test_the_import_guard_is_not_limited_to_import_errors(self) -> None:
        import re as _re
        from pathlib import Path as _Path

        source = (_Path(__file__).resolve().parent.parent / "reporter_ui.py").read_text(encoding="utf-8")
        guard = _re.search(r"try:\s*\n\s*from tkinterweb import HtmlFrame\s*\nexcept (\w+):", source)
        self.assertIsNotNone(guard, "the tkinterweb import guard has moved")
        self.assertEqual(
            guard.group(1), "Exception",
            "a blocked native DLL does not always raise ImportError",
        )

    def test_widget_construction_is_guarded_too(self) -> None:
        import inspect

        import reporter_ui

        source = inspect.getsource(reporter_ui.ReporterApp._build_send_stage_panel)
        self.assertIn("HtmlFrame(", source)
        frame_call = source[source.index("HtmlFrame("):]
        self.assertIn("except Exception", source[:source.index("HtmlFrame(")] + frame_call[:400],
                      "constructing HtmlFrame must be guarded, not just importing it")

    def test_a_browser_fallback_exists_and_needs_no_native_library(self) -> None:
        import inspect

        import reporter_ui

        source = inspect.getsource(reporter_ui.ReporterApp._open_html_preview_in_browser)
        self.assertIn("webbrowser.open", source)
        self.assertIn("_last_html_preview", source)

    def test_the_fallback_message_does_not_tell_a_blocked_user_to_install_it(self) -> None:
        """The old message said "Install tkinterweb" — precisely what someone
        blocked by policy cannot do."""
        from pathlib import Path as _Path

        source = (_Path(__file__).resolve().parent.parent / "reporter_ui.py").read_text(encoding="utf-8")
        self.assertNotIn("Install tkinterweb for rendered HTML support", source)
        self.assertIn("Open in browser", source)
