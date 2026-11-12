"""The browser front end cannot send email, and PDF copies land beside the reports.

Sending was removed deliberately: every action this app offers stops at something a
person then acts on — a draft in Outlook, an .eml on disk, a PDF copy — so no click in
a browser tab can put mail in front of a dealership. These tests pin that as a property
rather than as the absence of a button, which a later change could quietly restore.
"""

from __future__ import annotations

import re
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import report_engine
import sent_email_archive
from webapp import api


class TheAppCannotSendTests(unittest.TestCase):
    def test_the_drafts_route_refuses_an_explicit_send(self) -> None:
        with self.assertRaises(api.ApiError) as caught:
            api.start_drafts({"dealerships": ["Demo Site 1"], "send": True})
        self.assertIn("cannot send", caught.exception.message.lower())

    def test_the_drafts_route_refuses_a_test_address(self) -> None:
        """A test send is still the app sending mail, so it goes with the rest."""
        with self.assertRaises(api.ApiError) as caught:
            api.start_drafts({"dealerships": ["Demo Site 1"], "test_address": "me@example.com"})
        self.assertIn("cannot send", caught.exception.message.lower())

    def test_the_refusal_says_what_can_be_done_instead(self) -> None:
        with self.assertRaises(api.ApiError) as caught:
            api.start_drafts({"dealerships": ["Demo Site 1"], "send": True})
        message = caught.exception.message.lower()
        for offered in ("draft", ".eml", "pdf"):
            self.assertIn(offered, message)

    def test_no_code_in_the_web_layer_asks_the_engine_to_send(self) -> None:
        """report_engine.send_emails still supports send=True for the desktop app.

        The guarantee is that nothing under webapp/ ever passes it, which is stronger
        than checking one call site: a new route could otherwise reintroduce sending
        without touching the one this test knew about.
        """
        offenders = []
        for path in sorted((ROOT / "webapp").rglob("*.py")):
            # Comments explain why sending is absent and would otherwise trip this.
            code = "\n".join(
                line for line in path.read_text(encoding="utf-8").splitlines()
                if not line.lstrip().startswith("#")
            )
            if re.search(r"send\s*=\s*True", code):
                offenders.append(path.name)
        self.assertEqual(offenders, [], "the web layer must never ask the engine to send")

    def test_the_page_offers_no_send_control(self) -> None:
        markup = (ROOT / "webapp" / "static" / "index.html").read_text(encoding="utf-8")
        script = (ROOT / "webapp" / "static" / "app.js").read_text(encoding="utf-8")
        self.assertNotIn("test_send_address", api.SETTING_KEYS)
        for gone in ("send-ticked", "send-test", "test_send_address"):
            self.assertNotIn(gone, markup)
            self.assertNotIn(gone, script)


class ActionableSelectionTests(unittest.TestCase):
    def _require(self, row=None, *, in_session=True):
        with (
            mock.patch.object(report_engine, "has_active_session", return_value=True),
            mock.patch.object(report_engine, "get_send_queue", return_value=[row] if row else []),
            mock.patch.object(report_engine, "has_dealership", return_value=in_session),
        ):
            api._require_actionable_dealerships(["Demo Site 1"])

    def _row(self, **overrides):
        row = {
            "dealership": "Demo Site 1",
            "ready_to_send": True,
            "email_ready": True,
            "sent": False,
        }
        row.update(overrides)
        return row

    def test_only_approved_checked_unsent_rows_with_built_email_are_actionable(self) -> None:
        self._require(self._row())

    def test_unapproved_dealership_is_rejected(self) -> None:
        with self.assertRaisesRegex(api.ApiError, "must be approved"):
            self._require()

    def test_unchecked_dealership_is_rejected(self) -> None:
        with self.assertRaisesRegex(api.ApiError, "must be checked"):
            self._require(self._row(ready_to_send=False))

    def test_missing_email_preview_is_rejected(self) -> None:
        with self.assertRaisesRegex(api.ApiError, "Build the email"):
            self._require(self._row(email_ready=False))

    def test_already_sent_dealership_is_rejected(self) -> None:
        with self.assertRaisesRegex(api.ApiError, "already been sent"):
            self._require(self._row(sent=True))

    def test_unknown_dealership_is_rejected(self) -> None:
        with self.assertRaisesRegex(api.ApiError, "not part of the current"):
            self._require(in_session=False)

    def test_draft_pdf_and_eml_routes_all_apply_the_selection_guard(self) -> None:
        with (
            mock.patch.object(report_engine, "has_active_session", return_value=True),
            mock.patch.object(report_engine, "get_send_queue", return_value=[]),
            mock.patch.object(report_engine, "has_dealership", return_value=True),
            mock.patch.object(api.settings_manager, "server_mode", return_value=False),
            mock.patch.object(api.settings_manager, "load_settings", return_value={"eml_export_folder": "out"}),
        ):
            for action in (api.start_drafts, api.start_pdf_copies, api.start_export):
                with self.subTest(action=action.__name__), self.assertRaisesRegex(api.ApiError, "must be approved"):
                    action({"dealerships": ["Demo Site 1"]})


class PdfCopiesLandBesideTheReportsTests(unittest.TestCase):
    def test_the_source_folder_is_taken_from_the_report_pdf(self) -> None:
        """Not accounts_folder plus the dealership name: the folder is often spelled
        differently, and every operator arranges their accounts folder their own way."""
        folder = Path(tempfile.mkdtemp(prefix="acct-")) / "Demo Site 1 (QLD)"
        folder.mkdir(parents=True)
        sales = folder / "sales.pdf"
        sales.write_bytes(b"%PDF-1.4\n")

        state = mock.Mock()
        state.result.sales_pdf = sales
        state.result.service_pdf = None
        session = mock.Mock(dealerships={"Demo Site 1": state})
        with mock.patch.object(report_engine, "_ensure_session", return_value=session):
            self.assertEqual(report_engine.source_folder_for("Demo Site 1"), folder)

    def test_no_source_pdf_means_no_guessed_folder(self) -> None:
        state = mock.Mock()
        state.result.sales_pdf = None
        state.result.service_pdf = None
        session = mock.Mock(dealerships={"Ghost Motors": state})
        with mock.patch.object(report_engine, "_ensure_session", return_value=session):
            self.assertIsNone(report_engine.source_folder_for("Ghost Motors"))

    def test_beside_source_writes_no_subfolder_and_dates_the_filename(self) -> None:
        """Without a month folder, the month has to be in the name or January's copy
        overwrites February's."""
        folder = Path(tempfile.mkdtemp(prefix="beside-"))
        with mock.patch.object(sent_email_archive, "_html_to_pdf_via_word", return_value=False):
            result = sent_email_archive.archive_sent_email(
                dealership_name="Demo Site 1",
                month_label="Aug 2026",
                html_body="<p>hi</p>",
                destination_folder=folder,
                layout="beside_source",
            )
        written = Path(result["path"])
        self.assertEqual(written.parent, folder, "no subfolder may be created")
        self.assertIn("Aug 2026", written.name)
        self.assertIn("Demo Site 1", written.name)

    def test_the_month_folder_layout_is_unchanged_for_the_desktop_app(self) -> None:
        folder = Path(tempfile.mkdtemp(prefix="month-"))
        with mock.patch.object(sent_email_archive, "_html_to_pdf_via_word", return_value=False):
            result = sent_email_archive.archive_sent_email(
                dealership_name="Demo Site 1",
                month_label="Aug 2026",
                html_body="<p>hi</p>",
                destination_folder=folder,
            )
        written = Path(result["path"])
        self.assertEqual(written.parent.name, "Aug 2026")
        self.assertEqual(written.stem, "Demo Site 1")

    def test_each_row_may_name_its_own_destination(self) -> None:
        """Every dealership has a different answer for 'beside its own reports'."""
        one = Path(tempfile.mkdtemp(prefix="one-"))
        two = Path(tempfile.mkdtemp(prefix="two-"))
        with mock.patch.object(sent_email_archive, "_html_to_pdf_via_word", return_value=False), \
             mock.patch.object(sent_email_archive, "_open_word", return_value=None), \
             mock.patch.object(sent_email_archive, "_close_word"):
            rows = sent_email_archive.archive_sent_emails(
                sent_results=[
                    {"status": "ok", "dealership": "A", "archive_html": "<p>a</p>", "destination_folder": str(one)},
                    {"status": "ok", "dealership": "B", "archive_html": "<p>b</p>", "destination_folder": str(two)},
                ],
                month_label="Aug 2026",
                layout="beside_source",
            )
        self.assertEqual(Path(rows[0]["path"]).parent, one)
        self.assertEqual(Path(rows[1]["path"]).parent, two)


if __name__ == "__main__":
    unittest.main()
