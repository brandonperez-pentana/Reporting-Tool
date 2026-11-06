"""Contact edits must survive reopening, failed writes and concurrent editors."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest import mock

import email_builder
import report_engine
import settings_manager
from webapp import api


class ContactsPersistenceTests(unittest.TestCase):
    def setUp(self) -> None:
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.root = Path(folder.name)
        self.path = self.root / "personal" / "contacts.csv"
        self.settings = self.root / "user_settings.json"
        self.row = {"dealership": "Demo Site 1", "contact_name": "Sam", "email": "sam@example.com"}
        # All settings and contact writes are isolated from the operator's files.
        for patcher in (
            mock.patch.object(settings_manager, "get_settings_path", return_value=self.settings),
            mock.patch.object(settings_manager, "_multi_user", return_value=False),
            mock.patch.object(settings_manager, "server_mode", return_value=False),
            mock.patch.object(api, "dealership_names", return_value=["Demo Site 1"]),
            mock.patch.object(report_engine, "has_active_session", return_value=False),
            mock.patch.object(report_engine, "append_workflow_evidence_record"),
        ):
            patcher.start()
            self.addCleanup(patcher.stop)

    def save(self, **extra):
        return api.save_contacts({"path": str(self.path), "rows": [self.row], **extra})

    def test_selected_file_is_remembered_and_reopened_from_disk(self) -> None:
        result = self.save()
        self.assertEqual(settings_manager.load_settings()["recipient_mapping_path"], str(self.path.resolve()))
        self.assertEqual(api.contacts()["rows"][0]["email"], self.row["email"])
        self.row["email"] = "updated@example.com"
        self.save(revision=result["contacts"]["revision"])
        self.assertEqual(api.contacts()["rows"][0]["email"], "updated@example.com")
        self.assertEqual(email_builder.load_contact_rows(self.path)[0]["email"], "updated@example.com")

    def test_stale_editor_does_not_overwrite_newer_contacts(self) -> None:
        revision = self.save()["contacts"]["revision"]
        self.row["email"] = "new@example.com"
        self.save(revision=revision)
        self.row["email"] = "stale@example.com"
        with self.assertRaises(api.ApiError) as caught:
            self.save(revision=revision)
        self.assertEqual(caught.exception.status, 409)
        self.assertEqual(api.contacts()["rows"][0]["email"], "new@example.com")

    def test_failed_replace_preserves_original_and_removes_temporary_file(self) -> None:
        self.save()
        original = self.path.read_bytes()
        self.row["email"] = "unsaved@example.com"
        with mock.patch.object(email_builder.os, "replace", side_effect=PermissionError("File is locked")):
            with self.assertRaises(api.ApiError) as caught:
                self.save()
        self.assertIn("not saved", str(caught.exception))
        self.assertEqual(self.path.read_bytes(), original)
        self.assertEqual(list(self.path.parent.glob("*.tmp")), [])

    def test_invalid_row_does_not_modify_file(self) -> None:
        self.save()
        original = self.path.read_bytes()
        with self.assertRaises(api.ApiError):
            api.save_contacts({"path": str(self.path), "rows": ["not a contact"]})
        self.assertEqual(self.path.read_bytes(), original)

    def test_missing_or_malformed_rows_never_clear_contacts(self) -> None:
        self.save()
        original = self.path.read_bytes()
        for rows in (None, {}, "", "bad rows"):
            with self.subTest(rows=rows), self.assertRaises(api.ApiError):
                api.save_contacts({"path": str(self.path), "rows": rows})
            self.assertEqual(self.path.read_bytes(), original)

    def test_incomplete_contact_is_rejected_without_losing_saved_contacts(self) -> None:
        self.save()
        original = self.path.read_bytes()
        with self.assertRaises(api.ApiError):
            api.save_contacts({"path": str(self.path), "rows": [self.row, {"dealership": "Demo Site 2"}]})
        self.assertEqual(self.path.read_bytes(), original)

    def test_session_reloads_the_file_that_was_saved(self) -> None:
        with mock.patch.object(report_engine, "has_active_session", return_value=True), \
             mock.patch.object(report_engine, "reload_session_recipient_mapping") as reload_mapping:
            result = self.save()
        reload_mapping.assert_called_once_with(self.path.resolve())
        self.assertTrue(result["session_refreshed"])

    def test_server_uses_configured_file_not_browser_supplied_path(self) -> None:
        settings_manager.save_settings({"recipient_mapping_path": str(self.path)})
        with mock.patch.object(settings_manager, "server_mode", return_value=True):
            result = api.save_contacts({"path": str(self.root / "ignored.csv"), "rows": [self.row]})
        self.assertEqual(result["path"], str(self.path.resolve()))
        self.assertFalse((self.root / "ignored.csv").exists())

    def test_json_contacts_also_round_trip_atomically(self) -> None:
        path = self.root / "contacts.json"
        email_builder.save_contact_rows(path, [self.row])
        self.assertEqual(email_builder.load_contact_rows(path)[0]["email"], self.row["email"])

    def test_frontend_warns_about_unsaved_contacts_and_sends_revision(self) -> None:
        script = (Path(__file__).resolve().parents[1] / "webapp" / "static" / "app.js").read_text(encoding="utf-8")
        self.assertIn("state.dirty || state.contactsDirty", script)
        self.assertIn("revision: state.contacts.revision", script)
        self.assertIn("Close without saving?", script)


if __name__ == "__main__":
    unittest.main()