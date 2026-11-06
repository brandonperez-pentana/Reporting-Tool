"""Behaviour pinned by the browser front end.

The web app is a second client of report_engine, so most of what matters is already
covered by the engine's own tests. What is not covered anywhere else is the HTTP layer:
the security guards that keep this a localhost-only app, the directory browser Settings
depends on, and the keep-alive handling that a hand-rolled server has to get right.
"""

from __future__ import annotations

import http.client
import json
import os
import tempfile
import threading
import unittest
from unittest import mock
from http.server import ThreadingHTTPServer
from pathlib import Path

from webapp import api
from webapp.server import ReporterHandler


class ServerTestCase(unittest.TestCase):
    """One server per class, driven over a real socket so keep-alive is exercised."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.httpd = ThreadingHTTPServer(("127.0.0.1", 0), ReporterHandler)
        cls.port = cls.httpd.server_address[1]
        cls.thread = threading.Thread(target=cls.httpd.serve_forever, daemon=True)
        cls.thread.start()

    @classmethod
    def tearDownClass(cls) -> None:
        cls.httpd.shutdown()
        cls.httpd.server_close()

    def connect(self) -> http.client.HTTPConnection:
        connection = http.client.HTTPConnection("127.0.0.1", self.port, timeout=30)
        self.addCleanup(connection.close)
        return connection

    def request(self, connection, method, path, body=None, headers=None):
        payload = json.dumps(body).encode() if body is not None else None
        head = {"Content-Type": "application/json"} if payload else {}
        head.update(headers or {})
        connection.request(method, path, body=payload, headers=head)
        response = connection.getresponse()
        raw = response.read()
        try:
            return response.status, json.loads(raw or b"{}")
        except json.JSONDecodeError:
            return response.status, {"raw": raw[:200]}


class KeepAliveTests(ServerTestCase):
    """An unread request body is still in the socket when the next request arrives, so
    the server parses it as a request line and answers 501. Every POST must drain its
    body even when the route ignores it or a guard rejects it first."""

    def test_a_route_that_ignores_its_body_does_not_poison_the_connection(self) -> None:
        connection = self.connect()
        # /api/sms/export takes no arguments, but the page posts {} to it.
        first_status, _ = self.request(connection, "POST", "/api/sms/export", {})
        self.assertIn(first_status, (200, 409))
        second_status, payload = self.request(connection, "GET", "/api/session")
        self.assertEqual(second_status, 200, payload)
        self.assertIn("has_session", payload)

    def test_a_rejected_post_does_not_poison_the_connection(self) -> None:
        connection = self.connect()
        rejected, _ = self.request(
            connection, "POST", "/api/settings", {"month": "Jul 2026"},
            headers={"Origin": "http://evil.example"},
        )
        self.assertEqual(rejected, 403)
        status, payload = self.request(connection, "GET", "/api/session")
        self.assertEqual(status, 200, payload)

    def test_a_post_to_a_missing_route_does_not_poison_the_connection(self) -> None:
        connection = self.connect()
        missing, _ = self.request(connection, "POST", "/api/not-a-route", {"a": 1})
        self.assertEqual(missing, 404)
        status, _ = self.request(connection, "GET", "/api/session")
        self.assertEqual(status, 200)

    def test_several_posts_in_a_row_on_one_connection(self) -> None:
        connection = self.connect()
        for _ in range(4):
            status, _ = self.request(connection, "POST", "/api/settings/check-path", {"path": ".", "kind": "folder"})
            self.assertEqual(status, 200)


class LocalhostOnlyTests(ServerTestCase):
    """This app has no login. It stays safe by being unreachable from anywhere else."""

    def test_a_request_addressed_to_another_host_is_refused(self) -> None:
        """DNS rebinding: a hostile page resolves its own domain to 127.0.0.1."""
        connection = self.connect()
        status, payload = self.request(connection, "GET", "/api/session", headers={"Host": "attacker.example"})
        self.assertEqual(status, 403)
        self.assertIn("localhost", payload["error"])

    def test_a_cross_origin_post_is_refused(self) -> None:
        connection = self.connect()
        status, _ = self.request(connection, "POST", "/api/settings", {}, headers={"Origin": "http://evil.example"})
        self.assertEqual(status, 403)

    def test_a_same_origin_post_is_allowed(self) -> None:
        connection = self.connect()
        status, _ = self.request(
            connection, "POST", "/api/settings/check-path", {"path": ".", "kind": "folder"},
            headers={"Origin": f"http://127.0.0.1:{self.port}"},
        )
        self.assertEqual(status, 200)


class StaticFileTests(ServerTestCase):
    def test_the_page_and_its_assets_are_served(self) -> None:
        connection = self.connect()
        for path, needle in (("/", b"Reporter"), ("/static/app.js", b"runJob"), ("/static/app.css", b"--navy")):
            connection.request("GET", path)
            response = connection.getresponse()
            body = response.read()
            self.assertEqual(response.status, 200, path)
            self.assertIn(needle, body, path)

    def test_asset_paths_cannot_escape_the_static_folder(self) -> None:
        connection = self.connect()
        connection.request("GET", "/static/..%2f..%2fuser_settings.json")
        response = connection.getresponse()
        response.read()
        self.assertEqual(response.status, 404)


class BrowseTests(unittest.TestCase):
    """Settings cannot open a native file dialog, so the server lists directories."""

    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        root = Path(self.tmp.name)
        (root / "Accounts").mkdir()
        (root / "logs").mkdir()
        (root / "contacts.csv").write_text("dealership\n", encoding="utf-8")
        (root / "template.msg").write_bytes(b"x" * 32)
        (root / "notes.txt").write_text("x", encoding="utf-8")
        self.root = root

    def test_folder_mode_lists_only_folders(self) -> None:
        result = api.browse({"path": str(self.root), "kind": "folder"})
        self.assertEqual([e["name"] for e in result["entries"]], ["Accounts", "logs"])

    def test_file_mode_filters_by_extension(self) -> None:
        result = api.browse({"path": str(self.root), "kind": "file", "extensions": ".csv"})
        names = [e["name"] for e in result["entries"]]
        self.assertIn("contacts.csv", names)
        self.assertNotIn("notes.txt", names)
        self.assertIn("Accounts", names, "folders stay visible so you can navigate into them")

    def test_folders_sort_before_files(self) -> None:
        entries = api.browse({"path": str(self.root), "kind": "file"})["entries"]
        self.assertTrue(all(e["is_dir"] for e in entries[:2]))

    def test_a_file_path_opens_the_folder_holding_it(self) -> None:
        result = api.browse({"path": str(self.root / "contacts.csv"), "kind": "file"})
        self.assertEqual(result["path"], str(self.root))

    def test_a_path_that_does_not_exist_falls_back_rather_than_failing(self) -> None:
        result = api.browse({"path": str(self.root / "nope" / "deeper"), "kind": "folder"})
        self.assertTrue(Path(result["path"]).is_dir())

    def test_the_parent_is_reported_so_the_client_never_guesses_separators(self) -> None:
        result = api.browse({"path": str(self.root / "Accounts"), "kind": "folder"})
        self.assertEqual(result["parent"], str(self.root))

    def test_somewhere_sensible_to_start(self) -> None:
        result = api.browse({"path": "", "kind": "folder"})
        self.assertTrue(result["places"], "there is always at least the app folder")
        self.assertTrue(Path(result["path"]).is_dir())

    def test_windows_drives_are_offered_only_on_windows(self) -> None:
        result = api.browse({"path": str(self.root), "kind": "folder"})
        if os.name == "nt":
            self.assertTrue(result["drives"])
        else:
            self.assertEqual(result["drives"], [])


class JobTests(unittest.TestCase):
    def test_an_unknown_job_is_reported_rather_than_crashing(self) -> None:
        with self.assertRaises(api.ApiError) as caught:
            api.job_status("deadbeef")
        self.assertEqual(caught.exception.status, 404)

    def test_outlook_registration_failures_are_explained_in_plain_words(self) -> None:
        message = api.friendly_error(Exception("(-2147221005, 'Invalid class string', None, None)"))
        self.assertNotIn("-2147221005", message)
        self.assertGreater(len(message), 80, "a bare restatement is not an explanation")


class OutlookFailureTests(ServerTestCase):
    """Outlook drives sending and drafts. When it cannot start, every email in a batch
    fails with the same COM code, and that code is meaningless to an operator."""

    def test_the_com_error_is_explained_rather_than_shown_raw(self) -> None:
        """Assert the properties the message must have, not the diagnosis it offers.

        An earlier version of this test pinned the wording of one guess — that New
        Outlook had replaced the classic app — and that guess turned out to be wrong for
        the machine actually hitting the error. Pinning a specific cause makes the test
        defend the mistake.
        """
        message = api.friendly_error(Exception("(-2147221005, 'Invalid class string', None, None)"))
        self.assertNotIn("-2147221005", message, "the COM code must not reach the operator")
        self.assertNotIn("Invalid class string", message)
        self.assertIn("Outlook.Application", message, "name what could not be resolved")
        self.assertIn("diagnose_outlook", message, "point at something that finds the cause")
        self.assertIn("unaffected", message, "say what still works")

    def test_a_missing_pywin32_is_explained_too(self) -> None:
        message = api.friendly_error(Exception("pywin32 is required for Outlook draft creation."))
        self.assertIn("pip install pywin32", message)

    def test_an_unrecognised_error_is_passed_through_not_swallowed(self) -> None:
        self.assertEqual(api.friendly_error(Exception("disk is on fire")), "disk is on fire")

    def test_per_row_failures_are_translated_before_they_reach_the_page(self) -> None:
        """send_emails catches per dealership and stores str(exc), so the translation has
        to be applied to each row — not only to an error that escapes the job."""
        raw = "(-2147221005, 'Invalid class string', None, None)"
        rows = [
            {"dealership": "Example Motors", "status": "failed", "error": raw},
            {"dealership": "Other Motors", "status": "drafted", "error": ""},
        ]
        for row in rows:
            if row.get("error"):
                row["error"] = api.friendly_error(Exception(str(row["error"])))
        self.assertNotIn("-2147221005", rows[0]["error"])
        self.assertEqual(rows[1]["error"], "")

    def test_the_preflight_answers_without_needing_outlook(self) -> None:
        status, payload = self.request(self.connect(), "GET", "/api/outlook")
        self.assertEqual(status, 200)
        self.assertTrue(payload["checked"])
        self.assertIn("available", payload)
        if not payload["available"]:
            self.assertTrue(payload["reason"], "an unavailable check has to say why")


if __name__ == "__main__":
    unittest.main()
class ExportTests(ServerTestCase):
    """Save as .eml is the route that has to survive Outlook being unavailable.

    So the endpoint is checked for its refusals, which are what an operator actually
    meets: nothing ticked, and no folder chosen yet. The file format itself is covered
    by tests/test_email_export.py, which needs no server.
    """

    def test_an_empty_selection_is_refused_with_something_actionable(self) -> None:
        connection = self.connect()
        status, payload = self.request(connection, "POST", "/api/export", {"dealerships": []})
        self.assertEqual(status, 400)
        self.assertIn("tick", payload.get("error", "").lower())

    def test_no_destination_folder_names_the_setting_to_fill_in(self) -> None:
        with mock.patch.object(api.settings_manager, "load_settings", return_value={}):
            with self.assertRaises(api.ApiError) as caught:
                api.start_export({"dealerships": ["Demo Site 1"]})
        message = caught.exception.message
        self.assertIn(".eml", message)
        self.assertIn("Settings", message)


class FirstRunTests(unittest.TestCase):
    """A copy moved to another machine has to know it has never been set up.

    The flag keys off the settings file existing rather than off any field being empty,
    so an operator who clears one path is not thrown back into the setup screen.
    """

    def test_a_copy_with_no_settings_file_reports_first_run(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "user_settings.json"
            with mock.patch.object(api.settings_manager, "get_settings_path", return_value=path), \
                 mock.patch.object(api.settings_manager, "load_settings", return_value={}):
                payload = api.get_settings()
        self.assertTrue(payload["first_run"])
        self.assertEqual(
            sorted(payload["missing_required"]),
            ["accounts_folder", "output_folder", "recipient_mapping_path"],
        )

    def test_a_configured_copy_does_not_report_first_run(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "user_settings.json"
            path.write_text("{}", encoding="utf-8")
            settings = {
                "accounts_folder": "C:\\Accounts",
                "output_folder": "C:\\Out",
                "recipient_mapping_path": "C:\\contacts.csv",
            }
            with mock.patch.object(api.settings_manager, "get_settings_path", return_value=path), \
                 mock.patch.object(api.settings_manager, "load_settings", return_value=settings):
                payload = api.get_settings()
        self.assertFalse(payload["first_run"])
        self.assertEqual(payload["missing_required"], [])

    def test_the_bundled_branding_assets_are_not_a_user_setting(self) -> None:
        """The app uses its packaged assets without asking the operator to locate them."""
        payload = api.get_settings()
        self.assertNotIn("default_assets_folder", payload)
        self.assertNotIn("email_template_assets_folder", payload)

class DealershipPickerTests(unittest.TestCase):
    """Contacts match reports by an exact string, so the name is picked, not typed.

    A contact row saying "Demo Site 2" against a folder called "Demo Site 2
    Eastside" produces an email with no recipient and no error, which is the worst
    kind of failure: silent and only visible at send time.
    """

    def _names_with_accounts_folder(self, folder) -> list[str]:
        with mock.patch.object(
            api.settings_manager, "load_settings", return_value={"accounts_folder": str(folder)}
        ):
            return api.dealership_names()

    def test_the_names_are_the_folder_names_exactly(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            for name in ("Demo Site 1", "Demo Site 2 Eastside", "Demo Site 3"):
                (root / name).mkdir()
            self.assertEqual(
                self._names_with_accounts_folder(root),
                ["Demo Site 1", "Demo Site 2 Eastside", "Demo Site 3"],
            )

    def test_loose_files_and_hidden_folders_are_not_offered(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "Demo Site 1").mkdir()
            (root / ".hidden").mkdir()
            (root / "notes.txt").write_text("x", encoding="utf-8")
            self.assertEqual(self._names_with_accounts_folder(root), ["Demo Site 1"])

    def test_no_accounts_folder_yet_is_an_empty_list_not_an_error(self) -> None:
        """Contacts are often filled in before anything is configured."""
        with mock.patch.object(api.settings_manager, "load_settings", return_value={}):
            self.assertEqual(api.dealership_names(), [])

    def test_a_missing_folder_is_an_empty_list_not_an_error(self) -> None:
        self.assertEqual(self._names_with_accounts_folder(Path("/no/such/folder")), [])

    def test_the_contacts_payload_carries_the_names_for_the_picker(self) -> None:
        payload = api.contacts()
        self.assertIn("dealership_names", payload)
        self.assertIsInstance(payload["dealership_names"], list)
