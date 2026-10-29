"""Reporter on a server: settings, sessions and jobs belong to the signed-in Windows user."""

from __future__ import annotations

import http.client
import json
import os
import shutil
import sys
import tempfile
import threading
import time
import unittest
from http.server import ThreadingHTTPServer
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import report_engine  # noqa: E402
import settings_manager  # noqa: E402
from webapp import api  # noqa: E402
from webapp.server import ReporterHandler  # noqa: E402

ALICE = "CORP\\alice"
BOB = "CORP\\bob"


class ServerModeCase(unittest.TestCase):
    def setUp(self) -> None:
        self.root = Path(tempfile.mkdtemp(prefix="reporter-server-"))
        self.addCleanup(shutil.rmtree, self.root, True)
        env = mock.patch.dict(os.environ, {
            "REPORTER_MODE": "server",
            "REPORTER_DATA_ROOT": str(self.root),
            "REPORTER_ADMINS": "CORP\\Admin",
        })
        env.start()
        self.addCleanup(env.stop)


class PerUserSettingsTests(ServerModeCase):
    def test_each_user_keeps_their_own_portfolio(self) -> None:
        with settings_manager.as_user(ALICE):
            api.save_settings({"portfolio": ["Demo Site 1.xlsx"], "sender_name": "Alice"})
        with settings_manager.as_user(BOB):
            self.assertEqual(api.get_settings()["portfolio"], [])
            self.assertEqual(api.get_settings()["sender_name"], "")
        with settings_manager.as_user(ALICE):
            self.assertEqual(api.get_settings()["portfolio"], ["Demo Site 1.xlsx"])

    def test_shared_settings_reach_every_user_and_outputs_do_not_collide(self) -> None:
        with settings_manager.as_user("CORP\\admin"):
            api.save_settings({"accounts_folder": "D:\\Library"})
        with settings_manager.as_user(ALICE):
            alice = settings_manager.load_settings()
        with settings_manager.as_user(BOB):
            bob = settings_manager.load_settings()
        self.assertEqual(alice["accounts_folder"], "D:\\Library")
        self.assertEqual(bob["accounts_folder"], "D:\\Library")
        self.assertNotEqual(alice["output_folder"], bob["output_folder"])
        self.assertTrue(Path(alice["output_folder"]).is_relative_to(self.root / "users"))

    def test_a_user_cannot_change_shared_settings(self) -> None:
        with settings_manager.as_user(ALICE):
            with self.assertRaises(api.ApiError) as caught:
                api.save_settings({"accounts_folder": "C:\\Elsewhere"})
        self.assertEqual(caught.exception.status, 403)
        self.assertFalse(settings_manager.shared_settings_path().exists())

    def test_reposting_unchanged_shared_values_is_allowed(self) -> None:
        """The settings form posts every field, including the read-only ones."""
        with settings_manager.as_user(ALICE):
            current = api.get_settings()
            api.save_settings({"accounts_folder": current["accounts_folder"], "sender_title": "CRM"})
            self.assertEqual(api.get_settings()["sender_title"], "CRM")

    def test_the_user_file_never_holds_shared_keys(self) -> None:
        with settings_manager.as_user("CORP\\admin"):
            api.save_settings({"accounts_folder": "D:\\Library", "sender_name": "Admin"})
            personal = json.loads(settings_manager.get_settings_path().read_text(encoding="utf-8"))
        shared = json.loads(settings_manager.shared_settings_path().read_text(encoding="utf-8"))
        self.assertNotIn("accounts_folder", personal)
        self.assertNotIn("sender_name", shared)

    def test_a_login_cannot_escape_the_users_folder(self) -> None:
        for login in ("CORP\\..\\..\\evil", "../../evil", "a/b\\c", "CORP\\x:y"):
            with self.subTest(login=login):
                folder = settings_manager.user_folder(login).resolve()
                self.assertEqual(folder.parent, (self.root / "users").resolve())
        with self.assertRaises(ValueError):
            settings_manager.user_key("..")

    def test_the_desktop_copy_is_unaffected_without_a_user(self) -> None:
        legacy = Path(tempfile.mkdtemp()) / "user_settings.json"
        self.addCleanup(shutil.rmtree, legacy.parent, True)
        settings_manager.save_settings({"accounts_folder": "C:\\Mine"}, legacy)
        self.assertEqual(settings_manager.load_settings(legacy)["accounts_folder"], "C:\\Mine")


class PerUserSessionTests(ServerModeCase):
    def test_sessions_do_not_replace_each_other(self) -> None:
        session = mock.Mock(dealerships={"Demo Site 1": object()})
        with settings_manager.as_user(ALICE):
            report_engine._set_active_session(session)
            self.assertTrue(report_engine.has_dealership("Demo Site 1"))
        with settings_manager.as_user(BOB):
            self.assertFalse(report_engine.has_active_session())
        with settings_manager.as_user(ALICE):
            report_engine.clear_session()
            self.assertFalse(report_engine.has_active_session())


class JobOwnershipTests(ServerModeCase):
    def _wait(self, job_id: str) -> dict:
        for _ in range(100):
            job = api.job_status(job_id)
            if job["state"] != "running":
                return job
            time.sleep(0.05)
        self.fail("job did not finish")

    def test_a_job_runs_as_the_user_who_started_it(self) -> None:
        with settings_manager.as_user(ALICE):
            job_id = api._start_job("test", lambda _: settings_manager.current_user())["job_id"]
            job = self._wait(job_id)
        self.assertEqual(job["result"], ALICE.casefold())

    def test_another_user_cannot_see_the_job(self) -> None:
        with settings_manager.as_user(ALICE):
            job_id = api._start_job("test", lambda _: "done")["job_id"]
            self._wait(job_id)
        with settings_manager.as_user(BOB):
            with self.assertRaises(api.ApiError) as caught:
                api.job_status(job_id)
        self.assertEqual(caught.exception.status, 404)

    def test_drafts_and_browsing_are_refused_on_the_server(self) -> None:
        with settings_manager.as_user(ALICE):
            with self.assertRaises(api.ApiError):
                api.start_drafts({"dealerships": ["Demo Site 1"]})
            with self.assertRaises(api.ApiError) as caught:
                api.browse({"path": "C:\\"})
            self.assertEqual(caught.exception.status, 403)
            self.assertFalse(api.outlook_status()["available"])


class SignInOverHttpTests(ServerModeCase):
    def setUp(self) -> None:
        super().setUp()
        self.httpd = ThreadingHTTPServer(("127.0.0.1", 0), ReporterHandler)
        threading.Thread(target=self.httpd.serve_forever, daemon=True).start()
        self.addCleanup(self.httpd.server_close)
        self.addCleanup(self.httpd.shutdown)

    def _get(self, path: str, headers: dict[str, str] | None = None) -> tuple[int, dict]:
        connection = http.client.HTTPConnection("127.0.0.1", self.httpd.server_address[1], timeout=30)
        self.addCleanup(connection.close)
        connection.request("GET", path, headers=headers or {})
        response = connection.getresponse()
        return response.status, json.loads(response.read() or b"{}")

    def test_no_identity_is_refused(self) -> None:
        status, _ = self._get("/api/settings")
        self.assertEqual(status, 401)

    def test_the_user_header_is_ignored_unless_trusted(self) -> None:
        status, _ = self._get("/api/settings", {"X-Reporter-User": ALICE})
        self.assertEqual(status, 401)

    def test_a_trusted_user_header_signs_the_user_in(self) -> None:
        with mock.patch.dict(os.environ, {"REPORTER_TRUST_USER_HEADER": "1"}):
            status, payload = self._get("/api/settings", {"X-Reporter-User": ALICE})
        self.assertEqual(status, 200, payload)
        self.assertEqual(payload["user"], ALICE.casefold())
        self.assertFalse(payload["is_admin"])

    def test_the_iis_token_signs_the_user_in(self) -> None:
        with mock.patch("webapp.server._user_from_iis_token", return_value=BOB):
            status, payload = self._get("/api/settings", {"X-IIS-WindowsAuthToken": "1a2b"})
        self.assertEqual(status, 200, payload)
        self.assertEqual(payload["user"], BOB.casefold())

    def test_a_dev_user_signs_in_a_plain_browser_for_local_testing(self) -> None:
        with mock.patch.dict(os.environ, {"REPORTER_DEV_USER": ALICE}):
            status, payload = self._get("/api/settings")
        self.assertEqual(status, 200, payload)
        self.assertEqual(payload["user"], ALICE.casefold())

    def test_a_real_sign_in_wins_over_the_dev_user(self) -> None:
        with mock.patch.dict(os.environ, {"REPORTER_DEV_USER": ALICE}), \
             mock.patch("webapp.server._user_from_iis_token", return_value=BOB):
            status, payload = self._get("/api/settings", {"X-IIS-WindowsAuthToken": "1a2b"})
        self.assertEqual(status, 200, payload)
        self.assertEqual(payload["user"], BOB.casefold())


if __name__ == "__main__":
    unittest.main()
