from __future__ import annotations

import csv
import io
import unittest
import zipfile

from tools.build_release import CONTACTS_HEADER, PREFIX, validate_archive


def _archive(files: dict[str, str]) -> zipfile.ZipFile:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, content in files.items():
            archive.writestr(name, content)
    buffer.seek(0)
    return zipfile.ZipFile(buffer)


class ReleaseArchiveValidationTests(unittest.TestCase):
    def _minimal_files(self) -> dict[str, str]:
        return {
            f"{PREFIX}/reporter_web.py": "# app",
            f"{PREFIX}/README.md": "# Reporter",
            f"{PREFIX}/requirements-web.txt": "",
            f"{PREFIX}/VERSION": "2026.09.12\n",
            f"{PREFIX}/data/contacts.csv": ",".join(CONTACTS_HEADER) + "\n",
        }

    def test_clean_header_only_package_passes(self) -> None:
        with _archive(self._minimal_files()) as archive:
            self.assertEqual(validate_archive(archive), [])

    def test_machine_settings_and_root_report_files_are_rejected(self) -> None:
        files = self._minimal_files()
        files[f"{PREFIX}/user_settings.json"] = "{}"
        files[f"{PREFIX}/Example Dealer - Jul 2026.txt"] = "customer report"
        with _archive(files) as archive:
            errors = validate_archive(archive)
        self.assertTrue(any("user_settings.json" in error for error in errors))
        self.assertTrue(any("root-level data file" in error for error in errors))

    def test_contacts_must_be_header_only(self) -> None:
        files = self._minimal_files()
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(CONTACTS_HEADER)
        writer.writerow(["Demo Site 1", "Person", "person@example.com", "", "", "yes", "", ""])
        files[f"{PREFIX}/data/contacts.csv"] = output.getvalue()
        with _archive(files) as archive:
            errors = validate_archive(archive)
        self.assertIn("data/contacts.csv must contain only the standard header row", errors)

    def test_non_demo_accounts_are_rejected(self) -> None:
        files = self._minimal_files()
        files[f"{PREFIX}/Accounts/Actual Dealer/report.pdf"] = "client data"
        with _archive(files) as archive:
            errors = validate_archive(archive)
        self.assertTrue(any("non-demo account data" in error for error in errors))

    def test_expected_synthetic_demo_account_layout_is_allowed(self) -> None:
        files = self._minimal_files()
        files[f"{PREFIX}/Accounts/Demo Site 1.xlsx"] = "synthetic workbook"
        files[f"{PREFIX}/Accounts/Demo Site 1/Demo Site 1 - Sales Report - July 26.pdf"] = "synthetic report"
        files[f"{PREFIX}/Accounts/Demo Site 1/Demo Site 1 - Service Report - July 26.pdf"] = "synthetic report"
        with _archive(files) as archive:
            self.assertEqual(validate_archive(archive), [])

    def test_alternate_dotenv_files_are_rejected(self) -> None:
        files = self._minimal_files()
        files[f"{PREFIX}/.env.production"] = "REPORTER_AI_API_KEY=secret"
        with _archive(files) as archive:
            errors = validate_archive(archive)
        self.assertTrue(any("machine-specific or secret file" in error for error in errors))


if __name__ == "__main__":
    unittest.main()
