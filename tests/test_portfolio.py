""""My dealerships": each operator runs only their own workbooks from a shared library."""

from __future__ import annotations

import json
import shutil
import sys
import tempfile
import unittest
import warnings
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

warnings.filterwarnings("ignore")

import dealership_insights_local as core  # noqa: E402
import excel_reader  # noqa: E402
from webapp import api  # noqa: E402

ACCOUNTS = ROOT / "Accounts"


def workbooks() -> list[Path]:
    return sorted(ACCOUNTS.glob("*.xls[xm]"))


@unittest.skipUnless(len(workbooks()) >= 3, "demo workbooks not present")
class PortfolioDiscoveryTests(unittest.TestCase):
    def test_only_the_listed_workbooks_are_discovered(self) -> None:
        chosen = [workbooks()[0].name, workbooks()[2].name.upper()]
        found = excel_reader.discover_masters(ACCOUNTS, recursive=False, include=chosen)
        self.assertEqual(sorted(m.path.name for m in found),
                         sorted([workbooks()[0].name, workbooks()[2].name]))

    def test_an_empty_list_means_the_whole_library(self) -> None:
        everyone = excel_reader.discover_masters(ACCOUNTS, recursive=False)
        self.assertEqual(len(excel_reader.discover_masters(ACCOUNTS, recursive=False, include=[])),
                         len(everyone))

    def test_a_run_processes_only_the_portfolio(self) -> None:
        chosen = [workbooks()[1].name]
        groups = core.collect_report_groups(ACCOUNTS, 7, 2026, include=chosen)
        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0].workbook_months, excel_reader.discover_masters(
            ACCOUNTS, recursive=False, include=chosen)[0].months)

    def test_a_stale_portfolio_does_not_fall_back_to_the_whole_library(self) -> None:
        groups = core.collect_report_groups(ACCOUNTS, 7, 2026, include=["Renamed Dealer.xlsx"])
        self.assertEqual(groups, [])

    def test_a_pdf_only_library_ignores_the_portfolio(self) -> None:
        folder = Path(tempfile.mkdtemp())
        try:
            site = folder / "Some Dealership"
            site.mkdir()
            (site / "Some Dealership - Sales Report - July 26.pdf").write_bytes(b"%PDF-1.4\n")
            groups = core.collect_report_groups(folder, 7, 2026, include=["Demo Site 1.xlsx"])
            self.assertTrue(groups)
        finally:
            shutil.rmtree(folder, ignore_errors=True)


class PortfolioSettingsTests(unittest.TestCase):
    def test_entries_are_bare_filenames_without_duplicates(self) -> None:
        cleaned = api.clean_portfolio(["C:\\Lib\\Demo Site 1.xlsx", "demo site 1.XLSX", " ", None, "Demo Site 2.xlsx"])
        self.assertEqual(cleaned, ["Demo Site 1.xlsx", "Demo Site 2.xlsx"])
        self.assertEqual(api.clean_portfolio("Demo Site 1.xlsx"), [])

    def test_the_portfolio_is_saved_and_survives_other_settings_saves(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "user_settings.json"
            with mock.patch.object(api.settings_manager, "get_settings_path", return_value=path):
                saved = api.save_settings({"portfolio": ["Demo Site 1.xlsx", "Demo Site 2.xlsx"]})
                self.assertEqual(saved["portfolio"], ["Demo Site 1.xlsx", "Demo Site 2.xlsx"])
                api.save_settings({"sender_name": "Sam"})
            self.assertEqual(json.loads(path.read_text(encoding="utf-8"))["portfolio"],
                             ["Demo Site 1.xlsx", "Demo Site 2.xlsx"])

    def test_options_list_workbooks_and_flag_missing_selections(self) -> None:
        folder = Path(tempfile.mkdtemp())
        try:
            (folder / "Demo Site 1.xlsx").write_bytes(b"")
            (folder / "~$Demo Site 1.xlsx").write_bytes(b"")
            settings = {"accounts_folder": str(folder), "portfolio": ["demo site 1.xlsx", "Gone.xlsx"]}
            with mock.patch.object(api.settings_manager, "load_settings", return_value=settings):
                options = api.portfolio_options()
        finally:
            shutil.rmtree(folder, ignore_errors=True)
        self.assertEqual(options["workbooks"], ["Demo Site 1.xlsx"])
        self.assertEqual(options["missing"], ["Gone.xlsx"])


class SharedTitlesAndFilenamesTests(unittest.TestCase):
    def test_workbooks_sharing_a_title_are_each_named_after_their_file(self) -> None:
        """Sites in one group often share the group's title; only one used to run."""
        folder = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, folder, True)
        for name in ("Demo Site 3 North.xlsx", "Demo Site 3 South.xlsx", "Demo Site 1.xlsx"):
            (folder / name).write_bytes(b"")
        titles = {"Demo Site 3 North": "Demo Site 3", "Demo Site 3 South": "Demo Site 3",
                  "Demo Site 1": "Demo Site 1 Group"}
        with mock.patch.object(excel_reader, "_title_from_drawings",
                               side_effect=lambda path: (titles[Path(path).stem], "")), \
             mock.patch.object(excel_reader, "read_months", return_value=("Jun 2026", "Jul 2026")), \
             mock.patch.object(excel_reader, "list_months", return_value=["Jun 2026", "Jul 2026"]):
            found = excel_reader.discover_masters(folder, recursive=False)
        self.assertEqual(sorted(m.dealership for m in found),
                         ["Demo Site 1 Group", "Demo Site 3 North", "Demo Site 3 South"])

    def test_filenames_keep_their_words(self) -> None:
        self.assertEqual(core.safe_file_stem("Demo Site 4 CJD/Volvo"), "Demo Site 4 CJD-Volvo")
        self.assertEqual(core.safe_file_stem("Demo Site 5 O'Brien"), "Demo Site 5 O'Brien")
        self.assertEqual(core.safe_file_stem("Demo Site 6 Nissan & Kia"), "Demo Site 6 Nissan & Kia")


class PdfSourceFolderTests(unittest.TestCase):
    """The PDF folders can live somewhere other than beside the workbooks."""

    def setUp(self) -> None:
        self.accounts = Path(tempfile.mkdtemp())
        self.pdfs = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.accounts, True)
        self.addCleanup(shutil.rmtree, self.pdfs, True)
        self.site = self.pdfs / "Demo Site 7 Holdings NZ"
        self.site.mkdir()
        (self.site / "Demo Site 7 Holdings NZ - Sales Report - July 26.pdf").write_bytes(b"%PDF-1.4\n")

    def test_a_workbook_finds_its_pdfs_in_the_separate_folder_by_filename(self) -> None:
        """The title box says "DS7 NZ"; the PDF folder is named after the workbook file."""
        master = excel_reader.MasterFile(
            path=self.accounts / "Demo Site 7 Holdings NZ.xlsx", dealership="DS7 NZ",
            month_label="Jul 2026", previous_month_label="Jun 2026", months=["Jul 2026"])
        with mock.patch.object(excel_reader, "discover_masters", return_value=[master]):
            groups = core.find_workbook_report_groups(self.accounts, 7, 2026, pdf_dir=self.pdfs)
        self.assertEqual(groups[0].folder_path, self.site)
        self.assertIsNotNone(groups[0].sales_pdf)

    def test_a_pdf_only_library_is_read_from_the_pdf_folder(self) -> None:
        groups = core.collect_report_groups(self.accounts, 7, 2026, pdf_dir=self.pdfs)
        self.assertTrue(groups)

    def test_the_setting_is_saved(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "user_settings.json"
            with mock.patch.object(api.settings_manager, "get_settings_path", return_value=path):
                saved = api.save_settings({"pdf_source_folder": str(self.pdfs)})
        self.assertEqual(saved["pdf_source_folder"], str(self.pdfs))


if __name__ == "__main__":
    unittest.main()
