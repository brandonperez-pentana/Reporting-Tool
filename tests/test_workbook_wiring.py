"""The run loop reading a library of workbooks rather than a folder of PDFs.

The library keeps one workbook per dealership at the top level and a folder of
exported PDFs beside it. These tests pin the division of labour that follows from
that: the workbook decides who exists, which month it holds and every figure; the
PDFs are the attachment and nothing else.
"""

from __future__ import annotations

import sys
import unittest
import warnings
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

warnings.filterwarnings("ignore")

import dealership_insights_local as core  # noqa: E402
import report_engine  # noqa: E402

ACCOUNTS = ROOT / "Accounts"


def workbooks() -> list[Path]:
    """The library keeps one workbook per dealership at the top level."""
    return sorted(ACCOUNTS.glob("*.xls[xm]")) or sorted(ACCOUNTS.glob("*.xlsx"))


@unittest.skipUnless(workbooks(), "demo workbooks not present")
class DiscoveryPrefersWorkbooksTests(unittest.TestCase):
    def test_the_library_is_discovered_from_workbooks(self) -> None:
        groups = core.collect_report_groups(ACCOUNTS, 7, 2026)
        self.assertTrue(groups)
        self.assertTrue(any(g.workbook or g.workbook_month for g in groups),
                        "expected workbook-driven groups")

    def test_a_pdf_only_library_still_runs(self) -> None:
        """The two sources have to coexist while dealerships are moved across."""
        import tempfile, shutil

        folder = Path(tempfile.mkdtemp())
        try:
            site = folder / "Some Dealership"
            site.mkdir()
            (site / "Some Dealership - Sales Report - July 26.pdf").write_bytes(b"%PDF-1.4\n")
            groups = core.collect_report_groups(folder, 7, 2026)
            self.assertTrue(groups)
            self.assertTrue(all(g.workbook is None for g in groups))
        finally:
            shutil.rmtree(folder, ignore_errors=True)

    def test_months_come_from_the_workbooks_not_filenames(self) -> None:
        months = report_engine.detect_available_report_months(ACCOUNTS)
        self.assertTrue(months)
        for label in months:
            with self.subTest(label=label):
                self.assertRegex(label, r"^[A-Z][a-z]{2} \d{4}$")


@unittest.skipUnless(workbooks(), "demo workbooks not present")
class ProcessingFromAWorkbookTests(unittest.TestCase):
    def _group(self, month: int, year: int):
        groups = core.collect_report_groups(ACCOUNTS, month, year)
        self.assertTrue(groups)
        return groups[0]

    def test_a_month_the_workbook_holds_produces_figures(self) -> None:
        group = self._group(7, 2026)
        result = core.process_dealership_group(group, 7, 2026, "Jul 2026")
        self.assertEqual(result.report_status, "success")
        self.assertGreaterEqual(len(result.parsed_metrics), 15)
        self.assertEqual(len(result.bullet_insights), core.PRIORITY_INSIGHT_COUNT)

    def test_a_month_the_workbook_does_not_hold_says_which_it_has(self) -> None:
        """The operator should not go hunting for a file that is not the source.

        The month here has to be one genuinely outside the history. February used to
        qualify, back when only the workbook's latest month could be reported.
        """
        group = self._group(11, 2026)
        result = core.process_dealership_group(group, 11, 2026, "Nov 2026")
        self.assertEqual(result.report_status, "blocked")
        self.assertTrue(
            any("covers" in note or "currently holds" in note for note in result.notes),
            f"expected the months it does hold to be named, got {result.notes}")

    def test_a_missing_pdf_costs_the_attachment_not_the_month(self) -> None:
        group = self._group(7, 2026)
        stripped = core.DealershipReportGroup(
            dealership_name=group.dealership_name,
            group_key=group.group_key,
            folder_path=group.folder_path,
            sales_pdf=None,
            service_pdf=None,
            workbook=group.workbook,
            workbook_month=group.workbook_month,
        )
        result = core.process_dealership_group(stripped, 7, 2026, "Jul 2026")
        self.assertEqual(result.report_status, "success")
        self.assertGreaterEqual(len(result.parsed_metrics), 15)
        self.assertTrue(any("to attach" in note for note in result.notes))

    def test_pdfs_are_matched_to_the_workbook_for_attaching(self) -> None:
        group = self._group(7, 2026)
        self.assertIsNotNone(group.sales_pdf, "expected an exported PDF to attach")


if __name__ == "__main__":
    unittest.main()
