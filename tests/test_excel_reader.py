"""The Excel reader: same figures as the PDF path, or it is not a replacement.

The demo workbooks ship with the PDFs that were rendered from them, which makes the
strongest test available almost free: read the month both ways and require the engine
to produce the same copy. That is the property this whole ingestion change rests on,
so it is asserted directly rather than approximated by unit tests on the parts.

The rest guards the traps found while building it -- a percentage read as a fraction,
a campaign name read as a figure, a rate printed to more decimal places than the PDF
attached beside it.
"""

from __future__ import annotations

import datetime as dt
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
import settings_manager  # noqa: E402
from reporting import report_breakdowns  # noqa: E402

ACCOUNTS = ROOT / "Accounts"


def workbooks() -> list[Path]:
    """The library keeps one workbook per dealership at the top level."""
    return sorted(ACCOUNTS.glob("*.xls[xm]"))


def pdf_pair(book: Path) -> tuple[Path | None, Path | None]:
    """The exported PDFs beside a workbook, if they were exported at all."""
    import dealership_insights_local as _core

    folder = _core._pdf_folder_for(book.parent, book.stem)
    if folder is None:
        return (None, None)
    return (next(iter(folder.glob("*Sales Report*.pdf")), None),
            next(iter(folder.glob("*Service Report*.pdf")), None))


def with_engine(engine: str, work):
    original = settings_manager.load_settings
    settings_manager.load_settings = lambda *a, **k: {**original(*a, **k), "insight_engine": engine}
    try:
        return work()
    finally:
        settings_manager.load_settings = original


@unittest.skipUnless(workbooks(), "demo workbooks not present")
class ReadsTheMonthTests(unittest.TestCase):
    def setUp(self) -> None:
        self.book = workbooks()[0]
        self.read = excel_reader.read_master(self.book, fallback_name=self.book.stem)

    def test_it_finds_the_month_pair(self) -> None:
        self.assertTrue(self.read.month_label)
        self.assertTrue(self.read.previous_month_label)
        self.assertNotEqual(self.read.month_label, self.read.previous_month_label)

    def test_it_names_the_dealership(self) -> None:
        self.assertTrue(self.read.dealership)

    def test_it_reads_the_first_two_tabs_whatever_they_are_called(self) -> None:
        """Tab position is the contract, not sheet names."""
        self.assertGreaterEqual(len(self.read.metrics), 15)
        self.assertIn("sales_total_leads", self.read.metrics)
        self.assertIn("service_reminders_sent", self.read.metrics)

    def test_every_block_in_the_sheet_is_accounted_for(self) -> None:
        """An unmapped block means the workbook changed and nobody noticed."""
        unexpected = [u for u in self.read.unmapped if "|" not in u]
        self.assertEqual(unexpected, [], f"unmapped blocks: {unexpected}")

    def test_it_reports_no_warnings_on_a_healthy_workbook(self) -> None:
        self.assertEqual(self.read.warnings, [])

    def test_the_master_is_never_modified(self) -> None:
        before = self.book.stat().st_mtime_ns, self.book.stat().st_size
        excel_reader.read_master(self.book)
        excel_reader.build_breakdowns(self.read, self.book)
        self.assertEqual((self.book.stat().st_mtime_ns, self.book.stat().st_size), before)

    def test_it_reads_from_bytes_as_well_as_a_path(self) -> None:
        """Hosted, the workbook arrives over the network and never touches disk."""
        from_bytes = excel_reader.read_master(self.book.read_bytes())
        self.assertEqual(from_bytes.metrics.keys(), self.read.metrics.keys())
        self.assertEqual(from_bytes.month_label, self.read.month_label)


class ValueFormattingTests(unittest.TestCase):
    def test_a_percentage_is_scaled_by_its_number_format(self) -> None:
        """0.0795 in a percent cell is 8%, not 0.08%. Getting this wrong is silent."""
        self.assertEqual(excel_reader._raw_value(0.0795, "0%"), "8%")
        self.assertEqual(excel_reader._raw_value(0.5317, "0.0%"), "53.2%")

    def test_decimals_follow_the_sheet_not_our_taste(self) -> None:
        """The email must agree with the PDF attached beside it."""
        self.assertEqual(excel_reader._raw_value(0.0795, "0%"), "8%")
        self.assertEqual(excel_reader._decimals("0.00%"), 2)
        self.assertEqual(excel_reader._decimals("0%"), 0)

    def test_currency_keeps_its_marker(self) -> None:
        self.assertEqual(excel_reader._raw_value(44183, '"$"#,##0'), "$44,183")

    def test_a_campaign_name_is_not_a_figure(self) -> None:
        """The stop rule that keeps the reader out of the campaign list beside a block."""
        self.assertTrue(excel_reader._looks_like_value(136))
        self.assertTrue(excel_reader._looks_like_value("16:35 hrs"))
        self.assertTrue(excel_reader._looks_like_value("50:00 min"))
        self.assertFalse(excel_reader._looks_like_value("GWM Closed RO"))
        self.assertFalse(excel_reader._looks_like_value("NEW Prospect"))


@unittest.skipUnless(workbooks(), "demo workbooks not present")
class ParityWithThePdfPathTests(unittest.TestCase):
    """The workbooks and their PDFs are the same month, so both routes must agree."""

    def test_every_metric_matches_the_pdf_extraction(self) -> None:
        checked = 0
        for book in workbooks():
            sales_pdf, service_pdf = pdf_pair(book)
            if not (sales_pdf and service_pdf):
                continue
            from_pdf = core.parse_report_metrics(
                core.extract_pdf_text(sales_pdf), core.extract_pdf_text(service_pdf))
            from_excel = excel_reader.read_master(book).metrics
            shared = set(from_pdf) & set(from_excel)
            self.assertGreaterEqual(len(shared), 15, f"{book.stem}: too few shared metrics")
            for name in sorted(shared):
                with self.subTest(dealership=book.stem, metric=name):
                    for field in ("previous_value", "current_value"):
                        self.assertAlmostEqual(
                            float(from_pdf[name][field]), float(from_excel[name][field]), places=2,
                            msg=f"{name}.{field} differs between the PDF and the workbook")
                    checked += 1
        self.assertGreater(checked, 100, "parity check did not actually run")

    def test_the_engine_writes_the_same_copy_from_either_source(self) -> None:
        """The property the whole ingestion change rests on."""
        book = workbooks()[0]
        sales_pdf, service_pdf = pdf_pair(book)
        if not (sales_pdf and service_pdf):
            self.skipTest("no PDFs beside this workbook")

        sales_text = core.extract_pdf_text(sales_pdf)
        service_text = core.extract_pdf_text(service_pdf)
        from_pdf = with_engine("funnel", lambda: core.generate_insights_from_metrics(
            core.parse_report_metrics(sales_text, service_text),
            "success", "Demo Dealership", "Jul 2026", sales_text, service_text))

        read = excel_reader.read_master(book)
        breakdowns = excel_reader.build_breakdowns(read, book)
        with mock.patch.object(report_breakdowns, "parse_breakdowns", return_value=breakdowns):
            from_excel = with_engine("funnel", lambda: core.generate_insights_from_metrics(
                read.metrics, "success", "Demo Dealership", "Jul 2026", "excel", "excel"))

        self.assertEqual(from_pdf[0], from_excel[0], "insights differ")
        self.assertEqual(from_pdf[1], from_excel[1], "summary differs")
        self.assertEqual(from_pdf[2], from_excel[2], "SMS differs")


@unittest.skipUnless(workbooks(), "demo workbooks not present")
class BreakdownsForTheFunnelEngineTests(unittest.TestCase):
    def setUp(self) -> None:
        self.book = workbooks()[0]
        self.read = excel_reader.read_master(self.book)
        self.breakdowns = excel_reader.build_breakdowns(self.read, self.book)

    def test_it_returns_the_shape_the_funnel_engine_expects(self) -> None:
        self.assertEqual(
            set(self.breakdowns),
            {"lead_sources", "pipeline_stages", "sales_campaigns", "service_campaigns"})

    def test_lead_sources_carry_both_months(self) -> None:
        sources = self.breakdowns["lead_sources"]
        self.assertTrue(sources)
        for row in sources:
            self.assertIsNotNone(row.leads_current, f"{row.source} has no current figure")

    def test_campaign_blocks_carry_the_previous_month(self) -> None:
        """Without it the campaign-return insight silently drops out of the month."""
        block = self.breakdowns["service_campaigns"]
        self.assertIsNotNone(block.stated_roi)
        self.assertIsNotNone(block.previous_roi)

    def test_named_campaigns_are_capped_at_what_the_report_prints(self) -> None:
        for key in ("sales_campaigns", "service_campaigns"):
            self.assertLessEqual(len(self.breakdowns[key].named), 3)

    def test_pipeline_stages_come_from_the_third_tab(self) -> None:
        stages = self.breakdowns["pipeline_stages"]
        self.assertIn("Contacted", stages)
        self.assertIn(self.read.month_label, stages["Contacted"])


class WalkInIsALeadSourceTests(unittest.TestCase):
    def test_a_hyphenated_walk_in_row_is_carried(self) -> None:
        """Workbooks print "WALK-IN"; dropping it credited a dealership's walk-in fall to Phone Up."""
        def block(walk_in: tuple[int, int], phone: tuple[int, int]) -> dict:
            return {"WALK-IN": {"previous": walk_in[0], "current": walk_in[1]},
                    "PHONE UP": {"previous": phone[0], "current": phone[1]}}

        read = excel_reader.WorkbookRead(
            dealership="Demo", month_label="Aug 2026", previous_month_label="Jul 2026",
            breakdowns={"NUMBER OF LEADS": block((147, 133), (56, 55)),
                        "NUMBER OF DEALS": block((72, 53), (25, 20))},
            pipeline={"Contacted": {}},
        )
        rows = excel_reader.build_breakdowns(read, b"")["lead_sources"]
        walk_in = next(row for row in rows if row.source == "Walk-in")
        self.assertEqual((walk_in.deals_previous, walk_in.deals_current), (72, 53))


class DealershipTitleTests(unittest.TestCase):
    @staticmethod
    def package(*runs: str) -> bytes:
        import io
        import zipfile

        body = "".join(f"<a:t>{run}</a:t>" for run in runs)
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr("xl/drawings/drawing1.xml", f"<xdr:wsDr>{body}</xdr:wsDr>")
        return buffer.getvalue()

    def test_an_escaped_ampersand_reads_as_an_ampersand(self) -> None:
        title = excel_reader._title_from_drawings(self.package("Demo Site 6 Nissan &amp; Kia", "August 2026"))
        self.assertEqual(title, ("Demo Site 6 Nissan & Kia", "August 2026"))

    def test_a_section_heading_is_not_a_dealership_name(self) -> None:
        for heading in ("SALES", "Dealer Name", "MONTHLY REPORT"):
            with self.subTest(heading=heading):
                self.assertEqual(excel_reader._title_from_drawings(self.package(heading)), ("", ""))


if __name__ == "__main__":
    unittest.main()


@unittest.skipUnless(workbooks(), "demo workbooks not present")
class MonthComesFromTheWorkbookTests(unittest.TestCase):
    """A master has no month in its filename, so the sheet has to say."""

    def test_the_month_is_read_from_the_third_tab(self) -> None:
        previous, current = excel_reader.read_months(workbooks()[0])
        self.assertTrue(current)
        self.assertTrue(previous)
        self.assertNotEqual(previous, current)

    def test_read_master_agrees_with_the_third_tab(self) -> None:
        book = workbooks()[0]
        previous, current = excel_reader.read_months(book)
        read = excel_reader.read_master(book)
        self.assertEqual(read.month_label, current)
        self.assertEqual(read.previous_month_label, previous)

    def test_the_reported_month_is_one_the_workbook_has_data_for(self) -> None:
        """The month must come from a populated column, not merely a dated one."""
        for book in workbooks():
            _, label = excel_reader.read_months(book)
            with self.subTest(book=book.stem):
                self.assertTrue(label)
                self.assertIn(label, excel_reader.available_months(ACCOUNTS))


@unittest.skipUnless(workbooks(), "demo workbooks not present")
class DiscoveryTests(unittest.TestCase):
    def setUp(self) -> None:
        self.found = excel_reader.discover_masters(ACCOUNTS)

    def test_it_finds_one_workbook_per_dealership(self) -> None:
        self.assertEqual(len(self.found), len(workbooks()))
        names = [m.dealership for m in self.found]
        self.assertEqual(len(names), len(set(names)), "dealership names must be unique")

    def test_every_workbook_reports_its_month(self) -> None:
        for master in self.found:
            with self.subTest(dealership=master.dealership):
                self.assertTrue(master.is_readable)
                self.assertTrue(master.month_label)

    def test_it_names_dealerships_without_a_month_in_the_filename(self) -> None:
        for master in self.found:
            with self.subTest(dealership=master.dealership):
                self.assertNotRegex(master.dealership, r"(?i)\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\b")

    def test_excel_lock_files_are_ignored(self) -> None:
        """A workbook open in Excel leaves a ~$ file beside it that is not a workbook."""
        lock = ACCOUNTS / "~$Demo Site.xlsx"
        lock.write_bytes(b"not a workbook")
        try:
            self.assertEqual(len(excel_reader.discover_masters(ACCOUNTS)), len(self.found))
        finally:
            lock.unlink()

    def test_available_months_lists_what_exists(self) -> None:
        months = excel_reader.available_months(ACCOUNTS)
        self.assertTrue(months)
        self.assertIn(self.found[0].month_label, months)


class MultiYearDataTabTests(unittest.TestCase):
    """Real masters run across years and keep growing: 2025, 2026, then 2027.

    The columns for months that have not happened yet already exist, so "rightmost
    dated column" is the wrong rule and "rightmost column with figures under it" is the
    right one. The year matters as much as the month -- comparing month names alone
    puts December 2025 after January 2026.
    """

    @staticmethod
    def _workbook(path: Path, years, filled_through: "dt.datetime", *, skip=None) -> Path:
        from openpyxl import Workbook

        book = Workbook()
        book.active.title = "Sales Report"
        book.create_sheet("Service Report")
        sheet = book.create_sheet("Data")
        sheet.cell(row=2, column=2, value="Month")
        column = 4
        for year in years:
            for month in range(1, 13):
                moment = dt.datetime(year, month, 1)
                sheet.cell(row=2, column=column, value=moment)
                if moment <= filled_through and moment != skip:
                    for row in range(3, 40):
                        sheet.cell(row=row, column=column, value=5)
                column += 1
        book.save(path)
        return path

    def setUp(self) -> None:
        self.folder = Path(tempfile.mkdtemp())

    def tearDown(self) -> None:
        shutil.rmtree(self.folder, ignore_errors=True)

    def test_it_picks_the_latest_month_that_has_figures(self) -> None:
        book = self._workbook(self.folder / "a.xlsx", (2025, 2026, 2027), dt.datetime(2026, 3, 1))
        self.assertEqual(excel_reader.read_months(book), ("Feb 2026", "Mar 2026"))

    def test_columns_for_future_months_are_ignored(self) -> None:
        """2027 exists in the sheet long before it exists in the year."""
        book = self._workbook(self.folder / "b.xlsx", (2025, 2026, 2027), dt.datetime(2025, 12, 1))
        self.assertEqual(excel_reader.read_months(book), ("Nov 2025", "Dec 2025"))

    def test_the_previous_month_crosses_the_year_boundary(self) -> None:
        """December 2025 precedes January 2026; month names alone get this backwards."""
        book = self._workbook(self.folder / "c.xlsx", (2025, 2026, 2027), dt.datetime(2026, 1, 1))
        self.assertEqual(excel_reader.read_months(book), ("Dec 2025", "Jan 2026"))

    def test_a_month_nobody_filled_in_is_skipped(self) -> None:
        book = self._workbook(self.folder / "d.xlsx", (2025, 2026), dt.datetime(2026, 3, 1),
                              skip=dt.datetime(2025, 12, 1))
        self.assertEqual(excel_reader.read_months(book), ("Feb 2026", "Mar 2026"))

    def test_an_untouched_workbook_reports_no_month(self) -> None:
        book = self._workbook(self.folder / "e.xlsx", (2026, 2027), dt.datetime(2020, 1, 1))
        self.assertEqual(excel_reader.read_months(book), ("", ""))

    def test_available_months_sorts_across_years(self) -> None:
        """Newest first, by date -- December 2025 must not outrank January 2026."""
        self._workbook(self.folder / "one.xlsx", (2025, 2026), dt.datetime(2025, 11, 1))
        self._workbook(self.folder / "two.xlsx", (2025, 2026), dt.datetime(2026, 2, 1))
        months = excel_reader.available_months(self.folder)

        self.assertEqual(months[0], "Feb 2026")
        self.assertIn("Nov 2025", months)
        self.assertLess(months.index("Jan 2026"), months.index("Dec 2025"))

        def key(label: str) -> tuple[int, int]:
            name, year = label.split()
            return (int(year), excel_reader.MONTH_NAMES.index(name) + 1)

        self.assertEqual(months, sorted(months, key=key, reverse=True))
