"""Reporting a month other than the workbook's latest.

A master holds years of history in its third tab, so a dealership can be reported on
any month it has figures for -- not only the one the front tabs happen to render. These
tests pin the two rules that makes safe:

* the latest month still comes from the rendered report, so its figures are the ones
  printed on the PDF attached to the email;
* an earlier month read from the history tab agrees with the rendered report when both
  describe the same month.

Without the second, the same month would read differently depending on which tab it
came from, and nobody would know which figure the dealership was sent.
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
import excel_reader  # noqa: E402
import report_engine  # noqa: E402

ACCOUNTS = ROOT / "Accounts"


def workbooks() -> list[Path]:
    return sorted(ACCOUNTS.glob("*.xls[xm]"))


@unittest.skipUnless(workbooks(), "demo workbooks not present")
class TheWorkbookOffersEveryMonthItHoldsTests(unittest.TestCase):
    def setUp(self) -> None:
        self.book = workbooks()[0]

    def test_it_lists_more_than_the_latest_month(self) -> None:
        months = excel_reader.list_months(self.book)
        self.assertGreater(len(months), 1,
                           "a master holds history; one month means the history tab was missed")

    def test_the_months_are_in_order_oldest_first(self) -> None:
        months = excel_reader.list_months(self.book)
        self.assertEqual(months, sorted(months, key=_sort_key))

    def test_the_earliest_month_is_not_offered_for_reporting(self) -> None:
        """It has nothing before it, so there is no movement to report."""
        held = excel_reader.list_months(self.book)
        offered = excel_reader.discover_masters(ACCOUNTS, recursive=False)[0].months
        self.assertNotIn(held[0], offered)
        self.assertIn(held[-1], offered)

    def test_the_month_picker_offers_them_all(self) -> None:
        detected = report_engine.detect_available_report_months(ACCOUNTS)
        self.assertGreater(len(detected), 1)
        self.assertEqual(detected, sorted(detected, key=_sort_key, reverse=True),
                         "newest first, so the default lands on a month with data")


def _sort_key(label: str) -> tuple[int, int]:
    parts = label.split()
    return (int(parts[1]), excel_reader.MONTH_NAMES.index(parts[0]) + 1)


@unittest.skipUnless(workbooks(), "demo workbooks not present")
class BothTabsAgreeOnTheSameMonthTests(unittest.TestCase):
    """The property that lets the two sources coexist."""

    def test_the_history_tab_matches_the_rendered_report(self) -> None:
        checked = 0
        for book in workbooks():
            rendered = excel_reader.read_master(book)
            history = excel_reader.read_month(book, rendered.month_label)
            self.assertEqual(history.previous_month_label, rendered.previous_month_label,
                             f"{book.stem}: the two tabs disagree on the previous month")
            self.assertEqual(set(history.metrics), set(rendered.metrics),
                             f"{book.stem}: the tabs produced different metrics")
            for name in sorted(rendered.metrics):
                with self.subTest(dealership=book.stem, metric=name):
                    for field in ("previous_value", "current_value"):
                        self.assertAlmostEqual(
                            float(rendered.metrics[name][field]),
                            float(history.metrics[name][field]), places=2,
                            msg=f"{name}.{field} differs between the report and the history")
                    checked += 1
        self.assertGreater(checked, 100, "the agreement check did not actually run")

    def test_the_latest_month_is_still_read_from_the_rendered_report(self) -> None:
        book = workbooks()[0]
        latest = excel_reader.read_master(book).month_label
        dispatched = excel_reader.read_for_month(book, latest)
        self.assertEqual(dispatched.month_label, latest)
        self.assertEqual(set(dispatched.metrics), set(excel_reader.read_master(book).metrics))


@unittest.skipUnless(workbooks(), "demo workbooks not present")
class EveryOfferedMonthActuallyReportsTests(unittest.TestCase):
    def test_each_offered_month_produces_a_full_set_of_insights(self) -> None:
        offered = report_engine.detect_available_report_months(ACCOUNTS)
        self.assertTrue(offered)
        for label in offered:
            month_name, year = label.split()
            number = excel_reader.MONTH_NAMES.index(month_name) + 1
            # Two dealerships per month is enough to prove the month reports; running
            # all ten across every month turns a test suite into a coffee break.
            groups = core.collect_report_groups(ACCOUNTS, number, int(year))[:2]
            results = [core.process_dealership_group(g, number, int(year), label) for g in groups]
            with self.subTest(month=label):
                produced = [r for r in results if r.report_status == "success"]
                self.assertEqual(len(produced), len(results),
                                 f"{label} is offered but not every dealership reported")
                for result in produced:
                    self.assertGreaterEqual(len(result.parsed_metrics), 15)
                    self.assertEqual(len(result.bullet_insights), core.PRIORITY_INSIGHT_COUNT)

    def test_a_month_outside_the_history_names_the_range_it_covers(self) -> None:
        groups = core.collect_report_groups(ACCOUNTS, 11, 2026)
        result = core.process_dealership_group(groups[0], 11, 2026, "Nov 2026")
        self.assertEqual(result.report_status, "blocked")
        self.assertTrue(any("covers" in note for note in result.notes),
                        f"expected the covered range in the notes, got {result.notes}")

    def test_pipeline_size_reads_the_same_from_either_tab(self) -> None:
        """A snapshot the report prints without comparison stays a snapshot."""
        book = workbooks()[0]
        rendered = excel_reader.read_master(book)
        history = excel_reader.read_month(book, rendered.month_label)
        for source in (rendered, history):
            record = source.metrics.get("sales_pipeline_size")
            self.assertIsNotNone(record)
            self.assertEqual(float(record["previous_value"]), float(record["current_value"]))


if __name__ == "__main__":
    unittest.main()


@unittest.skipUnless(workbooks(), "demo workbooks not present")
class AttachmentsFollowTheMonthTests(unittest.TestCase):
    """A PDF is only ever attached to the month it was exported for.

    The figures come from the workbook now, so a month with no export still reports in
    full -- it just has nothing to attach. What must never happen is the reverse: one
    month's PDF turning up beside another month's figures, which would send a
    dealership an attachment that contradicts the email.
    """

    def test_a_month_with_exports_matches_them(self) -> None:
        groups = core.collect_report_groups(ACCOUNTS, 7, 2026)
        attached = [g for g in groups if g.sales_pdf]
        self.assertTrue(attached, "July has exported PDFs and should match them")

    def test_a_month_without_exports_attaches_nothing(self) -> None:
        for group in core.collect_report_groups(ACCOUNTS, 2, 2026):
            with self.subTest(dealership=group.dealership_name):
                self.assertIsNone(group.sales_pdf)
                self.assertIsNone(group.service_pdf)

    def test_a_month_without_exports_still_reports_in_full(self) -> None:
        groups = core.collect_report_groups(ACCOUNTS, 2, 2026)[:2]
        for group in groups:
            result = core.process_dealership_group(group, 2, 2026, "Feb 2026")
            with self.subTest(dealership=group.dealership_name):
                self.assertEqual(result.report_status, "success")
                self.assertEqual(len(result.bullet_insights), core.PRIORITY_INSIGHT_COUNT)
                self.assertTrue(any("to attach" in note for note in result.notes),
                                f"expected a missing-attachment note, got {result.notes}")

    def test_no_pdf_is_ever_borrowed_from_another_month(self) -> None:
        exported = {p.name for p in ACCOUNTS.rglob("*.pdf")}
        self.assertTrue(exported, "demo data should ship some exported PDFs")
        for month_number, label in ((2, "Feb 2026"), (3, "Mar 2026"), (5, "May 2026")):
            for group in core.collect_report_groups(ACCOUNTS, month_number, 2026):
                with self.subTest(month=label, dealership=group.dealership_name):
                    for pdf in (group.sales_pdf, group.service_pdf):
                        self.assertIsNone(
                            pdf, f"{label} picked up {pdf.name if pdf else ''}")
