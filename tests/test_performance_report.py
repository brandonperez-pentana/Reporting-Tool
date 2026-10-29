"""The Performance Intelligence report, built from the workbook's history tab.

The report and the email are sent together, so the figures on it have to be the month's
own. These tests pin the parts where that could quietly go wrong: which month is read,
which periods it is compared against, and the two cell formats that silently read as
zero if they are not handled.

Rendering needs Chromium, so those tests skip where it is not installed. Everything
about reading the workbook is checked regardless, because that is where a wrong number
would come from.
"""

from __future__ import annotations

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
import performance_report as pr  # noqa: E402
import settings_manager  # noqa: E402

ACCOUNTS = ROOT / "Accounts"
CAN_RENDER = pr.available()


def workbooks() -> list[Path]:
    return sorted(ACCOUNTS.glob("*.xls[xm]"))


@unittest.skipUnless(workbooks(), "demo workbooks not present")
class ReadingTheWorkbookTests(unittest.TestCase):
    def setUp(self) -> None:
        self.book = workbooks()[0]
        self.state = pr.build_state(self.book, fallback_name="Demo")

    def test_it_defaults_to_the_latest_month(self) -> None:
        meta = self.state["_meta"]
        self.assertEqual(meta["month_label"], meta["months_available"][-1])

    def test_the_four_periods_step_back_correctly(self) -> None:
        """Latest, prior, two prior, then the same month a year earlier."""
        periods = self.state["periods"]
        self.assertEqual(len(periods), 4)
        month, year = periods[0].split()
        self.assertEqual(periods[3], f"{month} {int(year) - 1}")

    def test_a_named_month_is_honoured(self) -> None:
        available = self.state["_meta"]["months_available"]
        target = available[2]
        state = pr.build_state(self.book, target)
        self.assertEqual(state["periods"][0], target)
        self.assertEqual(state["periods"][1], available[1])

    def test_history_never_runs_past_the_month_being_reported(self) -> None:
        """A March report must not chart April; the reader is asked for one month."""
        available = self.state["_meta"]["months_available"]
        target = available[2]
        state = pr.build_state(self.book, target)
        self.assertEqual(state["history"]["labels"][-1], target)
        self.assertLessEqual(len(state["history"]["labels"]), pr.HISTORY_MONTHS)

    def test_response_time_is_converted_from_its_text_form(self) -> None:
        """Stored as "16:35 hrs"; a plain number parse would report zero."""
        self.assertGreater(self.state["responseTime"][0], 0,
                           "response time should not read as zero")
        self.assertEqual(pr._num("16:35 hrs"), 16.5833)
        self.assertAlmostEqual(pr._num("50:00 min"), 0.8333, places=3)
        self.assertEqual(pr._num("N/A"), 0.0)

    def test_total_deals_survives_a_gap_in_the_pipeline_stage(self) -> None:
        """The pipeline's own Deals row is blank some months; the stated total is not."""
        self.assertGreater(self.state["deals"][0], 0)
        self.assertEqual(self.state["deals"][0], self.state["_meta"] and self.state["deals"][0])
        leads = self.state["leads"][0]
        if leads:
            self.assertAlmostEqual(self.state["conv"][0], self.state["deals"][0] / leads * 100, places=6)

    def test_missing_comparison_periods_are_reported_not_hidden(self) -> None:
        warnings_seen = " ".join(self.state["_meta"]["warnings"])
        if self.state["periods"][3] not in self.state["_meta"]["months_available"]:
            self.assertIn("year-on-year", warnings_seen)

    def test_every_block_the_template_reads_is_populated(self) -> None:
        for key in ("leads", "deals", "conv", "pipelineTotal", "lost", "appointments",
                    "emailCapture", "sameDay", "testDrives"):
            with self.subTest(field=key):
                self.assertEqual(len(self.state[key]), 4)
        self.assertEqual(len(self.state["salesCamps"]), 3)
        self.assertEqual(len(self.state["serviceCamps"]), 3)
        self.assertEqual(len(self.state["pipelineHistory"]["series"]), len(pr.PIPELINE_ROWS))
        for source in pr.SOURCE_NAMES:
            self.assertIn(source, self.state["leadSources"])


class TheTemplateIsSelfContainedTests(unittest.TestCase):
    """It has to render on a machine with no internet and no CDN reachable."""

    def test_the_template_ships_with_the_app(self) -> None:
        self.assertTrue(pr.TEMPLATE_PATH.is_file(), f"missing {pr.TEMPLATE_PATH}")

    def test_it_loads_nothing_from_the_network(self) -> None:
        import re

        html = pr.TEMPLATE_PATH.read_text(encoding="utf-8", errors="replace")
        external = set(re.findall(r'(?:src|href)="(https?://[^"]+)"', html))
        self.assertEqual(external, set(), f"template reaches out to {external}")

    def test_it_has_somewhere_to_put_the_figures(self) -> None:
        html = pr.TEMPLATE_PATH.read_text(encoding="utf-8", errors="replace")
        self.assertIn('id="reporter-data"', html)


@unittest.skipUnless(workbooks() and CAN_RENDER, "needs workbooks and Chromium")
class RenderingTests(unittest.TestCase):
    @classmethod
    def tearDownClass(cls) -> None:
        pr.shutdown()

    def setUp(self) -> None:
        self.out = Path(tempfile.mkdtemp())

    def tearDown(self) -> None:
        shutil.rmtree(self.out, ignore_errors=True)

    def test_it_produces_a_four_page_pdf(self) -> None:
        result = pr.build(workbooks()[0], self.out, "Demo Site 1")
        self.assertTrue(result.ok, result.reason)
        data = result.pdf.read_bytes()
        self.assertTrue(data.startswith(b"%PDF"))
        self.assertGreater(len(data), 50_000)
        self.assertEqual(data.count(b"/Type /Page\n") or data.count(b"/Type/Page"), 4)

    def test_the_filename_names_the_dealership_and_month(self) -> None:
        result = pr.build(workbooks()[0], self.out, "Demo Site 1")
        self.assertIn("Performance Intelligence", result.pdf.name)
        self.assertIn(result.month_label, result.pdf.name)

    def test_sites_sharing_a_group_name_get_their_own_file(self) -> None:
        first = pr.build(workbooks()[0], self.out, "Group Site A")
        second = pr.build(workbooks()[0], self.out, "Group Site B")
        self.assertTrue(first.ok and second.ok)
        self.assertNotEqual(first.pdf, second.pdf)
        self.assertTrue(first.pdf.is_file() and second.pdf.is_file())

    def test_a_second_build_is_served_from_cache(self) -> None:
        first = pr.build(workbooks()[0], self.out, "Demo Site 1")
        second = pr.build(workbooks()[0], self.out, "Demo Site 1")
        self.assertFalse(first.reused)
        self.assertTrue(second.reused)

    def test_a_changed_workbook_is_rendered_again(self) -> None:
        copy = self.out / "copy.xlsx"
        shutil.copy2(workbooks()[0], copy)
        self.assertFalse(pr.build(copy, self.out, "Copy").reused)
        self.assertTrue(pr.build(copy, self.out, "Copy").reused)
        copy.touch()
        self.assertFalse(pr.build(copy, self.out, "Copy").reused,
                         "a re-saved workbook must not serve a stale report")

    def test_the_workbook_is_never_modified(self) -> None:
        import hashlib

        book = workbooks()[0]
        before = hashlib.sha256(book.read_bytes()).hexdigest()
        pr.build(book, self.out, "Demo Site 1")
        self.assertEqual(hashlib.sha256(book.read_bytes()).hexdigest(), before)


@unittest.skipUnless(workbooks(), "demo workbooks not present")
class FailingSafelyTests(unittest.TestCase):
    def test_a_month_the_workbook_lacks_is_declined_with_its_range(self) -> None:
        result = pr.build(workbooks()[0], Path(tempfile.mkdtemp()), "Demo", "Nov 2029")
        self.assertFalse(result.ok)
        self.assertIn("no figures for Nov 2029", result.reason)
        self.assertIn("covers", result.reason)

    def test_a_missing_renderer_is_a_reason_not_an_exception(self) -> None:
        with mock.patch.object(pr, "available", return_value=False):
            with mock.patch.object(pr, "describe", return_value="unavailable - no Chromium"):
                result = pr.build(workbooks()[0], Path(tempfile.mkdtemp()), "Demo")
        self.assertFalse(result.ok)
        self.assertIn("Chromium", result.reason)

    def test_a_missing_workbook_is_reported(self) -> None:
        result = pr.build(ROOT / "nope.xlsx", Path(tempfile.mkdtemp()), "Demo")
        self.assertFalse(result.ok)
        self.assertIn("not there", result.reason)

    def test_the_month_still_runs_when_the_report_cannot_be_built(self) -> None:
        group = core.collect_report_groups(ACCOUNTS, 7, 2026)[0]
        base = dict(settings_manager.load_settings(), performance_report="off")
        with mock.patch.object(settings_manager, "load_settings", return_value=base):
            result = core.process_dealership_group(group, 7, 2026, "Jul 2026")
        self.assertEqual(result.report_status, "success")
        self.assertIsNone(result.performance_pdf)

    def test_switching_it_off_never_calls_the_renderer(self) -> None:
        group = core.collect_report_groups(ACCOUNTS, 7, 2026)[0]
        base = dict(settings_manager.load_settings(), performance_report="off")
        with mock.patch.object(settings_manager, "load_settings", return_value=base):
            with mock.patch.object(pr, "build") as build:
                core.build_performance_report(group, "Jul 2026")
        build.assert_not_called()

    def test_a_library_of_pdfs_with_no_workbook_is_skipped(self) -> None:
        group = core.DealershipReportGroup(
            dealership_name="No Workbook", group_key="no workbook",
            folder_path=ACCOUNTS, sales_pdf=None, service_pdf=None)
        pdf, note = core.build_performance_report(group, "Jul 2026")
        self.assertIsNone(pdf)
        self.assertEqual(note, "")


if __name__ == "__main__":
    unittest.main()


class SayingWhatIsActuallyMissingTests(unittest.TestCase):
    """Installing the package and downloading the browser are separate steps.

    Reporting the first as success when the second has not happened is worse than
    useless: Settings says the report is available, the run then fails, and the operator
    is shown several lines of the renderer's own box-drawn error instead of the one
    command that fixes it.
    """

    BROWSER_MISSING = (
        "BrowserType.launch: Executable doesn't exist at "
        "C:\\Users\\x\\AppData\\Local\\ms-playwright\\chromium_headless_shell-1243\\"
        "chrome-headless-shell.exe\n"
        "\u2554\u2550\u2550\u2550\u2557\n"
        "\u2551 Looks like Playwright was just installed or updated. \u2551\n"
        "\u2551     playwright install                                \u2551\n"
        "\u255a\u2550\u2550\u2550\u255d")

    def test_a_missing_browser_is_not_reported_as_available(self) -> None:
        with mock.patch.object(pr, "browser_ready", return_value=False):
            self.assertFalse(pr.available())
            self.assertIn("no browser to render with", pr.describe())

    def test_a_missing_package_is_distinguished_from_a_missing_browser(self) -> None:
        with mock.patch.object(pr, "_playwright_installed", return_value=False):
            self.assertIn("Playwright is not installed", pr.describe())

    def test_the_command_names_the_interpreter_that_is_running(self) -> None:
        """Installing into a different Python is the usual reason this stays broken."""
        self.assertIn(sys.executable, pr.install_command())
        self.assertIn("playwright install chromium", pr.install_command())

    @unittest.skipUnless(workbooks(), "demo workbooks not present")
    def test_the_renderer_error_becomes_one_actionable_line(self) -> None:
        with mock.patch.object(pr, "_browser", side_effect=Exception(self.BROWSER_MISSING)):
            result = pr.build(workbooks()[0], Path(tempfile.mkdtemp()), "Demo", "Jul 2026")
        self.assertFalse(result.ok)
        self.assertEqual(len(result.reason.splitlines()), 1, result.reason)
        self.assertFalse(any(ch in result.reason for ch in "\u2554\u2550\u2551\u255a"),
                         "the renderer's box drawing reached the operator")
        self.assertIn("playwright install chromium", result.reason)

    @unittest.skipUnless(workbooks(), "demo workbooks not present")
    def test_an_unrelated_render_failure_is_kept_short(self) -> None:
        noisy = "Something broke\n" + "\n".join(f"  frame {i}" for i in range(40))
        with mock.patch.object(pr, "_browser", side_effect=Exception(noisy)):
            result = pr.build(workbooks()[0], Path(tempfile.mkdtemp()), "Demo", "Jul 2026")
        self.assertFalse(result.ok)
        self.assertEqual(len(result.reason.splitlines()), 1)
        self.assertIn("Something broke", result.reason)


class WhicheverBrowserTheMachineHasTests(unittest.TestCase):
    """A managed build may block the 150MB Chromium download outright.

    Edge ships with Windows and Chrome is on most machines, and Playwright can drive
    either without downloading anything, so the report should not be unavailable
    merely because a CDN is unreachable.
    """

    class _Chromium:
        def __init__(self, works: set[str]) -> None:
            self.works, self.tried = works, []

        def launch(self, args=None, **options):
            channel = options.get("channel", "bundled")
            self.tried.append(channel)
            if channel not in self.works:
                raise RuntimeError(f"Executable doesn't exist ({channel})")
            return mock.MagicMock()

    class _Instance:
        def __init__(self, works: set[str]) -> None:
            self.chromium = WhicheverBrowserTheMachineHasTests._Chromium(works)

    def setUp(self) -> None:
        self._saved = pr._LAUNCH_CHOICE
        pr._LAUNCH_CHOICE = ""

    def tearDown(self) -> None:
        pr._LAUNCH_CHOICE = self._saved

    def test_the_download_is_preferred_when_it_is_there(self) -> None:
        """Same renderer on every machine, where the machine allows it."""
        instance = self._Instance({"bundled"})
        pr._launch(instance)
        self.assertEqual(instance.chromium.tried, ["bundled"])
        self.assertEqual(pr._LAUNCH_CHOICE, "bundled Chromium")

    def test_edge_is_used_when_there_is_no_download(self) -> None:
        instance = self._Instance({"msedge"})
        pr._launch(instance)
        self.assertEqual(instance.chromium.tried, ["bundled", "msedge"])
        self.assertEqual(pr._LAUNCH_CHOICE, "Microsoft Edge")

    def test_chrome_is_the_last_resort(self) -> None:
        instance = self._Instance({"chrome"})
        pr._launch(instance)
        self.assertEqual(pr._LAUNCH_CHOICE, "Google Chrome")

    def test_a_machine_with_none_of_them_raises_for_the_caller_to_report(self) -> None:
        instance = self._Instance(set())
        with self.assertRaises(Exception):
            pr._launch(instance)
        self.assertEqual(instance.chromium.tried, ["bundled", "msedge", "chrome"])

    def test_the_winner_is_tried_first_next_time(self) -> None:
        """Ten dealerships must not each re-probe a browser that is not installed."""
        pr._LAUNCH_CHOICE = "Microsoft Edge"
        instance = self._Instance({"msedge"})
        pr._launch(instance)
        self.assertEqual(instance.chromium.tried, ["msedge"])

    def test_the_unavailable_message_offers_both_ways_out(self) -> None:
        with mock.patch.object(pr, "browser_ready", return_value=False):
            message = pr.describe()
        self.assertIn("Edge", message)
        self.assertIn("playwright install chromium", message)
