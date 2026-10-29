"""Build the Performance Intelligence PDF for a dealership, from its master workbook.

This is the four-page report the operator reviews and the dealership receives: cover,
executive summary, sales and service, marketing. It is drawn by
``templates/performance_intelligence.html``, which is the design as it was authored --
Reporter writes the figures into it and prints the page, rather than re-implementing
the layout here. Editing the template changes the report; nothing in this file needs to
know what it looks like.

The figures come from the **history tab** of the workbook, not from the rendered report
tabs. That matters: the report tabs are pinned to two fixed columns, but the history tab
holds every month, so this report can be produced for any month the workbook covers --
including the year-on-year column and the thirteen-month charts.

Rendering needs a real browser: the layout is CSS grid and the charts are SVG built in
JavaScript. Chromium is driven through Playwright, one instance reused across a whole
run. Where Playwright is not installed the report is simply unavailable, reported as a
reason rather than an error, and the month still runs.
"""

from __future__ import annotations

import datetime as dt
import logging
import atexit
import re
import sys
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

LOGGER = logging.getLogger("reporter.performance")

TEMPLATE_PATH = Path(__file__).resolve().parent / "templates" / "performance_intelligence.html"

# Rows the template reads by position rather than by label, because the labels repeat
# down the sheet. Kept here so a workbook whose layout shifts fails loudly in one place.
ROW_LEAD_SOURCES = 3      # Other, Internet, Phone Enq, Walk-in
ROW_DEAL_SOURCES = 9
ROW_HISTORY_LEADS = 7
ROW_HISTORY_DEALS = 13
ROW_SALES_CAMPAIGNS = 48  # name, deals, ROI in threes
ROW_SERVICE_CAMPAIGNS = 67
PIPELINE_ROWS = (("Unqualified", 23), ("Contacted", 24), ("Store Visit", 25),
                 ("Demo", 26), ("Write-Up", 27), ("F&I", 28))
SOURCE_NAMES = ("Other lead Types", "Internet", "Phone Enq", "Walk-in")

HISTORY_MONTHS = 13
RENDER_TIMEOUT_MS = 60_000


@dataclass
class BuildResult:
    pdf: Path | None = None
    reused: bool = False
    reason: str = ""
    month_label: str = ""
    warnings: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.pdf is not None


# --- reading the workbook -------------------------------------------------------------

def _norm(value: Any) -> str:
    text = str(value if value is not None else "").replace("–", "-").replace("—", "-")
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def _num(value: Any) -> float:
    """A cell as a number. Anything unparseable is zero, as the template expects."""
    if isinstance(value, bool):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value or "").replace(",", "").strip()
    if not text or re.fullmatch(r"n/?a|na|null|-", text.lower()):
        return 0.0
    # Durations are typed as text in these sheets ("50:00 min", "16:35 hrs"). The report
    # shows response time in hours, so they are converted rather than read as zero.
    duration = re.fullmatch(r"(\d+):(\d{2})\s*(min|mins|hr|hrs|hour|hours)?", text, re.I)
    if duration:
        first, second = int(duration.group(1)), int(duration.group(2))
        unit = (duration.group(3) or "min").lower()
        minutes = first * 60 + second if unit.startswith("hr") or unit.startswith("hour") else first + second / 60
        return round(minutes / 60, 4)
    try:
        return float(text)
    except ValueError:
        return 0.0


def _month_label(moment: dt.date) -> str:
    return f"{moment.strftime('%b')} {moment.year}"


def _as_month(value: Any) -> dt.date | None:
    if isinstance(value, dt.datetime):
        return dt.date(value.year, value.month, 1)
    if isinstance(value, dt.date):
        return dt.date(value.year, value.month, 1)
    return None


def _shift(moment: dt.date, months_back: int) -> dt.date:
    total = moment.year * 12 + moment.month - 1 - months_back
    return dt.date(total // 12, total % 12 + 1, 1)


def build_state(workbook: str | Path, month_label: str = "", *, fallback_name: str = "") -> dict[str, Any]:
    """Everything the template needs, for one dealership and one month.

    ``month_label`` selects which month the report covers; left empty, the workbook's
    latest is used. The three comparison periods -- prior month, two months prior and
    the same month a year earlier -- are derived from it, and any that the workbook does
    not cover come back as zeros rather than shifting the others along.
    """
    path = Path(workbook)
    book = load_workbook(path, data_only=True)
    try:
        sheet_name = "Data" if "Data" in book.sheetnames else book.sheetnames[2]
        rows = [list(r) for r in book[sheet_name].iter_rows(values_only=True)]
        if len(rows) < 2:
            raise ValueError(f"{path.name} has no history tab to read.")

        # Month columns sit in row 2. A month can appear twice -- once in the running
        # year and again in the pair the report tabs render -- and the rightmost is the
        # one those tabs used, so it wins.
        columns: dict[dt.date, int] = {}
        order: list[dt.date] = []
        for column, value in enumerate(rows[1]):
            moment = _as_month(value)
            if moment is None:
                continue
            if moment not in columns:
                order.append(moment)
            columns[moment] = column
        if not columns:
            raise ValueError(f"{path.name}: no month columns in row 2 of '{sheet_name}'.")

        months = sorted(columns)
        target = None
        if month_label:
            target = next((m for m in months if _month_label(m) == month_label), None)
        if target is None:
            target = months[-1]

        wanted = [target, _shift(target, 1), _shift(target, 2),
                  dt.date(target.year - 1, target.month, 1)]
        col = [columns.get(w) for w in wanted]

        labels = [_norm(r[1]) if len(r) > 1 else "" for r in rows]

        def row_index(name: str, start: int = 0) -> int | None:
            wanted_label = _norm(name)
            for i in range(start, len(labels)):
                if labels[i] == wanted_label:
                    return i
            return None

        def cell(row: int | None, column: int | None) -> Any:
            if row is None or column is None or row >= len(rows) or column >= len(rows[row]):
                return None
            return rows[row][column]

        def vals(name: str, start: int = 0) -> list[float]:
            index = row_index(name, start)
            return [_num(cell(index, c)) for c in col]

        def vals_any(names: tuple[str, ...], start: int = 0) -> list[float]:
            for name in names:
                if row_index(name, start) is not None:
                    return vals(name, start)
            return [0.0, 0.0, 0.0, 0.0]

        def campaigns(base_row: int) -> tuple[list[str], list[float], list[float]]:
            names, deals, rois = [], [], []
            for slot in range(3):
                row = base_row + slot * 3
                raw = str(cell(row, col[0]) or "").strip()
                names.append(re.sub(r"\s+", " ", raw))
                deals.append(_num(cell(row + 1, col[0])))
                rois.append(_num(cell(row + 2, col[0])))
            return names, deals, rois

        lead_sources, deal_sources, source_conv = {}, {}, {}
        for offset, name in enumerate(SOURCE_NAMES):
            lead_sources[name] = [_num(cell(ROW_LEAD_SOURCES + offset, c)) for c in col]
            deal_sources[name] = [_num(cell(ROW_DEAL_SOURCES + offset, c)) for c in col]
            source_conv[name] = [
                (deal_sources[name][i] / leads * 100) if leads else 0.0
                for i, leads in enumerate(lead_sources[name])
            ]

        dealer = ""
        if "Dealer Information" in book.sheetnames:
            info = [list(r) for r in book["Dealer Information"].iter_rows(
                min_row=1, max_row=3, max_col=1, values_only=True)]
            if len(info) > 1 and info[1] and info[1][0]:
                candidate = str(info[1][0]).strip()
                if candidate and candidate != "0":
                    dealer = candidate
        dealer = dealer or fallback_name or path.stem

        sales_names, sales_deals, sales_rois = campaigns(ROW_SALES_CAMPAIGNS)
        service_names, service_deals, service_rois = campaigns(ROW_SERVICE_CAMPAIGNS)

        leads = vals("total leads")
        deals = vals("total deals")
        warnings: list[str] = []

        # The template reads deals a second time from the pipeline's own "Deals" stage,
        # which is blank in some months and would then report a dealership as having
        # sold nothing. The stated total is authoritative; the pipeline figure is only
        # used where it agrees that there were deals at all.
        pipeline_deals = vals("deals", 20)
        if any(pipeline_deals) and not any(deals):
            deals = pipeline_deals
        elif any(deals) and pipeline_deals[0] and abs(pipeline_deals[0] - deals[0]) > 0.5:
            warnings.append(
                f"Total deals ({deals[0]:.0f}) and the pipeline's Deals stage "
                f"({pipeline_deals[0]:.0f}) disagree for {_month_label(target)}.")

        state: dict[str, Any] = {
            "dealer": dealer,
            "periods": [_month_label(w) for w in wanted],
            "leads": leads,
            "deals": deals,
            "conv": [(deals[i] / leads[i] * 100) if leads[i] else 0.0 for i in range(4)],
            "reminders": vals("service reminders sent"),
            "enrolled": vals("number of enrolled", 58),
            "tasks": vals("number of overdue service tasks"),
            "overdueTasks": [a + b for a, b in zip(vals("number of overdue sales tasks"),
                                                   vals("number of overdue service tasks"))],
            "bookings": vals("service bookings from campaigns"),
            "appointments": vals("number of appointments"),
            "responseTime": vals("internet lead response time"),
            "testDrives": vals("number of test drives"),
            "emailCapture": [x * 100 for x in vals("email capture rate%")],
            "sameDay": [x * 100 for x in vals("sameday deal %")],
            "unqualified": vals_any(("unqualified",)),
            "contacted": vals_any(("contacted",)),
            "storevisit": vals_any(("store visit",)),
            "demo": vals_any(("demo",)),
            "writeup": vals_any(("write-up", "write up")),
            "fi": vals_any(("f&i",)),
            "pipelineTotal": vals_any(("pipeline total",)),
            "lost": vals_any(("lost",)),
            "salesCamps": sales_names, "salesDeals": sales_deals, "salesROI": sales_rois,
            "serviceCamps": service_names, "serviceDeals": service_deals,
            "serviceROI": service_rois,
            "salesEnrolled": vals("number of enrolled", 44),
            "salesCampaignTotalDeals": vals("number of deals", 44),
            "serviceCampaignTotalROI": vals("roi from service campaigns"),
            "leadSources": lead_sources, "dealSources": deal_sources, "sourceConv": source_conv,
        }

        history = [m for m in months if m <= target][-HISTORY_MONTHS:]
        history_cols = [columns[m] for m in history]
        state["history"] = {
            "labels": [_month_label(m) for m in history],
            "leads": [_num(cell(ROW_HISTORY_LEADS, c)) for c in history_cols],
            "deals": [_num(cell(ROW_HISTORY_DEALS, c)) for c in history_cols],
        }
        state["pipelineHistory"] = {
            "labels": [_month_label(m) for m in history],
            "series": [{"name": name,
                        "values": [_num(cell(row, c)) for c in history_cols],
                        "cls": f"p{index + 1}"}
                       for index, (name, row) in enumerate(PIPELINE_ROWS)],
        }

        if col[3] is None:
            warnings.append(
                f"No {_month_label(wanted[3])} in this workbook, so the year-on-year "
                "column is blank.")
        if len(history) < HISTORY_MONTHS:
            warnings.append(
                f"Only {len(history)} of {HISTORY_MONTHS} months available, so the "
                "history charts are shorter than usual.")

        state["_meta"] = {
            "month_label": _month_label(target),
            "months_available": [_month_label(m) for m in months],
            "warnings": warnings,
        }
        return state
    finally:
        book.close()


# --- rendering --------------------------------------------------------------------------

_BROWSER_LOCK = threading.Lock()
_PLAYWRIGHT = None
_BROWSER = None
# Playwright's sync objects only work on the thread that started them.
_OWNER_THREAD: int | None = None


# Installing the package and downloading the browser are two separate steps, and only
# the second one makes rendering possible. Checked separately so Settings never claims
# a report can be produced when it cannot.
_BROWSER_READY = False


def install_command() -> str:
    """The exact command for this machine, naming the interpreter that is running."""
    return f'"{sys.executable}" -m playwright install chromium'


def _playwright_installed() -> bool:
    try:
        import playwright.sync_api  # noqa: F401
    except ImportError:
        return False
    return True


def browser_ready() -> bool:
    """Whether Chromium has actually been downloaded, not just Playwright installed."""
    global _BROWSER_READY
    if _BROWSER_READY:
        return True
    if not _playwright_installed():
        return False
    with _BROWSER_LOCK:
        if _BROWSER is not None and _BROWSER.is_connected():
            _BROWSER_READY = True
            return True
    # Probing means launching: whether Edge or Chrome can be driven is not something a
    # file on disk answers. It costs about a second, once.
    try:
        from playwright.sync_api import sync_playwright

        with sync_playwright() as instance:
            browser = _launch(instance)
            browser.close()
    except Exception:  # noqa: BLE001 - treat any failure as "not ready"
        return False
    # Only the positive answer is remembered: someone installing a browser while
    # Reporter is open should not have to restart it.
    _BROWSER_READY = True
    return True


def available() -> bool:
    return TEMPLATE_PATH.is_file() and _playwright_installed() and browser_ready()


def describe() -> str:
    """One line for the operator about whether this report can be produced here."""
    if not TEMPLATE_PATH.is_file():
        return f"unavailable - the template is missing ({TEMPLATE_PATH.name})"
    if not _playwright_installed():
        return (f'unavailable - Playwright is not installed. Run: "{sys.executable}" '
                f"-m pip install playwright, then {install_command()}")
    if not browser_ready():
        return ("unavailable - no browser to render with. Either install Microsoft Edge "
                f"or Google Chrome, or download Chromium with: {install_command()}")
    return f"{_LAUNCH_CHOICE or 'Chromium'} (via Playwright)"


# Three ways to get a browser, in the order they are tried. The bundled Chromium is a
# 150MB download that a managed network may block outright; Edge is on every Windows
# build and Chrome on most, and Playwright can drive either without downloading
# anything. Preferring the download keeps every machine on the same renderer, but not
# at the cost of the feature being unavailable.
LAUNCH_OPTIONS = (
    ("bundled Chromium", {}),
    ("Microsoft Edge", {"channel": "msedge"}),
    ("Google Chrome", {"channel": "chrome"}),
)
LAUNCH_ARGS = ["--no-sandbox", "--disable-dev-shm-usage"]

_LAUNCH_CHOICE = ""


def _launch(instance):
    """Launch whichever browser this machine actually has."""
    global _LAUNCH_CHOICE
    order = list(LAUNCH_OPTIONS)
    if _LAUNCH_CHOICE:
        order.sort(key=lambda option: option[0] != _LAUNCH_CHOICE)
    failure = None
    for name, options in order:
        try:
            browser = instance.chromium.launch(args=LAUNCH_ARGS, **options)
        except Exception as exc:  # noqa: BLE001 - try the next one
            failure = exc
            continue
        if name != _LAUNCH_CHOICE:
            LOGGER.info("Rendering reports with %s", name)
        _LAUNCH_CHOICE = name
        return browser
    raise failure if failure is not None else RuntimeError("No browser could be launched.")


def _browser():
    """One browser for the whole run. Launching costs about a second; rendering does not."""
    global _PLAYWRIGHT, _BROWSER, _OWNER_THREAD
    if _OWNER_THREAD is not None and _OWNER_THREAD != threading.get_ident():
        # Started by an earlier run's worker thread; unusable here and cannot be closed from here.
        _PLAYWRIGHT = None
        _BROWSER = None
        _OWNER_THREAD = None
    if _BROWSER is not None and _BROWSER.is_connected():
        return _BROWSER
    from playwright.sync_api import sync_playwright

    if _PLAYWRIGHT is None:
        _PLAYWRIGHT = sync_playwright().start()
        _OWNER_THREAD = threading.get_ident()
    _BROWSER = _launch(_PLAYWRIGHT)
    return _BROWSER


def shutdown() -> None:
    """Close the shared browser. Safe to call when one was never started."""
    global _PLAYWRIGHT, _BROWSER, _OWNER_THREAD
    with _BROWSER_LOCK:
        try:
            if _BROWSER is not None:
                _BROWSER.close()
        except Exception:  # noqa: BLE001
            pass
        try:
            if _PLAYWRIGHT is not None:
                _PLAYWRIGHT.stop()
        except Exception:  # noqa: BLE001
            pass
        _BROWSER = None
        _PLAYWRIGHT = None
        _OWNER_THREAD = None


atexit.register(shutdown)


def render_html_file(html_path: str | Path, pdf_path: str | Path) -> bool:
    """Print a standalone HTML file to PDF with the shared browser. Never raises."""
    if not (_playwright_installed() and browser_ready()):
        return False
    try:
        with _BROWSER_LOCK:
            page = _browser().new_page()
            try:
                page.set_default_timeout(RENDER_TIMEOUT_MS)
                page.goto(Path(html_path).resolve().as_uri(), wait_until="load")
                page.pdf(path=str(pdf_path), format="A4", print_background=True)
            finally:
                page.close()
        return True
    except Exception as exc:  # noqa: BLE001
        LOGGER.warning("Could not render %s to PDF: %s", Path(html_path).name, exc)
        return False


def _fingerprint(workbook: Path, month_label: str) -> str:
    stat = workbook.stat()
    template = TEMPLATE_PATH.stat()
    return f"{int(stat.st_mtime)}-{stat.st_size}-{month_label}-{int(template.st_mtime)}"


def _safe_stem(text: str) -> str:
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "-", text).strip(" .")[:120] or "report"


def build(
    workbook: str | Path,
    out_dir: str | Path,
    dealership: str,
    month_label: str = "",
    *,
    force: bool = False,
) -> BuildResult:
    """Produce the report, or say why it could not be produced. Never raises."""
    workbook = Path(workbook)
    out_dir = Path(out_dir)
    if not workbook.is_file():
        return BuildResult(reason=f"{workbook.name} is not there to read.")
    if not TEMPLATE_PATH.is_file():
        return BuildResult(reason=f"The report template is missing ({TEMPLATE_PATH}).")

    try:
        state = build_state(workbook, month_label, fallback_name=dealership)
    except Exception as exc:  # noqa: BLE001 - a month must not fail over its report
        LOGGER.exception("Could not read %s for the performance report", workbook.name)
        return BuildResult(reason=f"Could not read {workbook.name}: {exc}")

    meta = state.pop("_meta", {})
    resolved_month = str(meta.get("month_label") or month_label)
    warnings = list(meta.get("warnings") or [])

    if month_label and resolved_month != month_label:
        return BuildResult(
            reason=(f"This workbook has no figures for {month_label}. "
                    f"It covers {', '.join(meta.get('months_available') or []) or 'nothing'}."),
            month_label=resolved_month)

    # Named per workbook: sites in one group share the same 'Dealer Information' name.
    stem = _safe_stem(f"{dealership or state['dealer']} - Performance Intelligence - {resolved_month}")
    target = out_dir / f"{stem}.pdf"
    stamp = out_dir / f".{stem}.source"
    signature = _fingerprint(workbook, resolved_month)

    if not force and target.is_file() and target.stat().st_size > 0 and stamp.is_file():
        try:
            if stamp.read_text(encoding="utf-8").strip() == signature:
                return BuildResult(pdf=target, reused=True, month_label=resolved_month,
                                   warnings=warnings)
        except OSError:
            pass

    if not available():
        return BuildResult(reason=describe(), month_label=resolved_month, warnings=warnings)

    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        return BuildResult(reason=f"Could not write to {out_dir}: {exc}", month_label=resolved_month)

    try:
        # One page at a time: the template keeps state in module scope, so two renders
        # sharing a page would overwrite each other's figures.
        with _BROWSER_LOCK:
            try:
                page = _browser().new_page()
            except Exception:  # noqa: BLE001 - a stale browser is relaunched, not fatal
                global _BROWSER
                _BROWSER = None
                page = _browser().new_page()
            try:
                page.set_default_timeout(RENDER_TIMEOUT_MS)
                errors: list[str] = []
                page.on("pageerror", lambda exc: errors.append(str(exc)))
                page.goto(TEMPLATE_PATH.resolve().as_uri(), wait_until="load")
                page.evaluate(
                    "data => { state = Object.assign({}, state, data); render(); }", state)
                page.wait_for_timeout(350)
                page.pdf(path=str(target), format="A4", print_background=True,
                         margin={"top": "0", "right": "0", "bottom": "0", "left": "0"})
                if errors:
                    LOGGER.info("Template reported %d script error(s) for %s: %s",
                                len(errors), dealership, errors[0][:160])
            finally:
                page.close()
    except Exception as exc:  # noqa: BLE001
        LOGGER.exception("Rendering the performance report failed for %s", dealership)
        text = str(exc)
        if "Executable doesn't exist" in text or "playwright install" in text:
            global _BROWSER_READY
            _BROWSER_READY = False
            return BuildResult(
                reason=("No browser is available to render the report. Install Microsoft "
                        "Edge or Google Chrome, or download Chromium with: "
                        f"{install_command()}"),
                month_label=resolved_month, warnings=warnings)
        return BuildResult(reason=f"Could not render the report: {text.splitlines()[0][:200]}",
                           month_label=resolved_month, warnings=warnings)

    if not target.is_file() or target.stat().st_size == 0:
        return BuildResult(reason="The renderer produced an empty file.",
                           month_label=resolved_month, warnings=warnings)

    try:
        stamp.write_text(signature, encoding="utf-8")
    except OSError:
        pass  # the cache is an optimisation, not a requirement
    return BuildResult(pdf=target, month_label=resolved_month, warnings=warnings)
