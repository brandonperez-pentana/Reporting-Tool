"""Read a month out of a DealerSocket master workbook.

This replaces the PDF path. The old route was PDF -> pdfplumber text -> regex, which
reconstructed numbers from prose that Excel had already laid out in cells; every
fragile step in ingestion lived in that reconstruction. Here the figures are read
where they were typed.

Only the first two tabs are read, whatever they are called. That is the operator's
own guarantee about these workbooks -- tab 1 is the sales report and tab 2 is the
service report -- and it is a stronger contract than sheet names, which vary.

Nothing is written back. A master is opened read-only, from bytes, and closed. The
files in SharePoint are shared originals and this module must never be the reason one
changes.

What comes out is deliberately the same shape the rest of Reporter already consumes:
metric records built by ``dealership_insights_local.build_metric_record``, so unit
detection, significance banding and reliability rules stay in one place and the
engine cannot tell which source it was handed.

The layout is read structurally rather than by cell address. Each block in these
reports is a label with two dated rows under it -- previous month, then current --
and the reader finds those pairs wherever they sit. A master with a row inserted
still reads correctly, which a fixed-address reader would not.
"""

from __future__ import annotations

import datetime as dt
import html
import io
import logging
import re
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook

LOGGER = logging.getLogger("reporter.excel")

MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# --- what a label in the workbook means to the engine ------------------------------
#
# Keyed by the block label as it is written in the sheet, normalised. The value is
# either an engine metric name, or a dict mapping each sub-column to one. Anything not
# listed is read and carried as an extra, not discarded: an unmapped label is a
# workbook change worth seeing, not an error worth stopping for.

METRIC_MAP: dict[str, Any] = {
    # Sales report, performance metrics
    "UNCONTACTED LEADS": "sales_uncontacted_leads",
    "INTERNET LEAD RESPONSE TIME": "sales_internet_response_time",
    "NEW LEADS TO APPOINTMENTS": "sales_new_leads_to_appointments",
    "NEW LEADS TO TEST DRIVES": "sales_new_leads_to_test_drives",
    "APPOINTMENT TO SALES": "sales_appointment_to_sales",
    "OVERDUE TASKS": "sales_overdue_tasks",
    "E-MAIL CAPTURE RATE": "sales_email_capture_rate",
    "SAME DAY DEAL %": "sales_same_day_deal_pct",
    "CURRENT PIPELINE SIZE": "sales_pipeline_size",
    # Sales report, targeted marketing: one block, two value columns
    "DEALS FROM CAMPAIGNS": {"DEALS": "sales_campaign_deals", "ROI": "sales_campaign_roi"},
    # Service report
    "SERVICE REMINDERS SENT": "service_reminders_sent",
    "CUSTOMERS ENROLLED": "service_customers_enrolled",
    "OUTSTANDING TASKS": "service_outstanding_tasks",
    "NO SHOWS": "service_no_shows",
    "REBOOKED NO SHOWS FROM CAMPAIGNS": "service_rebooked_no_shows",
    "SERVICE BOOKINGS FROM CAMPAIGNS": {
        "RO'S": "service_bookings_from_campaigns",
        "ROI": "service_campaign_roi",
    },
}

# The leads/deals/conversion matrix at the top of the sales report. Row label crossed
# with column group gives the metric.
MATRIX_TOTAL_MAP = {
    ("TOTAL", "NUMBER OF LEADS"): "sales_total_leads",
    ("TOTAL", "NUMBER OF DEALS"): "sales_total_deals",
    ("TOTAL", "CONVERSION RATIOS"): "sales_overall_conversion_rate",
}


def _normalise(text: Any) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip().upper()


def _month_label(value: dt.datetime | dt.date | Any) -> str:
    """A date cell to the 'Jul 2026' label the rest of Reporter speaks."""
    if isinstance(value, (dt.datetime, dt.date)):
        return f"{MONTH_NAMES[value.month - 1]} {value.year}"
    text = str(value or "").strip()
    match = re.search(r"([A-Za-z]{3,9})[-\s]+(\d{2,4})", text)
    if match:
        month = match.group(1)[:3].title()
        year = match.group(2)
        if len(year) == 2:
            year = f"20{year}"
        return f"{month} {year}"
    return text


def _is_date(value: Any) -> bool:
    return isinstance(value, (dt.datetime, dt.date))


def _decimals(number_format: str) -> int:
    """How many decimal places the workbook itself shows for this cell.

    Reporter renders the same figure twice: once in the email copy and once in the PDF
    attached beside it. Both have to agree, so the reader rounds the way the sheet
    does rather than to its own taste.
    """
    match = re.search(r"0\.(0+)", str(number_format or ""))
    return len(match.group(1)) if match else 0


def _raw_value(value: Any, number_format: str) -> str:
    """Render a cell the way ``parse_metric_value`` expects to be handed it.

    The number format is what says whether 0.53 is fifty-three percent or fifty-three
    cents, so it is read rather than guessed. Getting this wrong is silent: the engine
    would happily report a conversion rate of 0.53%.
    """
    if value is None or value == "":
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, (dt.datetime, dt.date)):
        return _month_label(value)
    if not isinstance(value, (int, float)):
        return str(value)

    fmt = str(number_format or "")
    if "%" in fmt:
        return f"{float(value) * 100:.{_decimals(fmt)}f}%"
    if "$" in fmt or "£" in fmt or "€" in fmt:
        return f"${float(value):,.0f}"
    number = float(value)
    if abs(number - round(number)) < 1e-9:
        return str(int(round(number)))
    return f"{number:.2f}"


def _looks_like_value(value: Any) -> bool:
    """Is this a figure, or a label that happens to sit next to one?

    Response times arrive as text ('16:35 hrs'), so 'is it a number' is too narrow;
    campaign names are text too, so 'is it non-empty' is too wide.
    """
    if isinstance(value, (int, float)):
        return True
    if not isinstance(value, str):
        return False
    text = value.strip()
    if not text:
        return False
    if text.lower() in {"n/a", "na", "-"}:
        return True
    return bool(re.match(r"^[\$\-+]?[\d,]+(\.\d+)?\s*(%|min|mins|hr|hrs|sec|secs)?$", text, re.I)
                or re.match(r"^\d{1,3}:\d{2}\s*(min|mins|hr|hrs)?$", text, re.I))


@dataclass
class Series:
    """One labelled figure with its previous and current month values."""

    block: str
    column: str
    previous_raw: str
    current_raw: str
    previous_month: str
    current_month: str


@dataclass
class WorkbookRead:
    dealership: str
    month_label: str
    previous_month_label: str
    metrics: dict[str, dict[str, Any]] = field(default_factory=dict)
    breakdowns: dict[str, dict[str, Any]] = field(default_factory=dict)
    campaigns: dict[str, list[dict[str, Any]]] = field(default_factory=dict)
    # Pipeline stages, when they were read from the history tab. Empty on a read of the
    # rendered report, where build_breakdowns goes and fetches them itself.
    pipeline: dict[str, dict[str, Any]] = field(default_factory=dict)
    unmapped: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def is_usable(self) -> bool:
        return bool(self.metrics) and bool(self.month_label)


# --- reading the sheet structurally -------------------------------------------------

def _grid(worksheet) -> tuple[list[list[Any]], list[list[str]]]:
    values: list[list[Any]] = []
    formats: list[list[str]] = []
    for row in worksheet.iter_rows():
        values.append([cell.value for cell in row])
        formats.append([cell.number_format for cell in row])
    return values, formats


def _cell(grid: list[list[Any]], row: int, col: int) -> Any:
    if 0 <= row < len(grid) and 0 <= col < len(grid[row]):
        return grid[row][col]
    return None


def _label_above(values: list[list[Any]], row: int, col: int, limit: int = 6) -> str:
    """The nearest text cell above, which is how these blocks are titled."""
    for offset in range(1, limit + 1):
        candidate = _cell(values, row - offset, col)
        if isinstance(candidate, str) and candidate.strip():
            return _normalise(candidate)
    return ""


def _paired_blocks(values: list[list[Any]], formats: list[list[str]]) -> list[Series]:
    """Find every 'label, then two dated rows' block on a sheet.

    A block is recognised by its shape rather than its position: a date in some column
    with another date directly below it, and figures to the right of both.
    """
    series: list[Series] = []
    seen: set[tuple[int, int]] = set()

    for row_index, row in enumerate(values):
        for col_index, value in enumerate(row):
            if not _is_date(value) or (row_index, col_index) in seen:
                continue
            below = _cell(values, row_index + 1, col_index)
            if not _is_date(below):
                continue
            seen.add((row_index, col_index))
            seen.add((row_index + 1, col_index))

            block = _label_above(values, row_index, col_index)
            previous_month = _month_label(value)
            current_month = _month_label(below)

            # Walk right while either dated row still holds a figure. The stop rule
            # matters: these blocks sit beside a list of campaign names, and without
            # it the reader walks straight into that list and reads 'GWM Closed RO'
            # as if it were this block's third value column.
            offset = 1
            while True:
                top = _cell(values, row_index, col_index + offset)
                bottom = _cell(values, row_index + 1, col_index + offset)
                if (top is None or top == "") and (bottom is None or bottom == ""):
                    break
                if _is_date(top) or _is_date(bottom):
                    break
                if not (_looks_like_value(top) or _looks_like_value(bottom)):
                    break
                # A sub-column heading sits directly above the first dated row, or
                # there is none. Looking further up finds the block above instead.
                heading = _cell(values, row_index - 1, col_index + offset)
                column_label = _normalise(heading) if isinstance(heading, str) else ""
                if column_label == block:
                    column_label = ""
                series.append(Series(
                    block=block,
                    column=column_label,
                    previous_raw=_raw_value(top, _cell(formats, row_index, col_index + offset)),
                    current_raw=_raw_value(bottom, _cell(formats, row_index + 1, col_index + offset)),
                    previous_month=previous_month,
                    current_month=current_month,
                ))
                offset += 1
                if offset > 4:
                    break
    return series


def _matrix_block(values: list[list[Any]], formats: list[list[str]]) -> tuple[list[Series], dict[str, dict[str, Any]]]:
    """The leads / deals / conversion table, which is dated across a row, not down one.

    This is also where the lead-source split comes from, which the funnel engine needs
    and which the PDF path had to recover from prose.
    """
    series: list[Series] = []
    breakdowns: dict[str, dict[str, Any]] = {}

    for row_index, row in enumerate(values):
        date_columns = [i for i, v in enumerate(row) if _is_date(v)]
        if len(date_columns) < 4:
            continue
        # Groups of two: previous and current under each heading.
        pairs = [(date_columns[i], date_columns[i + 1]) for i in range(0, len(date_columns) - 1, 2)]
        headings: list[str] = []
        for previous_col, _ in pairs:
            headings.append(_label_above(values, row_index, previous_col, limit=2))

        for below in range(row_index + 1, min(row_index + 12, len(values))):
            label = None
            for col in range(0, 4):
                candidate = _cell(values, below, col)
                if isinstance(candidate, str) and candidate.strip():
                    label = _normalise(candidate)
                    break
            if not label:
                break
            for (previous_col, current_col), heading in zip(pairs, headings):
                previous_value = _cell(values, below, previous_col)
                current_value = _cell(values, below, current_col)
                if previous_value is None and current_value is None:
                    continue
                series.append(Series(
                    block=heading,
                    column=label,
                    previous_raw=_raw_value(previous_value, _cell(formats, below, previous_col)),
                    current_raw=_raw_value(current_value, _cell(formats, below, current_col)),
                    previous_month=_month_label(row[previous_col]),
                    current_month=_month_label(row[current_col]),
                ))
                breakdowns.setdefault(heading, {})[label] = {
                    "previous": previous_value,
                    "current": current_value,
                }
        break  # only the first such table on a sheet
    return series, breakdowns


def _single_figures(values: list[list[Any]], formats: list[list[str]]) -> list[Series]:
    """Figures stated once rather than as a dated pair.

    Current pipeline size is a snapshot, not a month-on-month movement, so it appears
    as a single labelled number. The engine still wants a record, so the one value is
    used for both sides and the movement comes out flat -- which is honest: the sheet
    does not say what it was last month.
    """
    found: list[Series] = []
    for row_index, row in enumerate(values):
        for col_index, value in enumerate(row):
            if not isinstance(value, str):
                continue
            if "PIPELINE SIZE" not in _normalise(value):
                continue
            for offset in range(1, 4):
                figure = _cell(values, row_index, col_index + offset)
                if isinstance(figure, (int, float)):
                    raw = _raw_value(figure, _cell(formats, row_index, col_index + offset))
                    found.append(Series(
                        block="CURRENT PIPELINE SIZE", column="",
                        previous_raw=raw, current_raw=raw,
                        previous_month="", current_month="",
                    ))
                    break
    return found


def _campaign_rows(values: list[list[Any]]) -> list[dict[str, Any]]:
    """Named campaigns with their deals and return, listed beside the dated blocks."""
    campaigns: list[dict[str, Any]] = []
    for row_index, row in enumerate(values):
        for col_index, value in enumerate(row):
            if not isinstance(value, str) or not value.strip():
                continue
            if _normalise(value) not in {"SUCCESSFUL CAMPAIGNS", "ADDITIONAL SUCCESSFUL CAMPAIGNS"}:
                continue
            for below in range(row_index + 1, min(row_index + 8, len(values))):
                name = _cell(values, below, col_index)
                if not isinstance(name, str) or not name.strip():
                    continue
                if _normalise(name) in {"N/A", "NA"}:
                    continue
                numbers = [v for v in values[below][col_index + 1:col_index + 6]
                           if isinstance(v, (int, float))]
                if not numbers:
                    continue
                campaigns.append({
                    "name": name.strip(),
                    "deals": numbers[0] if numbers else None,
                    "roi": numbers[1] if len(numbers) > 1 else None,
                })
    return campaigns


# First text boxes that are a section heading or template placeholder, not a dealership.
_GENERIC_TITLES = {"MONTHLY REPORT", "DEMO DEALERSHIP", "DEALER NAME", "DEALERSHIP NAME",
                   "SALES", "SERVICE"}


def _title_from_drawings(source: bytes | Path) -> tuple[str, str]:
    """Dealership and month, which live in the report's text boxes rather than cells.

    openpyxl cannot see drawing text, so the drawing XML is read straight out of the
    package. Falls back to empty strings, and the caller falls back to the filename --
    a workbook whose title box has been restyled must not stop a month.
    """
    try:
        handle = io.BytesIO(source) if isinstance(source, bytes) else source
        with zipfile.ZipFile(handle) as package:
            names = sorted(n for n in package.namelist()
                           if n.startswith("xl/drawings/drawing") and n.endswith(".xml"))
            for name in names[:2]:
                text = package.read(name).decode("utf-8", "ignore")
                runs = [html.unescape(t).strip()
                        for t in re.findall(r"<a:t>([^<]*)</a:t>", text) if t.strip()]
                if not runs:
                    continue
                dealership = runs[0]
                month = ""
                for run in runs[1:5]:
                    if re.search(r"[A-Za-z]{3,9}\s+\d{4}", run):
                        month = run
                        break
                if dealership and _normalise(dealership) not in _GENERIC_TITLES:
                    return dealership, month
    except Exception as exc:  # noqa: BLE001 - never let cosmetics stop a read
        LOGGER.info("Could not read the title from drawings: %s", exc)
    return "", ""


# --- the public entry point ---------------------------------------------------------

def read_master(source: bytes | str | Path, *, fallback_name: str = "") -> WorkbookRead:
    """Read one master workbook into metrics the engine already understands."""
    from dealership_insights_local import build_metric_record

    handle = io.BytesIO(source) if isinstance(source, bytes) else source
    workbook = load_workbook(handle, data_only=True, read_only=False)
    try:
        sheets = workbook.sheetnames[:2]
        if not sheets:
            raise ValueError("The workbook has no sheets.")

        all_series: list[Series] = []
        breakdowns: dict[str, dict[str, Any]] = {}
        campaigns: dict[str, list[dict[str, Any]]] = {}

        for position, sheet_name in enumerate(sheets):
            worksheet = workbook[sheet_name]
            values, formats = _grid(worksheet)
            matrix_series, matrix_breakdowns = _matrix_block(values, formats)
            all_series.extend(matrix_series)
            all_series.extend(_paired_blocks(values, formats))
            all_series.extend(_single_figures(values, formats))
            breakdowns.update(matrix_breakdowns)
            found = _campaign_rows(values)
            if found:
                campaigns["sales" if position == 0 else "service"] = found

        dealership, drawing_month = _title_from_drawings(source)
        dealership = dealership or fallback_name

        # Tab 3 is authoritative for which month this workbook holds: its rightmost
        # populated column is the month, whatever the filename or the title box says.
        # The dated pairs on tabs 1 and 2 are the fallback, and the title box is the
        # last resort.
        try:
            previous_label, current_label = read_months(source)
        except Exception as exc:  # noqa: BLE001
            LOGGER.info("Could not read the month from tab 3: %s", exc)
            previous_label = current_label = ""

        if not current_label:
            months = [s.current_month for s in all_series if s.current_month]
            current_label = max(set(months), key=months.count) if months else _month_label(drawing_month)
        if not previous_label:
            previous_months = [s.previous_month for s in all_series if s.previous_month]
            previous_label = max(set(previous_months), key=previous_months.count) if previous_months else ""

        result = WorkbookRead(
            dealership=dealership,
            month_label=current_label,
            previous_month_label=previous_label,
            breakdowns=breakdowns,
            campaigns=campaigns,
        )

        for item in all_series:
            metric_name = _metric_name_for(item)
            if not metric_name:
                if item.block and item.block not in result.unmapped:
                    result.unmapped.append(f"{item.block}|{item.column}".strip("|"))
                continue
            record = build_metric_record(metric_name, item.previous_raw, item.current_raw)
            if record is None:
                result.warnings.append(
                    f"{metric_name}: could not read "
                    f"'{item.previous_raw}' -> '{item.current_raw}'"
                )
                continue
            result.metrics[metric_name] = record

        return result
    finally:
        workbook.close()


def _metric_name_for(item: Series) -> str:
    total = MATRIX_TOTAL_MAP.get((item.column, item.block))
    if total:
        return total

    mapped = METRIC_MAP.get(item.block)
    if isinstance(mapped, str):
        return mapped if not item.column else ""
    if isinstance(mapped, dict):
        return mapped.get(item.column, "")
    return ""


# --- the funnel engine's breakdowns -------------------------------------------------

PIPELINE_STAGES = ["Unqualified", "Contacted", "Store Visit", "Demo",
                   "Write-Up", "F&I", "Deals", "Lost"]

SOURCE_LABELS = {"WALK IN": "Walk-in", "PHONE UP": "Phone Up",
                 "INTERNET": "Internet", "OTHER": "Other"}


def _pipeline_from_data_tab(workbook, previous_label: str, current_label: str):
    """Pipeline stages, which are charted on tab 1 but tabulated on tab 3.

    The chart is a picture; the third tab holds the same figures as cells. Reading them
    there avoids trying to recover numbers from a rendered chart, which is what the PDF
    path had to do.
    """
    stages: dict[str, dict[str, int | None]] = {}
    if len(workbook.sheetnames) < 3:
        return stages
    worksheet = workbook[workbook.sheetnames[2]]
    values = [[cell.value for cell in row] for row in worksheet.iter_rows()]

    header_row = next((i for i, row in enumerate(values)
                       if any(isinstance(v, str) and _normalise(v) == "MONTH" for v in row)), None)
    if header_row is None:
        return stages
    dated = [(i, _month_label(v)) for i, v in enumerate(values[header_row]) if _is_date(v)]
    wanted = {previous_label: None, current_label: None}
    for column, label in dated:
        if label in wanted:
            wanted[label] = column   # later columns win: the report pair sits rightmost

    for row in values:
        label = next((v for v in row[:3] if isinstance(v, str) and v.strip()), None)
        if not label:
            continue
        match = next((s for s in PIPELINE_STAGES if _normalise(s) == _normalise(label)), None)
        if not match or match in stages:
            continue
        entry: dict[str, int | None] = {}
        for month_label, column in wanted.items():
            value = row[column] if column is not None and column < len(row) else None
            entry[month_label] = int(value) if isinstance(value, (int, float)) else None
        stages[match] = entry
    return stages


def build_breakdowns(read: "WorkbookRead", workbook_source: bytes | str | Path) -> dict[str, Any]:
    """The same structure ``report_breakdowns.parse_breakdowns`` returns, from Excel.

    Built to that shape on purpose. The funnel engine, the month model and every test
    around them stay untouched: they cannot tell whether the figures were scraped out
    of a PDF or read from the cells they were typed into.
    """
    from reporting.report_breakdowns import Campaign, CampaignBlock, SourceRow

    def by_source(block: str) -> dict[str, Any]:
        # Workbooks print "WALK-IN"; SOURCE_LABELS keys use a space.
        return {re.sub(r"[\s-]+", " ", key).strip(): value
                for key, value in (read.breakdowns.get(block) or {}).items()}

    leads = by_source("NUMBER OF LEADS")
    deals = by_source("NUMBER OF DEALS")
    conversion = by_source("CONVERSION RATIOS")

    sources: list[SourceRow] = []
    for key, label in SOURCE_LABELS.items():
        if key not in leads and key not in deals:
            continue

        def figure(block: dict[str, Any], month: str, scale: float = 1.0):
            entry = block.get(key) or {}
            value = entry.get(month)
            return round(float(value) * scale, 2) if isinstance(value, (int, float)) else None

        sources.append(SourceRow(
            source=label,
            leads_previous=figure(leads, "previous"),
            leads_current=figure(leads, "current"),
            deals_previous=figure(deals, "previous"),
            deals_current=figure(deals, "current"),
            conversion_previous=figure(conversion, "previous", 100.0),
            conversion_current=figure(conversion, "current", 100.0),
        ))

    def block_for(kind: str, count_metric: str, roi_metric: str) -> CampaignBlock:
        named = [
            Campaign(name=str(item["name"]),
                     count=int(item["deals"] or 0),
                     roi=int(item["roi"] or 0))
            for item in read.campaigns.get(kind, [])
            if item.get("deals") is not None
        ]
        count_record = read.metrics.get(count_metric) or {}
        roi_record = read.metrics.get(roi_metric) or {}

        def figure(record: dict[str, Any], key: str) -> int | None:
            value = record.get(key)
            return int(value) if isinstance(value, (int, float)) else None

        # Both months, not just the current one. The funnel engine's campaign finding
        # is a movement -- "return reached X, up from Y" -- so a block with no previous
        # figure silently drops that insight and a lesser one takes its place.
        return CampaignBlock(
            named=named[:3],
            stated_count=figure(count_record, "current_value"),
            stated_roi=figure(roi_record, "current_value"),
            previous_count=figure(count_record, "previous_value"),
            previous_roi=figure(roi_record, "previous_value"),
        )

    if read.pipeline:
        stages = read.pipeline
    else:
        handle = io.BytesIO(workbook_source) if isinstance(workbook_source, bytes) else workbook_source
        workbook = load_workbook(handle, data_only=True, read_only=True)
        try:
            stages = _pipeline_from_data_tab(workbook, read.previous_month_label, read.month_label)
        finally:
            workbook.close()

    return {
        "lead_sources": sources,
        "pipeline_stages": stages,
        "sales_campaigns": block_for("sales", "sales_campaign_deals", "sales_campaign_roi"),
        "service_campaigns": block_for("service", "service_bookings_from_campaigns", "service_campaign_roi"),
    }


# --- discovery: one workbook per dealership, no month in the filename ---------------
#
# The PDF era put the month in the filename, so discovery scored filenames for month
# tokens and a dealership had twelve files a year. A master workbook is one file per
# dealership that keeps being updated, so the filename says who and the sheet says
# when. Reading the month from the file rather than its name also removes a whole
# class of failure: a workbook saved as "August" that actually holds July's figures
# now reports July, because that is what is in it.

@dataclass
class MasterFile:
    """One dealership's workbook, and the month it currently holds."""

    path: Path
    dealership: str
    month_label: str
    previous_month_label: str
    error: str = ""
    # Every month with figures in it, oldest first. A workbook holds years of history,
    # so the month picker offers all of them rather than only the newest.
    months: list[str] = field(default_factory=list)

    def holds(self, month_label: str) -> bool:
        return month_label in self.months

    @property
    def is_readable(self) -> bool:
        return not self.error and bool(self.month_label)


def read_months(source: bytes | str | Path) -> tuple[str, str]:
    """The month pair this workbook currently holds, from the latest columns of tab 3.

    Tab 3 lays the year out in dated columns, so the rightmost populated one is the
    month the reports are showing. A column is only counted when something was
    actually entered under it -- these sheets are laid out a year at a time, so
    empty columns for months yet to come are normal and must not be mistaken for
    the current month.

    Returns ``(previous, current)``. Either may be empty on a workbook that has not
    been filled in yet.
    """
    handle = io.BytesIO(source) if isinstance(source, bytes) else source
    workbook = load_workbook(handle, data_only=True, read_only=True)
    try:
        if len(workbook.sheetnames) < 3:
            return "", ""
        worksheet = workbook[workbook.sheetnames[2]]
        rows = [[cell for cell in row] for row in worksheet.iter_rows(max_row=90, values_only=True)]

        header_index = next(
            (i for i, row in enumerate(rows)
             if any(isinstance(v, str) and _normalise(v) == "MONTH" for v in row)),
            None,
        )
        if header_index is None:
            header_index = next(
                (i for i, row in enumerate(rows) if sum(1 for v in row if _is_date(v)) >= 2),
                None,
            )
        if header_index is None:
            return "", ""

        header = rows[header_index]
        populated: list[tuple[dt.date, int]] = []
        for column, value in enumerate(header):
            if not _is_date(value):
                continue
            has_data = any(
                isinstance(_cell(rows, r, column), (int, float, str))
                and str(_cell(rows, r, column)).strip() not in ("", "N/A", "#REF!")
                for r in range(header_index + 1, min(header_index + 80, len(rows)))
            )
            if has_data:
                moment = value.date() if isinstance(value, dt.datetime) else value
                populated.append((moment, column))

        if not populated:
            return "", ""

        # Rightmost wins a tie: these sheets repeat the reported month in a later
        # column beside the previous one, and that pair is what the reports show.
        populated.sort(key=lambda item: (item[0], item[1]))
        current = populated[-1][0]
        earlier = [moment for moment, _ in populated if moment < current]
        previous = max(earlier) if earlier else None
        return (_month_label(previous) if previous else "", _month_label(current))
    finally:
        workbook.close()


def workbook_paths(root: str | Path, *, recursive: bool = True) -> list[Path]:
    """Workbook files under a folder, found by name alone -- none of them is opened."""
    folder = Path(root).expanduser()
    if not folder.is_dir():
        return []
    # In the library the workbooks sit at the top level, one per dealership, with a
    # folder of exported PDFs beside each. Searching below that level would find the
    # same dealership twice.
    paths = folder.glob("*.xls[xm]") if not recursive else folder.rglob("*.xls[xm]")
    return [
        path for path in sorted(paths)
        if not path.name.startswith("~$")      # Excel's lock files
        and not any(part.lower() in {"archive", "archived", "old"} for part in path.parts)
    ]


def discover_masters(
    root: str | Path,
    *,
    limit: int = 2000,
    recursive: bool = True,
    include: Iterable[str] | None = None,
) -> list[MasterFile]:
    """Every dealership workbook under a folder, with the month each one holds.

    Deliberately tolerant. One unreadable workbook -- locked by someone with it open,
    half-synced, saved in an older format -- records its error and does not stop the
    other ninety-nine from being listed.

    ``include`` limits discovery to those workbook filenames (case-insensitive). The
    filter runs before a file is opened, which is where all the time goes.
    """
    folder = Path(root).expanduser()
    found: list[MasterFile] = []
    if not folder.is_dir():
        return found

    wanted = {str(name).casefold() for name in include} if include else None
    for path in workbook_paths(folder, recursive=recursive):
        if len(found) >= limit:
            break
        if wanted is not None and path.name.casefold() not in wanted:
            continue

        dealership = ""
        try:
            dealership, _ = _title_from_drawings(path)
        except Exception:  # noqa: BLE001
            dealership = ""
        # The folder is a better name than the file when the file is called
        # something generic, which in a shared library it usually is.
        if not dealership or _normalise(dealership) in _GENERIC_TITLES:
            dealership = path.parent.name if path.parent != folder else path.stem

        try:
            previous, current = read_months(path)
            # The earliest month has nothing before it to compare against, so it is not
            # offered: a month on its own produces no movement and no insight.
            held = list_months(path)
            entry = MasterFile(path=path, dealership=dealership,
                               month_label=current, previous_month_label=previous,
                               months=held[1:] if len(held) > 1 else [])
        except Exception as exc:  # noqa: BLE001
            entry = MasterFile(path=path, dealership=dealership, month_label="",
                               previous_month_label="", error=str(exc))
        found.append(entry)

    # Sites in one group often all carry the group's title; a shared title is no name at all.
    titles = Counter(_normalise(entry.dealership) for entry in found)
    for entry in found:
        if titles[_normalise(entry.dealership)] > 1:
            entry.dealership = entry.path.stem

    unique: list[MasterFile] = []
    seen: set[str] = set()
    for entry in found:
        key = _normalise(entry.dealership) or str(entry.path)
        if key in seen:
            LOGGER.warning("Skipping %s: another workbook is already named %s", entry.path, entry.dealership)
            continue
        seen.add(key)
        unique.append(entry)
    return unique


def available_months(root: str | Path, include: Iterable[str] | None = None) -> list[str]:
    """Which months the workbooks in a folder actually hold, newest first.

    This replaces scanning filenames for month tokens. The operator picks from months
    that exist rather than from a calendar, so choosing a month with no data stops
    being the most common way a run comes back empty.
    """
    labels = [label for master in discover_masters(root, include=include) for label in master.months]

    def sort_key(label: str) -> tuple[int, int]:
        parts = label.split()
        month = MONTH_NAMES.index(parts[0]) + 1 if parts and parts[0] in MONTH_NAMES else 0
        year = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0
        return (year, month)

    return sorted(set(labels), key=sort_key, reverse=True)


# --- reading any month, not just the latest -----------------------------------------
#
# Tabs 1 and 2 are the rendered report: they show one month against the one before it,
# and they are what the exported PDF shows. Tab 3 is the history behind that report --
# every month the dealership has, laid out in dated columns.
#
# So the reader uses whichever is right for the month asked for. The latest month comes
# from tabs 1 and 2, whose figures are the ones printed on the PDF beside the email. An
# earlier month has no PDF and never will, so it is read from tab 3.

DATA_SECTIONS = {
    "NEW LEADS", "DEALS", "CLOSING %", "APPOINTMENTS", "SALES PIPLINE", "RESPONSE TIME",
    "TEST DRIVES", "TASKS", "EMAIL CAPTURE", "SAMEDAY DEALS", "SALES CAMPAIGNS",
    "SUCCESSFUL CAMPAIGNS", "SERVICE REMINDERS", "NO SHOWS", "SERVICE CAMPAIGNS",
    "SUCCESFUL CAMPAIGNS", "INTERNAL SCORING", "TOTAL FOR WORKFRONT",
}

# (section, row label) -> (engine metric name, kind). "rate" is stored as a fraction in
# the sheet and printed as a percentage; "currency" gets its dollar sign back.
DATA_METRIC_MAP: dict[tuple[str, str], tuple[str, str]] = {
    ("NEW LEADS", "TOTAL LEADS"): ("sales_total_leads", "count"),
    ("DEALS", "TOTAL DEALS"): ("sales_total_deals", "count"),
    ("CLOSING %", "OVERALL"): ("sales_overall_conversion_rate", "rate"),
    ("SALES PIPLINE", "UNQUALIFIED"): ("sales_uncontacted_leads", "count"),
    ("SALES PIPLINE", "PIPELINE TOTAL"): ("sales_pipeline_size", "count"),
    ("RESPONSE TIME", "INTERNET LEAD RESPONSE TIME"): ("sales_internet_response_time", "text"),
    ("TASKS", "NUMBER OF OVERDUE SALES TASKS"): ("sales_overdue_tasks", "count"),
    ("TASKS", "NUMBER OF OVERDUE SERVICE TASKS"): ("service_outstanding_tasks", "count"),
    ("EMAIL CAPTURE", "EMAIL CAPTURE RATE%"): ("sales_email_capture_rate", "rate"),
    ("SAMEDAY DEALS", "SAMEDAY DEAL %"): ("sales_same_day_deal_pct", "rate"),
    ("SALES CAMPAIGNS", "NUMBER OF DEALS"): ("sales_campaign_deals", "count"),
    ("SALES CAMPAIGNS", "ROI"): ("sales_campaign_roi", "currency"),
    ("SERVICE REMINDERS", "SERVICE REMINDERS SENT"): ("service_reminders_sent", "count"),
    ("NO SHOWS", "NO SHOWS"): ("service_no_shows", "count"),
    ("NO SHOWS", "REBOOKED NO SHOWS"): ("service_rebooked_no_shows", "count"),
    ("SERVICE CAMPAIGNS", "NUMBER OF ENROLLED"): ("service_customers_enrolled", "count"),
    ("SERVICE CAMPAIGNS", "SERVICE BOOKINGS FROM CAMPAIGNS"): ("service_bookings_from_campaigns", "count"),
    ("SERVICE CAMPAIGNS", "ROI FROM SERVICE CAMPAIGNS"): ("service_campaign_roi", "currency"),
}

# Rates the report prints but the history tab stores only the ingredients for.
# (metric, numerator, denominator) -- all three as (section, label) pairs.
DATA_DERIVED_RATES = (
    ("sales_new_leads_to_appointments",
     ("APPOINTMENTS", "NUMBER OF APPOINTMENTS"), ("NEW LEADS", "TOTAL LEADS")),
    ("sales_new_leads_to_test_drives",
     ("TEST DRIVES", "NUMBER OF TEST DRIVES"), ("NEW LEADS", "TOTAL LEADS")),
    ("sales_appointment_to_sales",
     ("APPOINTMENTS", "NUMBER OF APPOINTMENTS"), ("DEALS", "TOTAL DEALS")),
)

DATA_SOURCE_ROWS = {
    "WALK-IN": "Walk-in", "PHONE ENQ": "Phone Up",
    "INTERNET": "Internet", "OTHER LEAD TYPES": "Other",
}

_CAMPAIGN_ROW = re.compile(r"CAMPAIGN\s*(\d+)\s*-?\s*(NAME|NUMBER OF DEALS|ROI)", re.I)


def _data_table(source: bytes | str | Path):
    """The history tab as (months, rows).

    ``months`` is every month with figures under it, oldest first. ``rows`` is one entry
    per labelled row: its section, its label, and its value per month.

    Section matters because the labels repeat: "Other lead Types" appears under leads,
    under deals and again under closing rates, and "Number of Enrolled" belongs to both
    the sales and the service campaign blocks.
    """
    handle = io.BytesIO(source) if isinstance(source, bytes) else source
    workbook = load_workbook(handle, data_only=True)
    try:
        if len(workbook.sheetnames) < 3:
            return [], []
        values, formats = _grid(workbook[workbook.sheetnames[2]])

        header_index = next(
            (i for i, row in enumerate(values)
             if any(isinstance(v, str) and _normalise(v) == "MONTH" for v in row)),
            None,
        )
        if header_index is None:
            header_index = next(
                (i for i, row in enumerate(values) if sum(1 for v in row if _is_date(v)) >= 2),
                None,
            )
        if header_index is None:
            return [], []

        dated = [(v.date() if isinstance(v, dt.datetime) else v, c)
                 for c, v in enumerate(values[header_index]) if _is_date(v)]

        def has_figures(column: int) -> bool:
            for r in range(header_index + 1, len(values)):
                cell = _cell(values, r, column)
                if isinstance(cell, (int, float)):
                    return True
                if isinstance(cell, str) and cell.strip() and cell.strip().upper() not in {"N/A", "NA", "#REF!", "."}:
                    return True
            return False

        # The same month can appear twice: once in the running year and again in the
        # pair the report renders. The rightmost is the one the report used.
        by_month: dict[dt.date, int] = {}
        for moment, column in dated:
            if has_figures(column):
                by_month[moment] = column
        months = sorted(by_month)

        rows: list[dict[str, Any]] = []
        section = ""
        for index in range(header_index + 1, len(values)):
            label = next((v for v in values[index][:3] if isinstance(v, str) and v.strip()), None)
            if not label:
                continue
            clean = _normalise(label)
            numeric = any(isinstance(_cell(values, index, by_month[m]), (int, float)) for m in months)
            if clean in DATA_SECTIONS and not numeric:
                section = clean
                continue
            rows.append({
                "section": section,
                "label": clean,
                "values": {m: _cell(values, index, by_month[m]) for m in months},
                "formats": {m: _cell(formats, index, by_month[m]) for m in months},
            })
        return months, rows
    finally:
        workbook.close()


def _find_row(rows, key: tuple[str, str]):
    section, label = key
    for row in rows:
        if row["section"] == section and row["label"] == label:
            return row
    return None


def _data_raw(row, month, kind: str) -> str:
    """One cell from the history tab, rendered the way the report renders it."""
    if row is None:
        return ""
    value = row["values"].get(month)
    if value is None or value == "":
        return ""
    if isinstance(value, str):
        return value.strip()
    number_format = str(row["formats"].get(month) or "")
    if kind == "rate":
        places = _decimals(number_format) if "%" in number_format else 0
        return f"{float(value) * 100:.{places}f}%"
    if kind == "currency":
        return f"${float(value):,.0f}"
    return _raw_value(value, number_format)


def list_months(source: bytes | str | Path) -> list[str]:
    """Every month this workbook has figures for, oldest first."""
    months, _ = _data_table(source)
    return [_month_label(m) for m in months]


def read_month(source: bytes | str | Path, month_label: str, *, fallback_name: str = "") -> WorkbookRead:
    """Read one particular month out of the history tab.

    Used for any month except the workbook's latest, which comes from the rendered
    report on tabs 1 and 2 instead so that its figures match the PDF exactly.
    """
    from dealership_insights_local import build_metric_record

    months, rows = _data_table(source)
    wanted = next((m for m in months if _month_label(m) == month_label), None)
    dealership, _ = _title_from_drawings(source)
    result = WorkbookRead(
        dealership=dealership or fallback_name,
        month_label=month_label,
        previous_month_label="",
    )
    if wanted is None:
        result.warnings.append(f"{month_label} is not in this workbook.")
        return result

    earlier = [m for m in months if m < wanted]
    if not earlier:
        result.warnings.append(f"{month_label} is the earliest month in this workbook, so "
                               "there is nothing to compare it against.")
        return result
    previous = earlier[-1]
    result.previous_month_label = _month_label(previous)

    for key, (metric_name, kind) in DATA_METRIC_MAP.items():
        row = _find_row(rows, key)
        current_raw = _data_raw(row, wanted, kind)
        # Pipeline size is a snapshot: the report prints "current pipeline size" on its
        # own, with nothing to compare it against. The history tab does hold last
        # month's, but using it would make the same month read differently depending on
        # which tab it came from, and would put a movement in a historical report that
        # the current month's report would never mention.
        previous_raw = current_raw if metric_name == "sales_pipeline_size" else _data_raw(row, previous, kind)
        record = build_metric_record(metric_name, previous_raw, current_raw)
        if record is not None:
            result.metrics[metric_name] = record

    for metric_name, numerator_key, denominator_key in DATA_DERIVED_RATES:
        numerator = _find_row(rows, numerator_key)
        denominator = _find_row(rows, denominator_key)

        def rate(month) -> str:
            if numerator is None or denominator is None:
                return ""
            top, bottom = numerator["values"].get(month), denominator["values"].get(month)
            if not isinstance(top, (int, float)) or not isinstance(bottom, (int, float)):
                return ""
            if abs(float(bottom)) < 1e-9:
                return ""
            return f"{float(top) / float(bottom) * 100:.0f}%"

        record = build_metric_record(metric_name, rate(previous), rate(wanted))
        if record is not None:
            result.metrics[metric_name] = record

    # Lead sources, for the funnel engine's "with Other accounting for N of the drop".
    for block, section in (("NUMBER OF LEADS", "NEW LEADS"),
                           ("NUMBER OF DEALS", "DEALS"),
                           ("CONVERSION RATIOS", "CLOSING %")):
        for label, display in DATA_SOURCE_ROWS.items():
            row = _find_row(rows, (section, label))
            if row is None:
                continue
            result.breakdowns.setdefault(block, {})[display.upper()] = {
                "previous": row["values"].get(previous),
                "current": row["values"].get(wanted),
            }

    for kind, section in (("sales", "SUCCESSFUL CAMPAIGNS"), ("service", "SUCCESFUL CAMPAIGNS")):
        named: dict[int, dict[str, Any]] = {}
        for row in rows:
            if row["section"] != section:
                continue
            match = _CAMPAIGN_ROW.match(row["label"])
            if not match:
                continue
            slot = int(match.group(1))
            field = match.group(2).upper()
            value = row["values"].get(wanted)
            entry = named.setdefault(slot, {})
            if field == "NAME":
                entry["name"] = str(value).strip() if isinstance(value, str) else ""
            elif field == "NUMBER OF DEALS":
                entry["deals"] = value if isinstance(value, (int, float)) else None
            else:
                entry["roi"] = value if isinstance(value, (int, float)) else None
        campaigns = [entry for _, entry in sorted(named.items())
                     if entry.get("name") and entry.get("name").upper() not in {"N/A", "NA"}
                     and entry.get("deals") is not None]
        if campaigns:
            result.campaigns[kind] = campaigns

    result.pipeline = {}
    for stage in PIPELINE_STAGES:
        row = _find_row(rows, ("SALES PIPLINE", _normalise(stage)))
        if row is None:
            continue
        result.pipeline[stage] = {
            result.previous_month_label: row["values"].get(previous)
            if isinstance(row["values"].get(previous), (int, float)) else None,
            month_label: row["values"].get(wanted)
            if isinstance(row["values"].get(wanted), (int, float)) else None,
        }
    return result


def read_for_month(
    source: bytes | str | Path,
    month_label: str = "",
    *,
    fallback_name: str = "",
) -> WorkbookRead:
    """Read a workbook for a particular month, from whichever tab is right for it.

    The workbook's latest month is read from the rendered report on tabs 1 and 2, whose
    figures are the ones printed on the PDF that gets attached to the email. Any earlier
    month is read from the history tab, because no PDF of it exists.
    """
    latest = ""
    try:
        _, latest = read_months(source)
    except Exception:  # noqa: BLE001
        latest = ""

    if not month_label or month_label == latest:
        return read_master(source, fallback_name=fallback_name)
    return read_month(source, month_label, fallback_name=fallback_name)
