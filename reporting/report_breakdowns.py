"""The three tables the reports carry and the metric parser throws away.

Every DealerSocket monthly report states a handful of totals and, beside each one, the
breakdown that explains it. The existing parser keeps the totals:

    lead source table   ->  total_leads, total_deals, overall_conversion_rate
    pipeline stages     ->  pipeline_size
    named campaigns     ->  campaign_roi, campaign_deals

The discarded columns are the only part of the report that answers *why* a number
moved, which is the whole point of the exercise: a month where campaign return
collapsed reads very differently once you can see which campaign went negative.

This module reads those three tables and nothing else. It computes no insights, makes
no judgements and produces no sentences -- it hands back what the page says, so a later
layer can reason about it.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

# "$44,183", "-$608", "$0". Negative campaign returns are real and are the single most
# valuable figure in the block: a campaign that costs money rather than making it.
MONEY = re.compile(r"(-?)\$\s*([\d,]+)")
MONTH_TAG = re.compile(r"\b([A-Z][a-z]{2})-(\d{2})\b")

# The left column of an interleaved block bleeds onto the same physical line as the
# right column, so a campaign row can arrive with "Jun-26 7 $10,059 " or a stray header
# fragment in front of its name. Both are stripped before the name is kept.
LEFT_COLUMN_BLEED = re.compile(
    r"^(?:"
    r"[A-Z][a-z]{2}-\d{2}\s+[\d,]+\s+-?\$[\d,]+"      # a month row: "Jul-26 65 $44,183"
    r"|RO'?s\s+ROI|DEALS\s+ROI|ROI"                     # a header fragment
    r")\s*",
    re.I,
)
CAMPAIGN_ROW = re.compile(r"^(?P<name>.+?)\s+(?P<count>\d+)\s+(?P<roi>-?\$[\d,]+)\s*$")

NOT_A_CAMPAIGN = {"n/a", "na", "none", "-", ""}


def _money(text: str) -> int | None:
    found = MONEY.search(text or "")
    if not found:
        return None
    value = int(found.group(2).replace(",", ""))
    return -value if found.group(1) else value


@dataclass
class Campaign:
    """One named campaign and what it returned."""

    name: str
    count: int          # deals on the sales report, RO's on the service report
    roi: int


@dataclass
class CampaignBlock:
    """The named campaigns beside a stated monthly total.

    `named` holds at most the three the report prints. It is emphatically **not** a
    decomposition of the total: a dealership running four campaigns names three, so the
    counts fall short. Anything written from this has to read as "X of the month's
    total", never "made up of X, Y and Z", or it is silently wrong the moment a
    dealership grows past three campaigns.
    """

    named: list[Campaign] = field(default_factory=list)
    stated_count: int | None = None
    stated_roi: int | None = None
    previous_count: int | None = None
    previous_roi: int | None = None

    @property
    def runs_campaigns(self) -> bool:
        """False when the dealership listed n/a across the board — a real, common state."""
        return bool(self.named)

    @property
    def named_count(self) -> int:
        return sum(campaign.count for campaign in self.named)

    @property
    def accounts_for_whole_total(self) -> bool:
        """Do the named campaigns cover the stated total, or is there an unnamed tail?

        Judged on counts, not money: RO and deal counts tie out exactly where the block
        is complete, while the dollar figures drift by a few hundred against the stated
        total even then.
        """
        return self.stated_count is not None and self.named_count >= self.stated_count

    @property
    def unnamed_count(self) -> int:
        if self.stated_count is None:
            return 0
        return max(0, self.stated_count - self.named_count)

    def leader(self) -> Campaign | None:
        """The campaign that returned the most. The one worth naming, and worth repeating."""
        return max(self.named, key=lambda c: c.roi, default=None)

    def losing(self) -> list[Campaign]:
        """Campaigns that returned less than nothing."""
        return [campaign for campaign in self.named if campaign.roi < 0]


@dataclass
class SourceRow:
    """One lead source across the two months the report compares."""

    source: str
    leads_previous: int | None = None
    leads_current: int | None = None
    deals_previous: int | None = None
    deals_current: int | None = None
    conversion_previous: float | None = None
    conversion_current: float | None = None


def parse_lead_sources(sales_text: str) -> list[SourceRow]:
    """The leads / deals / conversion table, one row per source plus the total.

    Every channel is carried at full weight, "OTHER" included: it stands for whichever
    lead channels a given dealership has set up outside the named three, so it varies by
    dealership and is not a residual to be discounted.
    """
    rows: list[SourceRow] = []
    for line in (sales_text or "").splitlines():
        stripped = line.strip()
        match = re.match(
            r"^(?P<source>WALK\s*IN|PHONE\s*UP|INTERNET|OTHER|TOTAL)\s+(?P<numbers>[\d\s,%]+)$",
            stripped,
            re.I,
        )
        if not match:
            continue
        numbers = re.findall(r"([\d,]+)\s*%?", match.group("numbers"))
        if len(numbers) < 6:
            continue
        values = [int(value.replace(",", "")) for value in numbers[:6]]
        rows.append(
            SourceRow(
                source=re.sub(r"\s+", " ", match.group("source")).title(),
                leads_previous=values[0], leads_current=values[1],
                deals_previous=values[2], deals_current=values[3],
                conversion_previous=float(values[4]), conversion_current=float(values[5]),
            )
        )
    return rows


def parse_pipeline_stages(sales_text: str) -> dict[str, dict[str, int | None]]:
    """The pipeline block: Unqualified through F&I, plus Deals and Lost, per month.

    A stage row can arrive with fewer numbers than there are headings — the extraction
    drops a column when the underlying chart squeezes it. Rather than shifting every
    later stage one place left, a short row is left with `None` from the point it runs
    out, so a missing value reads as missing instead of as another stage's figure.
    """
    lines = [line.strip() for line in (sales_text or "").splitlines()]
    header_index = next(
        (i for i, line in enumerate(lines) if re.search(r"Unqualified\s+Contacted", line, re.I)),
        None,
    )
    if header_index is None:
        return {}
    headers = [h.strip() for h in re.split(r"\s{1,}", lines[header_index]) if h.strip()]
    headers = _regroup_stage_headers(headers)

    stages: dict[str, dict[str, int | None]] = {name: {} for name in headers}
    for line in lines[header_index + 1: header_index + 4]:
        month = MONTH_TAG.search(line)
        if not month:
            continue
        label = f"{month.group(1)}-{month.group(2)}"
        values = [int(v) for v in re.findall(r"\b(\d+)\b", line[month.end():])]
        for position, name in enumerate(headers):
            stages[name][label] = values[position] if position < len(values) else None
    return stages


def _regroup_stage_headers(tokens: list[str]) -> list[str]:
    """Rejoin stage names the text extraction split on their internal spaces."""
    joined = " ".join(tokens)
    known = ["Unqualified", "Contacted", "Store Visit", "Demo", "Write-Up", "F&I", "Deals", "Lost"]
    return [name for name in known if name.lower() in joined.lower()]


def parse_campaign_block(text: str, *, kind: str) -> CampaignBlock:
    """Named campaigns and the stated monthly total they sit beside.

    `kind` is "sales" or "service" and only decides which stated-total row to read;
    the two reports lay the block out identically otherwise.
    """
    block = CampaignBlock()
    anchor = "DEALS FROM CAMPAIGNS" if kind == "sales" else "SERVICE BOOKINGS FROM"
    lines = (text or "").splitlines()
    start = next((i for i, line in enumerate(lines) if anchor.lower() in line.lower()), None)
    if start is None:
        return block

    month_rows: list[tuple[str, int, int]] = []
    for raw in lines[start: start + 8]:
        line = raw.strip()

        month = MONTH_TAG.match(line)
        if month:
            counts = re.match(r"^[A-Z][a-z]{2}-\d{2}\s+([\d,]+)\s+(-?\$[\d,]+)", line)
            if counts:
                month_rows.append(
                    (f"{month.group(1)}-{month.group(2)}",
                     int(counts.group(1).replace(",", "")),
                     _money(counts.group(2)) or 0)
                )

        candidate = LEFT_COLUMN_BLEED.sub("", line)
        row = CAMPAIGN_ROW.match(candidate)
        if not row:
            continue
        name = row.group("name").strip(" -")
        if name.lower() in NOT_A_CAMPAIGN or MONTH_TAG.search(name):
            continue
        if re.fullmatch(r"[\d,$%\s.-]+", name):
            continue
        block.named.append(
            Campaign(name=name, count=int(row.group("count")), roi=_money(row.group("roi")) or 0)
        )

    if month_rows:
        month_rows.sort(key=lambda item: _month_key(item[0]))
        block.previous_count, block.previous_roi = month_rows[0][1], month_rows[0][2]
        block.stated_count, block.stated_roi = month_rows[-1][1], month_rows[-1][2]
    return block


_MONTH_ORDER = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]


def _month_key(label: str) -> tuple[int, int]:
    month, year = label.split("-")
    return (int(year), _MONTH_ORDER.index(month.lower()) + 1)


def parse_breakdowns(sales_text: str, service_text: str) -> dict[str, Any]:
    """Everything this module can read, from one dealership's pair of reports."""
    return {
        "lead_sources": parse_lead_sources(sales_text),
        "pipeline_stages": parse_pipeline_stages(sales_text),
        "sales_campaigns": parse_campaign_block(sales_text, kind="sales"),
        "service_campaigns": parse_campaign_block(service_text, kind="service"),
    }
