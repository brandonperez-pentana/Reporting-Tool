"""What the month actually did, computed once, before a word is written.

The engine's long-standing problem is that it has no such thing. Rules fuse three jobs
-- deciding what is significant, judging whether it is good or bad, and phrasing it --
and every later stage recovers the first two by re-reading the third. Four separate
bugs this month traced back to that single absence, and each needed its own local patch
because nothing could answer "what happened".

This module answers it. Its output is data: stations, movements, drivers, a verdict. No
sentences appear anywhere in it, and nothing here decides what to say -- that belongs to
the layer above, which reads this and only this.

The shape follows the operator's brief. A dealership is a sales funnel and a bottom
line, nothing more complicated. Two things fall out of that:

  * Position beats magnitude. Metrics are stations on a chain, not peers to be ranked by
    how far they moved -- a loss high in the chain outweighs a bigger percentage lower
    down, because everything downstream inherits it.
  * Every movement carries its driver. "Campaign return fell" is a number; "campaign
    return fell, and this campaign lost money" is the reason, and the reason is the part
    that saves the reader the work.

Reliability gates are imported from the insight engine rather than reimplemented. Two
independent notions of "is this metric usable" is precisely the class of divergence this
architecture exists to remove.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

import dealership_insights_local as core

from reporting.report_breakdowns import CampaignBlock, SourceRow

SALESPERSON = "salesperson"
BUSINESS = "business"


@dataclass
class Movement:
    """One metric across the two months the report compares."""

    metric: str
    label: str
    previous: float | None
    current: float | None
    percent_change: float | None
    improving: bool
    usable: bool
    band: str
    unit: str = "count"
    metric_type: str = "count"
    point_change: float | None = None

    @property
    def moved(self) -> bool:
        return self.usable and self.band in {"meaningful", "strong"}


@dataclass
class Attribution:
    """What inside a movement did the moving.

    `share_of_total` is a share of the *stated* total, and `unnamed` records how much of
    that total is not covered by anything nameable. Both matter: the report only ever
    names its top three campaigns, so a driver is always a part of the month rather than
    an account of it, and copy written from this must say so.
    """

    driver: str
    value: float
    unit: str
    share_of_total: float | None = None
    unnamed: int = 0
    kind: str = ""


@dataclass
class Station:
    """One stop on a funnel: what it measures, who controls it, how it moved."""

    key: str
    label: str
    owner: str
    movement: Movement | None = None
    attribution: Attribution | None = None
    note: str = ""

    @property
    def reported(self) -> bool:
        return self.movement is not None and self.movement.usable


@dataclass
class Funnel:
    name: str
    stations: list[Station] = field(default_factory=list)

    def station(self, key: str) -> Station | None:
        return next((s for s in self.stations if s.key == key), None)

    def largest_leak(self) -> Station | None:
        """The earliest station with a material adverse movement.

        Earliest, not largest: a fall at the top of the funnel is inherited by every
        station below it, so reporting the downstream symptom as the headline points the
        reader at the wrong place. Magnitude only breaks ties among unreported stations.
        """
        for station in self.stations:
            if station.reported and station.movement.moved and not station.movement.improving:
                return station
        return None


# Sales funnel. Order is the chain itself, and largest_leak() depends on it.
SALES_STATIONS: list[tuple[str, str, str, str]] = [
    ("leads", "Lead volume", "sales_total_leads", SALESPERSON),
    ("contact", "Uncontacted leads", "sales_uncontacted_leads", SALESPERSON),
    ("appointments", "Lead-to-appointment conversion", "sales_new_leads_to_appointments", SALESPERSON),
    ("test_drives", "Lead-to-test-drive conversion", "sales_new_leads_to_test_drives", SALESPERSON),
    ("close", "Appointment-to-sale conversion", "sales_appointment_to_sales", SALESPERSON),
    ("units", "Deal volume", "sales_total_deals", SALESPERSON),
]

# Service is a retention funnel and behaves identically: reminders go out, customers
# enrol, bookings land, work is invoiced.
SERVICE_STATIONS: list[tuple[str, str, str, str]] = [
    ("reminders", "Service reminders sent", "service_reminders_sent", BUSINESS),
    ("enrolled", "Customers enrolled in service", "service_customers_enrolled", BUSINESS),
    ("bookings", "Service bookings from campaigns", "service_bookings_from_campaigns", BUSINESS),
]

# Reminders sent and customers enrolled print as separate figures on the service report
# and are, in practice, the same number: across a real month of 23 dealerships they were
# identical in all 23. Two stations carrying one figure state the same fact twice in the
# insights and count it twice in breadth(), which then leans the verdict. They are
# collapsed when they agree and kept apart when they do not, so a dealership that ever
# reports them separately still gets both.
DUPLICATE_SERVICE_STATIONS = ("reminders", "enrolled")

# Reported alongside the funnels rather than inside them: they are how well the follow-up
# is being kept up, not stops a lead passes through.
HYGIENE_METRICS = [
    ("sales_overdue_tasks", "Overdue sales tasks", SALESPERSON),
    ("sales_email_capture_rate", "Email capture rate", SALESPERSON),
    ("service_outstanding_tasks", "Outstanding service tasks", BUSINESS),
    ("service_no_shows", "Service no-shows", BUSINESS),
]

# Response time is measured from when a lead is actioned, so an uncontacted lead *is*
# the response-time figure. They are one phenomenon, and reporting both counts a single
# problem twice. Contact is the cause and holds the station; this is kept only as
# supporting evidence for it.
RESPONSE_TIME_METRIC = "sales_internet_response_time"


def _movement(metrics: dict[str, dict[str, Any]], name: str, label: str) -> Movement | None:
    record = (metrics or {}).get(name)
    if not record:
        return None
    usable = core.metric_usable_for_reporting(record)
    return Movement(
        metric=name,
        label=label,
        previous=record.get("previous_value"),
        current=record.get("current_value"),
        percent_change=record.get("percent_change"),
        improving=not core.metric_is_adverse(metrics, name),
        usable=usable,
        band=str(record.get("significance_band", "")),
        unit=str(record.get("unit", "count")),
        metric_type=str(record.get("metric_type", "count")),
        point_change=record.get("percentage_point_change"),
    )


def _same_figure(first: Station | None, second: Station) -> bool:
    """Do two stations carry the identical pair of figures?

    Compared rather than assumed: the collapse only happens where the report really did
    print one number twice, so a dealership whose figures diverge keeps both stations.
    """
    if first is None or first.movement is None or second.movement is None:
        return False
    return (
        (first.movement.previous, first.movement.current)
        == (second.movement.previous, second.movement.current)
    )


def _source_driver(rows: list[SourceRow], measure: str) -> Attribution | None:
    """Which lead source moved a total the most, in absolute terms.

    Every channel counts at full weight, "Other" included: it stands for whichever
    channels a given dealership set up beyond the named three, so it varies between
    dealerships and is not a residual to be discounted.
    """
    best: tuple[float, SourceRow] | None = None
    deltas: list[tuple[float, SourceRow]] = []
    for row in rows:
        if row.source.lower() == "total":
            continue
        previous = getattr(row, f"{measure}_previous")
        current = getattr(row, f"{measure}_current")
        if previous is None or current is None:
            continue
        deltas.append((current - previous, row))
    # "Deals rose, with Walk-in accounting for 5 of the drop" names a source that
    # moved against the total; only one moving with it explains the movement.
    net = sum(delta for delta, _ in deltas)
    for delta, row in deltas:
        if delta * net <= 0:
            continue
        if best is None or abs(delta) > abs(best[0]):
            best = (delta, row)
    if best is None:
        return None
    delta, row = best
    return Attribution(driver=row.source, value=delta, unit=measure, kind="lead_source")


def _campaign_driver(block: CampaignBlock) -> Attribution | None:
    """The campaign that returned the most, and how much of the month it accounts for."""
    leader = block.leader()
    if leader is None:
        return None
    share = None
    if block.stated_roi:
        share = round(leader.roi / block.stated_roi * 100, 1)
    return Attribution(
        driver=leader.name,
        value=leader.roi,
        unit="roi",
        share_of_total=share,
        unnamed=block.unnamed_count,
        kind="campaign",
    )


@dataclass
class MonthModel:
    """One dealership, one month, as data."""

    dealership: str
    month_label: str
    sales: Funnel
    service: Funnel
    hygiene: list[Station] = field(default_factory=list)
    lead_sources: list[SourceRow] = field(default_factory=list)
    pipeline_stages: dict[str, dict[str, int | None]] = field(default_factory=dict)
    sales_campaigns: CampaignBlock | None = None
    service_campaigns: CampaignBlock | None = None

    # --- what the layer above asks it -------------------------------------------

    def diagnostic_pair(self) -> list[Station]:
        """Lead-to-appointment and lead-to-test-drive.

        The operator's own rule: when these two are healthy the floor is working the
        leads it gets, and the rest generally follows -- campaigns excepted, since they
        run outside any salesperson's control.
        """
        return [s for s in (self.sales.station("appointments"), self.sales.station("test_drives")) if s]

    def reported_stations(self) -> list[Station]:
        return [s for s in [*self.sales.stations, *self.service.stations, *self.hygiene] if s.reported]

    def bottom_line(self) -> dict[str, Movement | None]:
        """Units sold and what the campaigns returned. The only scoreboard that counts."""
        units = self.sales.station("units")
        return {
            "units": units.movement if units else None,
            "sales_campaign_roi": self._campaign_roi_movement(self.sales_campaigns),
            "service_campaign_roi": self._campaign_roi_movement(self.service_campaigns),
        }

    @staticmethod
    def _campaign_roi_movement(block: CampaignBlock | None) -> Movement | None:
        if block is None or block.stated_roi is None or block.previous_roi is None:
            return None
        previous, current = float(block.previous_roi), float(block.stated_roi)
        if previous == 0 and current == 0:
            return None
        percent = round((current - previous) / previous * 100, 1) if previous else None
        magnitude = abs(percent) if percent is not None else None
        if magnitude is None:
            band = "strong"          # off a zero base, any return at all is the story
        elif magnitude >= 25:
            band = "strong"
        elif magnitude >= 10:
            band = "meaningful"
        elif magnitude >= 3:
            band = "slight"
        else:
            band = "negligible"
        return Movement(
            metric="campaign_roi", label="Campaign return",
            previous=previous, current=current, percent_change=percent,
            improving=current >= previous, usable=True, band=band,
            unit="currency", metric_type="currency",
        )

    def breadth(self) -> dict[str, int]:
        """How many measures firmed and how many softened, materially."""
        firmed = softened = 0
        for station in self.reported_stations():
            if not station.movement.moved:
                continue
            if station.movement.improving:
                firmed += 1
            else:
                softened += 1
        return {"firmed": firmed, "softened": softened}

    def bottom_line_directions(self) -> set[bool]:
        """Which way each part of the scoreboard moved, materially.

        Two entries mean the scoreboard disagrees with itself, which is the one thing
        no single-direction verdict can honestly describe.
        """
        return {m.improving for m in self.bottom_line().values() if m is not None and m.moved}

    def verdict(self) -> str:
        """good / soft / mixed / quiet, judged the way the operator judges it.

        Three rules, in order.

        A scoreboard that disagrees with itself is a split month, whatever the breadth
        of the smaller metrics says. Units up while campaign return fell 40% is not
        "ahead across the main measures" -- the opener claims a direction the next
        sentence then contradicts, which is exactly the complaint this engine was built
        to answer.

        Where the scoreboard is unanimous, breadth decides whether it holds up. Units
        rising while more of the funnel softened than firmed is a warning, not a pass,
        because next month is already in trouble; and one campaign cannot carry a month
        on its own while everything feeding it slips. An even split counts against the
        scoreboard for the same reason: half the month moving the other way is not
        "across the board". A funnel that reported nothing cannot argue with anything,
        so it does not get a vote -- silence is not disagreement.

        Only when nothing on the scoreboard moved does the funnel judge alone.
        """
        counts = self.breadth()
        directions = self.bottom_line_directions()

        if not directions and counts["firmed"] == 0 and counts["softened"] == 0:
            return "quiet"

        if len(directions) > 1:
            return "mixed"

        if directions:
            firmer = next(iter(directions))
            if counts["firmed"] + counts["softened"]:
                if firmer and counts["softened"] >= counts["firmed"]:
                    return "mixed"
                if not firmer and counts["firmed"] >= counts["softened"]:
                    return "mixed"
            return "good" if firmer else "soft"

        if counts["softened"] and counts["firmed"] == 0:
            return "soft"
        if counts["firmed"] and counts["softened"] == 0:
            return "good"
        return "mixed"

    def missing(self) -> list[str]:
        """Stations the reports gave nothing usable for."""
        return [s.key for s in [*self.sales.stations, *self.service.stations] if not s.reported]


def build_month_model(
    *,
    dealership: str,
    month_label: str,
    metrics: dict[str, dict[str, Any]],
    breakdowns: dict[str, Any],
) -> MonthModel:
    """Assemble the model from parsed metrics plus the three breakdown tables."""
    lead_sources = breakdowns.get("lead_sources") or []
    sales_campaigns = breakdowns.get("sales_campaigns")
    service_campaigns = breakdowns.get("service_campaigns")

    sales = Funnel("sales")
    for key, label, metric_name, owner in SALES_STATIONS:
        station = Station(key=key, label=label, owner=owner, movement=_movement(metrics, metric_name, label))
        if key == "leads":
            station.attribution = _source_driver(lead_sources, "leads")
        elif key == "units":
            station.attribution = _source_driver(lead_sources, "deals")
        elif key == "contact":
            response = _movement(metrics, RESPONSE_TIME_METRIC, "Internet lead response time")
            if response and response.usable:
                station.note = "response_time_follows_contact"
        sales.stations.append(station)

    service = Funnel("service")
    for key, label, metric_name, owner in SERVICE_STATIONS:
        station = Station(key=key, label=label, owner=owner, movement=_movement(metrics, metric_name, label))
        if key == DUPLICATE_SERVICE_STATIONS[1]:
            first = service.station(DUPLICATE_SERVICE_STATIONS[0])
            if _same_figure(first, station):
                continue
        if key == "bookings" and service_campaigns is not None:
            station.attribution = _campaign_driver(service_campaigns)
            gated = station.movement is not None and not station.movement.usable
            if gated and service_campaigns.runs_campaigns:
                station.note = "gated_metric_corroborated_by_named_campaigns"
        service.stations.append(station)

    hygiene = [
        Station(key=name, label=label, owner=owner, movement=_movement(metrics, name, label))
        for name, label, owner in HYGIENE_METRICS
    ]

    model = MonthModel(
        dealership=dealership,
        month_label=month_label,
        sales=sales,
        service=service,
        hygiene=hygiene,
        lead_sources=lead_sources,
        pipeline_stages=breakdowns.get("pipeline_stages") or {},
        sales_campaigns=sales_campaigns,
        service_campaigns=service_campaigns,
    )
    return model
