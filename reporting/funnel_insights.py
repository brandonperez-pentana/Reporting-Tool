"""Findings and copy, read from the month model and nothing else.

This is the layer that decides what to say. It never reads a rendered sentence to work
out what it means, because every finding already carries its own direction, size and
position -- which is what the old engine had to reconstruct from prose, and got wrong.

The copy rules come from the operator, and they are narrow on purpose:

  * Movement, then driver, then the driver's share. "Campaign return reached $44,183,
    up from $10,059, with Closed RO returning $23,114 of it."
  * No counterfactuals. "At last month's conversion those leads would have returned 172
    deals" is a hypothetical, and a hypothetical is advice wearing a number as a
    disguise.
  * No instruction. Naming the driver *is* the instruction; a dealer knows what to do
    with "this campaign returned the most" without being told.
  * Never the word money. The reader converts units and dollars themselves, instantly,
    because they do it for a living.
"""

from __future__ import annotations

from dataclasses import dataclass
from itertools import combinations

from reporting.month_model import Attribution, MonthModel, Movement, Station, _campaign_driver

MAX_INSIGHTS = 4
SMS_LIMIT = 160

# One scale for every finding, so nothing is ranked against a number that means
# something different. Ordering findings by their index inside their own funnel was the
# bug this replaces: a service station at index 0 tied with lead volume at index 0, and
# then won the tie on raw percentage -- which is how a month whose leads fell by two
# thirds got reported as a rise in reminders sent.
#
# The scale is the operator's own order of interest: what was sold, what the campaigns
# returned, where the sales funnel leaked, then the service funnel, then a campaign that
# lost money, then follow-up workload.
ORDER_UNITS = 0
ORDER_CAMPAIGN_RETURN = 10
ORDER_SALES_FUNNEL = 20      # + the station's position in the sales chain
ORDER_SERVICE_FUNNEL = 40    # + the station's position in the service chain
ORDER_LOSING_CAMPAIGN = 60
ORDER_HYGIENE = 80


@dataclass
class Finding:
    """One thing worth saying, with everything the renderer needs already attached."""

    key: str
    text: str
    improving: bool
    order: int          # position in the chain; lower leads
    magnitude: float
    short: str = ""     # the SMS form: the movement, nothing else
    brief: str = ""     # the summary form: the movement and where it landed
    statement: str = ""  # the summary form when it stands alone, so it carries a verb
    hollow: bool = False  # a rate that rose only because its base shrank
    minor: bool = False   # a count too small for its percentage to mean anything
    compact: str = ""   # the SMS clause: "deals up (33→40)", "email capture up to 47%"


# 1 -> 4 is "up 300%" and still four tasks; below this a count is not a point of interest.
SMALL_COUNT = 10


def _is_small_count(movement: Movement) -> bool:
    if _is_rate(movement) or movement.metric_type == "currency" or movement.unit == "currency":
        return False
    values = (movement.previous, movement.current)
    if any(not isinstance(v, (int, float)) for v in values):
        return False
    return max(abs(v) for v in values) < SMALL_COUNT


def _is_rate(movement: Movement) -> bool:
    return movement.metric_type == "rate" or movement.unit == "percentage"


def _value(movement: Movement, value: float | None) -> str:
    """A figure carrying its own unit, so 4 -> 9 is never mistaken for four deals."""
    if value is None:
        return "—"
    if movement.metric_type == "currency" or movement.unit == "currency":
        return _money(value)
    body = _number(value)
    if _is_rate(movement):
        return f"{body}%"
    if movement.metric_type == "time" or movement.unit == "minutes":
        return f"{body} min"
    return body


def _change_note(movement: Movement) -> str:
    """How much it moved, in the unit the movement is actually measured in.

    A conversion rate going from 4% to 9% did not rise by 125%; it rose by five
    percentage points. Reporting the relative change of a rate is the kind of number
    that survives review and then loses an argument in front of the dealer.
    """
    if _is_rate(movement):
        if movement.point_change is None:
            return ""
        points = abs(movement.point_change)
        unit = "pt" if abs(points - 1) < 0.05 else "pts"
        return f"{points:,.1f} {unit}".replace(".0 ", " ")
    if movement.percent_change is None:
        return ""
    # "fell 100% to $0" says the same thing twice.
    if movement.current == 0:
        return ""
    # "down 775% to -$6,104": a percentage of a figure that crossed zero means nothing.
    if any(isinstance(v, (int, float)) and v < 0 for v in (movement.previous, movement.current)):
        return ""
    # "339.2%" is false precision on a jump of more than four times.
    if abs(movement.percent_change) >= 100:
        return f"{abs(movement.percent_change):.0f}%"
    return f"{abs(movement.percent_change):.1f}%"


def _size(movement: Movement) -> float:
    """How far it moved, in the unit it is reported in: points for a rate, % otherwise.

    Ranking a rate by its relative change made 2% -> 8% a "300%" move that outranked
    every count in the month.
    """
    if _is_rate(movement) and movement.point_change is not None:
        return abs(movement.point_change)
    return abs(movement.percent_change or 0)


def _pct(movement: Movement) -> str:
    """Kept for the SMS, where the change note is all the room there is."""
    return _change_note(movement)


def _number(value: float | None) -> str:
    if value is None:
        return "—"
    if abs(value - round(value)) < 0.05:
        return f"{round(value):,}"
    return f"{value:,.1f}"


def _money(value: float) -> str:
    sign = "-" if value < 0 else ""
    return f"{sign}${abs(value):,.0f}"


def _direction_word(movement: Movement) -> str:
    return "rose" if movement.current > movement.previous else "fell"


def _direction_short(movement: Movement) -> str:
    """up / down describes the movement, never whether it was welcome.

    Overdue tasks rising is bad news, but it is still a rise, and an SMS that calls it a
    fall is simply wrong. Polarity is carried separately, by `improving`.
    """
    return "up" if movement.current > movement.previous else "down"


def _station_sentence(station: Station, qualifier: str = "") -> str:
    """The movement, then whatever inside it did the moving."""
    movement = station.movement
    base = (
        f"{station.label} {_direction_word(movement)} from "
        f"{_value(movement, movement.previous)} to {_value(movement, movement.current)}"
    )
    change = _change_note(movement)
    if change:
        base += f" ({change})"
    attribution = station.attribution
    if attribution and attribution.kind == "lead_source" and attribution.value:
        share = _number(abs(attribution.value))
        moved = "of the drop" if attribution.value < 0 else "of the rise"
        base += f", with {_source_name(attribution.driver)} accounting for {share} {moved}"
    if qualifier:
        base += f", {qualifier}"
    return base + "."


def _source_name(source: str) -> str:
    # The report prints "OTHER"; the workbook calls the same row "Other lead types".
    return "other lead types" if source.strip().lower() == "other" else source


# Rates measured against lead volume. When leads collapse, these can rise while the
# number of appointments or test drives behind them falls.
LEAD_BASED_RATES = ("appointments", "test_drives")


def _shrunken_base(model: MonthModel, station: Station) -> Movement | None:
    """The lead movement, when this rate rose only because there were fewer leads.

    Rate times leads is what the rate produced. If that fell while the rate rose, the
    rise is a smaller denominator, and saying "conversion rose" alone tells the dealer
    the opposite of what happened on the floor.
    """
    if station.key not in LEAD_BASED_RATES:
        return None
    rate = station.movement
    leads_station = model.sales.station("leads")
    leads = leads_station.movement if leads_station else None
    if leads is None or not leads.usable:
        return None
    values = (rate.previous, rate.current, leads.previous, leads.current)
    if any(not isinstance(v, (int, float)) for v in values):
        return None
    if rate.current <= rate.previous or leads.current >= leads.previous:
        return None
    if rate.current * leads.current >= rate.previous * leads.previous:
        return None
    return leads


def _movement_clause(label: str, movement: Movement, *, with_level: bool, verb: bool = False) -> str:
    """A reduced clause -- "lead volume down 72.5% to 91" -- not a sentence.

    Both compressed forms come through here so they cannot disagree with each other or
    with the full sentence about which way the month went. `with_level` is the only
    difference between them: the summary has room to say where the figure landed, an
    SMS mostly does not. `verb` gives the form that can stand as a sentence on its own.
    """
    direction = _direction_word(movement) if verb else _direction_short(movement)
    parts = [label.lower(), direction]
    change = _change_note(movement)
    if change:
        parts.append(change)
    clause = " ".join(parts)
    if with_level:
        clause += f" to {_value(movement, movement.current)}"
    return clause


def _sentence_case(clause: str) -> str:
    return clause[0].upper() + clause[1:] if clause[:1].isalpha() else clause


def _campaign_sentence(label: str, movement: Movement, attribution: Attribution | None) -> str:
    """The month's return, then the campaign that produced most of it, then its share.

    The share is the point. "Campaign return reached $44,183" is a number the reader
    already has; "and one campaign returned 52% of it" is the sentence that tells them
    which campaign to run again, and it is the work the report otherwise leaves them to
    do by hand.
    """
    base = (
        f"{label} reached {_money(movement.current)}, up from {_money(movement.previous)}"
        if movement.current >= movement.previous
        else f"{label} fell to {_money(movement.current)}, down from {_money(movement.previous)}"
    )
    # A campaign that returned nothing is not what produced the month, even when it is
    # the largest of the three the report names. Naming it reads as an accusation and
    # tells the reader nothing they can act on.
    if attribution and attribution.driver and attribution.value:
        share = attribution.share_of_total
        if share is not None and 0 < share <= 100:
            base += f", with {attribution.driver} returning {_money(attribution.value)} of it ({share:.0f}%)"
        else:
            # A campaign can exceed net return when other campaigns offset it. Avoid
            # calling that an impossible share of the total.
            base += f"; {attribution.driver} returned {_money(attribution.value)}"
    return base + "."


def build_findings(model: MonthModel) -> list[Finding]:
    """Rank what the month has to say, by position in the chain rather than by size."""
    findings: list[Finding] = []
    seen: set[str] = set()

    def add(key: str, text: str, improving: bool, order: int, magnitude: float,
            short: str = "", brief: str = "", statement: str = "", hollow: bool = False,
            minor: bool = False, movement: Movement | None = None, compact: str = "") -> None:
        if key in seen or not text:
            return
        seen.add(key)
        brief = brief or short or text.rstrip(".")
        if movement is not None and not compact:
            compact = _compact_clause(key, movement)
        findings.append(Finding(key=key, text=text, improving=improving, order=order,
                                magnitude=magnitude, short=short or text.rstrip("."),
                                brief=brief, statement=statement or brief, hollow=hollow,
                                minor=minor, compact=compact))

    # 1. The bottom line: units and what the campaigns returned.
    bottom = model.bottom_line()
    units_station = model.sales.station("units")
    if units_station and units_station.reported and units_station.movement.moved:
        movement = units_station.movement
        add("units", _station_sentence(units_station), movement.improving,
            ORDER_UNITS, abs(movement.percent_change or 0),
            short=_movement_clause(units_station.label, movement, with_level=True),
            brief=_movement_clause(units_station.label, movement, with_level=True),
            statement=_movement_clause(units_station.label, movement, with_level=True, verb=True),
            movement=movement)

    for key, label, block in (
        ("service_campaign_roi", "Service campaign return", model.service_campaigns),
        ("sales_campaign_roi", "Sales campaign return", model.sales_campaigns),
    ):
        movement = bottom.get(key)
        if movement is None or not movement.moved:
            continue
        attribution = None
        if block is not None and block.runs_campaigns:
            # The model already works out the leader and what share of the stated total
            # it accounts for; rebuilding a thinner Attribution here threw the share
            # away, which is the half of the sentence worth reading.
            attribution = _campaign_driver(block)
        add(key, _campaign_sentence(label, movement, attribution), movement.improving,
            ORDER_CAMPAIGN_RETURN, abs(movement.percent_change or 0),
            short=_movement_clause(label, movement, with_level=True),
            brief=_movement_clause(label, movement, with_level=True),
            statement=_movement_clause(label, movement, with_level=True, verb=True),
            movement=movement)

    # 2. The first leak, and the funnel around it. Earliest before largest: everything
    #    below a station inherits what happens at it. Sales before service, because the
    #    sales chain is where the units the dealer is judged on come from.
    for funnel, base_order in ((model.sales, ORDER_SALES_FUNNEL), (model.service, ORDER_SERVICE_FUNNEL)):
        for position, station in enumerate(funnel.stations):
            if not station.reported or not station.movement.moved:
                continue
            movement = station.movement
            shrunken = _shrunken_base(model, station) if funnel is model.sales else None
            qualifier = f"on {_change_note(shrunken)} fewer leads" if shrunken else ""
            suffix = f" {qualifier}" if qualifier else ""
            add(station.key, _station_sentence(station, qualifier), movement.improving,
                base_order + position, _size(movement),
                short=_movement_clause(station.label, movement, with_level=False)
                + (" on fewer leads" if shrunken else ""),
                brief=_movement_clause(station.label, movement, with_level=True) + suffix,
                statement=_movement_clause(station.label, movement, with_level=True, verb=True) + suffix,
                hollow=shrunken is not None, minor=_is_small_count(movement), movement=movement)

    # 3. A campaign that cost money is always worth a line of its own.
    for block in (model.sales_campaigns, model.service_campaigns):
        if block is None:
            continue
        for campaign in block.losing():
            add(f"losing:{campaign.name}", f"{campaign.name} returned {_money(campaign.roi)}.",
                False, ORDER_LOSING_CAMPAIGN, abs(campaign.roi),
                short=f"{campaign.name} returned {_money(campaign.roi)}",
                brief=f"{campaign.name} returned {_money(campaign.roi)}",
                compact=f"{campaign.name} ROI {_money(campaign.roi)}")

    # 4. Follow-up workload, once the funnel itself has been covered.
    for station in model.hygiene:
        if not station.reported or not station.movement.moved:
            continue
        movement = station.movement
        add(station.key, _station_sentence(station), movement.improving,
            ORDER_HYGIENE, _size(movement),
            short=_movement_clause(station.label, movement, with_level=False),
            brief=_movement_clause(station.label, movement, with_level=True),
            statement=_movement_clause(station.label, movement, with_level=True, verb=True),
            minor=_is_small_count(movement), movement=movement)

    findings.sort(key=lambda f: (f.order, -f.magnitude))
    return findings


def select_insights(model: MonthModel, findings: list[Finding]) -> list[Finding]:
    """Four findings that between them describe the month, not one side of it."""
    chosen = findings[:MAX_INSIGHTS]
    counts = model.breadth()
    if counts["firmed"] and counts["softened"] and len(findings) > MAX_INSIGHTS:
        directions = {f.improving for f in chosen}
        if len(directions) == 1:
            missing = not directions.pop()
            replacement = next((f for f in findings[MAX_INSIGHTS:] if f.improving is missing), None)
            if replacement is not None:
                chosen = chosen[:-1] + [replacement]
    return chosen


VERDICT_OPENERS = {
    "good": "{month} was ahead of {previous} overall",
    "soft": "{month} was a softer month than {previous}",
    "mixed": "{month} was a mixed month",
    "quiet": "{month} was broadly steady",
}

SUMMARY_SENTENCES = 3
SUMMARY_EXTRA_POINTS = 3


def _tier(finding: Finding) -> int:
    """Sales chain, then service chain, then losing campaigns, then follow-up workload."""
    if finding.order < ORDER_SERVICE_FUNNEL:
        return 0
    if finding.order < ORDER_LOSING_CAMPAIGN:
        return 1
    if finding.order < ORDER_HYGIENE:
        return 2
    return 3


def _by_interest(findings: list[Finding]) -> list[Finding]:
    return sorted(findings, key=lambda f: (_tier(f), -f.magnitude, f.order))


def _points_of_interest(chosen: list[Finding], findings: list[Finding]) -> list[Finding]:
    """The most telling movements that did not make the four insights.

    Tiered before sized: a 130% rise in overdue tasks is worth a sentence only when the
    funnel has nothing left to say, not ahead of a 64% fall in leads.
    """
    said = {f.key for f in chosen}
    return _by_interest([f for f in findings if f.key not in said and not f.minor])[:SUMMARY_EXTRA_POINTS]


def _closing_positive(stated: list[Finding], pool: list[Finding]) -> Finding | None:
    """Something that went right, not already said.

    A summary that ends on the worst number of the month is the last thing the reader
    takes away from it, and a month almost always has something working. This finds it
    rather than inventing one: where nothing improved, the summary ends where it ends.
    """
    said = {f.key for f in stated}
    return next((f for f in pool if f.improving and not f.hollow and f.key not in said), None)


def render_summary(model: MonthModel, chosen: list[Finding], month_name: str,
                   previous_name: str, findings: list[Finding] | None = None) -> str:
    """A conversational frame, then the most useful movements not in the lead.

    Use a reduced clause after the opener, so the summary reads as a sentence rather
    than a metrics dump. Two supporting movements follow (outside the bullets where
    possible), good news last; a closing positive is added only when the summary would
    otherwise end on bad news. Never introduced with the repetitive "Elsewhere".
    """
    opener = VERDICT_OPENERS[model.verdict()].format(month=month_name, previous=previous_name)
    if not chosen:
        return f"{opener}."
    wanted = SUMMARY_SENTENCES - 1
    supporting = _points_of_interest(chosen, list(findings or []))[:wanted]
    if len(supporting) < wanted:
        supporting += _by_interest([f for f in chosen[1:] if not f.minor])[:wanted - len(supporting)]
    # Good news last, so the summary does not end on the worst number when it need not.
    supporting.sort(key=lambda f: f.improving and not f.hollow)
    stated = [chosen[0], *supporting]
    sentences = [f"{opener}, with {chosen[0].brief}."]
    sentences.extend(f"{_sentence_case(f.statement)}." for f in supporting)
    if not stated[-1].improving or stated[-1].hollow:
        closer = _closing_positive(stated, list(chosen) + list(findings or []))
        if closer is not None:
            sentences.append(f"{_sentence_case(closer.statement)}.")
    return " ".join(sentences)


SMS_LABELS = {
    "units": "deals",
    "leads": "sales leads",
    "contact": "uncontacted leads",
    "appointments": "appointment rate",
    "test_drives": "test-drive rate",
    "close": "close rate",
    "sales_campaign_roi": "sales ROI",
    "service_campaign_roi": "service ROI",
    "reminders": "service reminders",
    "enrolled": "customers enrolled",
    "bookings": "campaign bookings",
    "sales_overdue_tasks": "overdue sales tasks",
    "sales_email_capture_rate": "email capture",
    "service_outstanding_tasks": "outstanding tasks",
    "service_no_shows": "no-shows",
}
SERVICE_KEYS = {"service_campaign_roi", "reminders", "enrolled", "bookings",
                "service_outstanding_tasks", "service_no_shows"}


def _compact_clause(key: str, movement: Movement) -> str:
    """"deals up (33→40)" for a count; "email capture up to 47%" for a rate, dollars or time."""
    name = SMS_LABELS.get(key, movement.label.lower())
    direction = _direction_short(movement)
    if _is_rate(movement) or movement.metric_type in ("currency", "time") or movement.unit in ("currency", "minutes"):
        return f"{name} {direction} to {_value(movement, movement.current)}"
    return f"{name} {direction} ({_value(movement, movement.previous)}→{_value(movement, movement.current)})"


def _sms_area(finding: Finding) -> int:
    if finding.key.startswith("losing:"):
        return 2
    return 1 if finding.key in SERVICE_KEYS else 0


def render_sms(chosen: list[Finding], month_name: str) -> str:
    """As much of the month as fits, and never one-sided when the month was not.

    Each side's largest movement goes in first; then the most remaining bullets that fit,
    preferring the fuller message on a tie. Laid out as one sentence per area (sales,
    service, losing campaigns), in chain order.
    """
    header = f"DealerSocket {month_name} Insights:\n"
    usable = [f for f in chosen if not f.hollow]
    if not usable:
        return f"{header}No material movement."

    def message(picked: list[Finding]) -> str:
        areas: dict[int, list[str]] = {}
        for f in sorted(picked, key=lambda f: (_sms_area(f), f.order)):
            areas.setdefault(_sms_area(f), []).append(f.compact or f.short)
        body = " ".join(f"{_sentence_case(', '.join(clauses))}." for clauses in areas.values())
        return f"{header}{body}"

    leads = [
        max(side, key=lambda f: f.magnitude)
        for side in ([f for f in usable if f.improving], [f for f in usable if not f.improving])
        if side
    ]
    picked = next(
        (option for option in (leads, *([lead] for lead in leads)) if len(message(option)) <= SMS_LIMIT),
        leads[:1],
    )
    rest = [f for f in usable if f not in picked]
    fitting = (
        [*picked, *extra]
        for size in range(len(rest) + 1)
        for extra in combinations(rest, size)
        if len(message([*picked, *extra])) <= SMS_LIMIT
    )
    # Most facts first, then the fullest message; at most four bullets, so this stays tiny.
    return message(max(fitting, key=lambda option: (len(option), len(message(option))), default=picked))


def build_client_copy(model: MonthModel, month_name: str, previous_name: str) -> tuple[list[str], str, str]:
    """The four insights, the summary and the SMS — three views of one model."""
    findings = build_findings(model)
    chosen = select_insights(model, findings)
    bullets = [finding.text for finding in chosen]
    summary = render_summary(model, chosen, month_name, previous_name, findings)
    return bullets, summary, render_sms(chosen, month_name)
