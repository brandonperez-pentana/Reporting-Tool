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
    # "339.2%" is false precision on a jump of more than four times.
    if abs(movement.percent_change) >= 100:
        return f"{abs(movement.percent_change):.0f}%"
    return f"{abs(movement.percent_change):.1f}%"


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
        base += f", with {attribution.driver} returning {_money(attribution.value)} of it"
        if attribution.share_of_total:
            base += f" ({attribution.share_of_total:.0f}%)"
    return base + "."


def build_findings(model: MonthModel) -> list[Finding]:
    """Rank what the month has to say, by position in the chain rather than by size."""
    findings: list[Finding] = []
    seen: set[str] = set()

    def add(key: str, text: str, improving: bool, order: int, magnitude: float,
            short: str = "", brief: str = "", statement: str = "", hollow: bool = False) -> None:
        if key in seen or not text:
            return
        seen.add(key)
        brief = brief or short or text.rstrip(".")
        findings.append(Finding(key=key, text=text, improving=improving, order=order,
                                magnitude=magnitude, short=short or text.rstrip("."),
                                brief=brief, statement=statement or brief, hollow=hollow))

    # 1. The bottom line: units and what the campaigns returned.
    bottom = model.bottom_line()
    units_station = model.sales.station("units")
    if units_station and units_station.reported and units_station.movement.moved:
        movement = units_station.movement
        add("units", _station_sentence(units_station), movement.improving,
            ORDER_UNITS, abs(movement.percent_change or 0),
            short=_movement_clause(units_station.label, movement, with_level=True),
            brief=_movement_clause(units_station.label, movement, with_level=True),
            statement=_movement_clause(units_station.label, movement, with_level=True, verb=True))

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
            statement=_movement_clause(label, movement, with_level=True, verb=True))

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
                base_order + position, abs(movement.percent_change or 0),
                short=_movement_clause(station.label, movement, with_level=False)
                + (" on fewer leads" if shrunken else ""),
                brief=_movement_clause(station.label, movement, with_level=True) + suffix,
                statement=_movement_clause(station.label, movement, with_level=True, verb=True) + suffix,
                hollow=shrunken is not None)

    # 3. A campaign that cost money is always worth a line of its own.
    for block in (model.sales_campaigns, model.service_campaigns):
        if block is None:
            continue
        for campaign in block.losing():
            add(f"losing:{campaign.name}", f"{campaign.name} returned {_money(campaign.roi)}.",
                False, ORDER_LOSING_CAMPAIGN, abs(campaign.roi),
                short=f"{campaign.name} returned {_money(campaign.roi)}",
                brief=f"{campaign.name} returned {_money(campaign.roi)}")

    # 4. Follow-up workload, once the funnel itself has been covered.
    for station in model.hygiene:
        if not station.reported or not station.movement.moved:
            continue
        movement = station.movement
        add(station.key, _station_sentence(station), movement.improving,
            ORDER_HYGIENE, abs(movement.percent_change or 0),
            short=_movement_clause(station.label, movement, with_level=False),
            brief=_movement_clause(station.label, movement, with_level=True),
            statement=_movement_clause(station.label, movement, with_level=True, verb=True))

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
    "good": "{month} finished ahead of {previous} across the main measures",
    "soft": "{month} came in behind {previous} across the main measures",
    "mixed": "{month} moved in both directions",
    "quiet": "{month} moved little on the main measures",
}

# When some measures moved against the verdict, "across the main measures" overclaims.
QUALIFIED_OPENERS = {
    "good": ("softened", "{month} finished ahead of {previous} overall"),
    "soft": ("firmed", "{month} came in behind {previous} overall"),
}

SUMMARY_SENTENCES = 3


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
    """A factual frame, then the same movements in short form. No mood, no advice.

    Two things this has to avoid. It cannot splice a finished sentence onto the frame
    with "with", which produced "July was split, with reminders sent rose from 70 to
    300"; and it cannot simply repeat the bullets, which puts the same three sentences
    in front of the reader twice in one email. So it takes the reduced clause instead:
    the frame absorbs the first, and the rest stand alone. The drivers stay in the
    bullets, which is the only place they are worth the words.
    """
    verdict = model.verdict()
    template = VERDICT_OPENERS[verdict]
    if verdict in QUALIFIED_OPENERS:
        against, qualified = QUALIFIED_OPENERS[verdict]
        if model.breadth()[against]:
            template = qualified
    opener = template.format(month=month_name, previous=previous_name)
    if not chosen:
        return f"{opener}."
    stated = chosen[:SUMMARY_SENTENCES]
    sentences = [f"{opener}, with {stated[0].brief}."]
    sentences.extend(f"{_sentence_case(f.statement)}." for f in stated[1:])
    if not stated[-1].improving or stated[-1].hollow:
        closer = _closing_positive(stated, list(chosen) + list(findings or []))
        if closer is not None:
            sentences.append(f"{_sentence_case(closer.statement)}.")
    return " ".join(sentences)


def render_sms(chosen: list[Finding], month_name: str) -> str:
    """Under the limit, and never one-sided when the month was not."""
    positives = [f for f in chosen if f.improving and not f.hollow]
    negatives = [f for f in chosen if not f.improving]
    if positives and negatives:
        body = f"{positives[0].short}, but {negatives[0].short}"
        if len(negatives) > 1 and len(body) + len(negatives[1].short) + 6 < SMS_LIMIT - 30:
            body += f" and {negatives[1].short}"
    else:
        parts = [f.short for f in chosen[:2]]
        body = " and ".join(parts) if parts else "no material movement"
    message = f"DealerSocket {month_name} Insights:\n{_sentence_case(body)}."
    if len(message) > SMS_LIMIT:
        message = f"DealerSocket {month_name} Insights:\n{_sentence_case(chosen[0].short)}."
    return message


def build_client_copy(model: MonthModel, month_name: str, previous_name: str) -> tuple[list[str], str, str]:
    """The four insights, the summary and the SMS — three views of one model."""
    findings = build_findings(model)
    chosen = select_insights(model, findings)
    bullets = [finding.text for finding in chosen]
    summary = render_summary(model, chosen, month_name, previous_name, findings)
    return bullets, summary, render_sms(chosen, month_name)
