"""The month model: what happened, computed before anything is written.

Every fixture here is synthetic. The point of the model is that the layer above reads
it instead of re-reading prose, so what these tests pin is the contract that layer will
depend on — ordering, ownership, attribution, and the verdict rule.
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from reporting import month_model as mm
from reporting.report_breakdowns import Campaign, CampaignBlock, SourceRow


def metric(previous: float, current: float, band: str = "strong", higher_is_better: bool = True) -> dict:
    return {
        "usable_for_reporting": True,
        "previous_value": float(previous),
        "current_value": float(current),
        "percent_change": round((current - previous) / previous * 100, 2) if previous else 0.0,
        "direction": "up" if current > previous else "down",
        "unit": "count",
        "metric_type": "volume",
        "significance_band": band,
        "higher_is_better": higher_is_better,
    }


def build(metrics=None, breakdowns=None) -> mm.MonthModel:
    return mm.build_month_model(
        dealership="Demo Site 1", month_label="Jul 2026",
        metrics=metrics or {}, breakdowns=breakdowns or {},
    )


class FunnelOrderTests(unittest.TestCase):
    def test_the_first_leak_is_the_earliest_not_the_largest(self) -> None:
        """A fall at the top of the funnel is inherited by everything below it, so
        naming the downstream symptom as the headline points at the wrong place."""
        model = build({
            "sales_total_leads": metric(300, 240),          # -20%, upstream
            "sales_total_deals": metric(150, 60),           # -60%, downstream and bigger
        })
        self.assertEqual(model.sales.largest_leak().key, "leads")

    def test_an_improving_station_is_never_the_leak(self) -> None:
        model = build({
            "sales_total_leads": metric(240, 300),
            "sales_total_deals": metric(150, 60),
        })
        self.assertEqual(model.sales.largest_leak().key, "units")

    def test_a_month_with_no_adverse_station_has_no_leak(self) -> None:
        self.assertIsNone(build({"sales_total_leads": metric(240, 300)}).sales.largest_leak())

    def test_the_diagnostic_pair_is_appointments_and_test_drives(self) -> None:
        """The operator's rule: healthy here means the floor is working its leads."""
        model = build()
        self.assertEqual([s.key for s in model.diagnostic_pair()], ["appointments", "test_drives"])


class OwnershipTests(unittest.TestCase):
    def test_the_funnel_belongs_to_the_salesperson_and_campaigns_do_not(self) -> None:
        """Two levers, two owners. A marketing problem reported as a sales problem is
        advice pointed at the wrong person."""
        model = build()
        self.assertTrue(all(s.owner == mm.SALESPERSON for s in model.sales.stations))
        self.assertTrue(all(s.owner == mm.BUSINESS for s in model.service.stations))

    def test_response_time_is_not_a_station_of_its_own(self) -> None:
        """Response time is measured from when a lead is actioned, so uncontacted leads
        are the response-time figure. Two stations would count one problem twice."""
        keys = [s.key for s in build().sales.stations]
        self.assertIn("contact", keys)
        self.assertNotIn("response_time", keys)
        self.assertEqual(
            [s[2] for s in mm.SALES_STATIONS if s[0] == "contact"], ["sales_uncontacted_leads"]
        )

    def test_the_response_time_metric_is_kept_as_evidence_for_contact(self) -> None:
        model = build({
            "sales_uncontacted_leads": metric(10, 40),
            "sales_internet_response_time": metric(50, 90, higher_is_better=False),
        })
        self.assertEqual(model.sales.station("contact").note, "response_time_follows_contact")


class AttributionTests(unittest.TestCase):
    SOURCES = [
        SourceRow("Walk In", 26, 21, 10, 3, 38.0, 14.0),
        SourceRow("Internet", 18, 27, 0, 5, 0.0, 19.0),
        SourceRow("Other", 276, 34, 156, 100, 57.0, 294.0),
        SourceRow("Total", 331, 91, 176, 114, 53.0, 125.0),
    ]

    def test_the_biggest_mover_drives_the_total(self) -> None:
        model = build({"sales_total_leads": metric(331, 91)}, {"lead_sources": self.SOURCES})
        attribution = model.sales.station("leads").attribution
        self.assertEqual(attribution.driver, "Other")
        self.assertEqual(attribution.value, -242)

    def test_other_is_carried_at_full_weight(self) -> None:
        """It stands for whichever channels a dealership set up beyond the named three,
        so it varies between dealerships and is not a residual to discount."""
        model = build({"sales_total_leads": metric(331, 91)}, {"lead_sources": self.SOURCES})
        self.assertEqual(model.sales.station("leads").attribution.driver, "Other")

    def test_the_total_row_is_never_its_own_driver(self) -> None:
        model = build({"sales_total_deals": metric(176, 114)}, {"lead_sources": self.SOURCES})
        self.assertNotEqual(model.sales.station("units").attribution.driver, "Total")

    def test_a_source_moving_against_the_total_is_not_its_driver(self) -> None:
        """A summary read "Deal volume rose ..., with Walk-in accounting for 5 of the drop"."""
        sources = [
            SourceRow("Walk-in", 60, 50, 10, 5, None, None),
            SourceRow("Internet", 80, 60, 5, 9, None, None),
            SourceRow("Other", 32, 18, 5, 8, None, None),
        ]
        model = build({"sales_total_deals": metric(20, 22)}, {"lead_sources": sources})
        attribution = model.sales.station("units").attribution
        self.assertEqual((attribution.driver, attribution.value), ("Internet", 4))

    def test_a_campaign_driver_reports_its_share_and_what_is_unnamed(self) -> None:
        """Never a decomposition: the report names three campaigns out of however many."""
        block = CampaignBlock(
            named=[Campaign("Closed RO", 28, 18201), Campaign("New Closed RO", 6, 4891)],
            stated_count=69, stated_roi=37163, previous_count=51, previous_roi=30110,
        )
        model = build({}, {"service_campaigns": block})
        attribution = model.service.station("bookings").attribution
        self.assertEqual(attribution.driver, "Closed RO")
        self.assertEqual(attribution.share_of_total, 49.0)
        self.assertEqual(attribution.unnamed, 35)

    def test_no_campaigns_means_no_driver(self) -> None:
        model = build({}, {"service_campaigns": CampaignBlock()})
        self.assertIsNone(model.service.station("bookings").attribution)


class GatedButExplainedTests(unittest.TestCase):
    def test_a_gated_metric_with_named_campaigns_behind_it_is_flagged(self) -> None:
        """Bookings jumping sevenfold is held back as suspicious. The campaign block
        names the campaigns that account for exactly that jump, which is the point of
        reading it — so the model records the corroboration and lets the layer above
        decide whether the number is now safe to print."""
        gated = metric(7, 65)
        gated["is_suspicious"] = True
        block = CampaignBlock(named=[Campaign("Closed RO", 32, 23114)], stated_count=65, stated_roi=44183,
                              previous_count=7, previous_roi=10059)
        model = build({"service_bookings_from_campaigns": gated}, {"service_campaigns": block})
        station = model.service.station("bookings")
        self.assertFalse(station.reported)
        self.assertEqual(station.note, "gated_metric_corroborated_by_named_campaigns")


class BottomLineTests(unittest.TestCase):
    def test_a_lane_that_returned_nothing_in_either_month_is_not_an_improvement(self) -> None:
        """$0 to $0 marked as a strong gain is how a dead campaign lane reaches copy."""
        block = CampaignBlock(named=[], stated_count=0, stated_roi=0, previous_count=0, previous_roi=0)
        self.assertIsNone(build({}, {"service_campaigns": block}).bottom_line()["service_campaign_roi"])

    def test_campaign_return_is_measured_off_the_stated_totals(self) -> None:
        block = CampaignBlock(named=[], stated_count=65, stated_roi=44183, previous_count=7, previous_roi=10059)
        movement = build({}, {"service_campaigns": block}).bottom_line()["service_campaign_roi"]
        self.assertEqual((movement.previous, movement.current), (10059.0, 44183.0))
        self.assertTrue(movement.improving)

    def test_a_return_off_a_zero_base_is_still_the_story(self) -> None:
        block = CampaignBlock(named=[], stated_count=12, stated_roi=9000, previous_count=0, previous_roi=0)
        movement = build({}, {"service_campaigns": block}).bottom_line()["service_campaign_roi"]
        self.assertEqual(movement.band, "strong")
        self.assertIsNone(movement.percent_change)


class VerdictTests(unittest.TestCase):
    """A bad month is one where the indicators softened across the board, judged against
    the bottom line — not a weighted score over whichever metrics happened to parse."""

    def test_everything_softening_is_a_bad_month(self) -> None:
        model = build({
            "sales_total_leads": metric(300, 200),
            "sales_total_deals": metric(150, 90),
            "sales_new_leads_to_appointments": metric(9, 4),
        })
        self.assertEqual(model.verdict(), "soft")

    def test_everything_firming_is_a_good_month(self) -> None:
        model = build({
            "sales_total_leads": metric(200, 300),
            "sales_total_deals": metric(90, 150),
        })
        self.assertEqual(model.verdict(), "good")

    def test_an_even_split_is_mixed_not_bad(self) -> None:
        """Half up and half down is not 'across the board'."""
        model = build({
            "sales_total_leads": metric(300, 200),
            "sales_total_deals": metric(150, 90),
            "sales_new_leads_to_appointments": metric(4, 9),
            "sales_reminders": metric(1, 2),
            "service_reminders_sent": metric(70, 300),
        })
        self.assertEqual(model.verdict(), "mixed")

    def test_units_holding_while_the_funnel_slips_is_not_a_pass(self) -> None:
        """The following month is already in trouble, so it must not read as good."""
        model = build({
            "sales_total_deals": metric(150, 152, band="slight"),
            "sales_total_leads": metric(300, 200),
            "sales_new_leads_to_appointments": metric(9, 4),
            "sales_appointment_to_sales": metric(10, 6),
        })
        self.assertEqual(model.verdict(), "soft")

    def test_a_month_with_nothing_material_is_quiet(self) -> None:
        self.assertEqual(build({"sales_total_leads": metric(300, 303, band="negligible")}).verdict(), "quiet")

    def test_a_scoreboard_that_disagrees_with_itself_is_split(self) -> None:
        """Units up and campaign return down cannot be "ahead across the main measures".

        Five of twenty-three real dealerships read this way in one month: an opener
        claiming a direction that the very next sentence contradicted.
        """
        model = build(
            {
                "sales_total_leads": metric(300, 380),
                "sales_total_deals": metric(116, 180),
                "sales_new_leads_to_appointments": metric(4, 9),
                "sales_appointment_to_sales": metric(6, 10),
            },
            {"service_campaigns": CampaignBlock(named=[], stated_roi=22155, previous_roi=37163)},
        )
        self.assertEqual(model.bottom_line_directions(), {True, False})
        self.assertEqual(model.verdict(), "mixed")

    def test_a_unanimous_scoreboard_still_reads_the_room(self) -> None:
        """One campaign rising cannot carry a month while the funnel slips underneath."""
        model = build(
            {
                "sales_total_leads": metric(300, 200),
                "sales_new_leads_to_appointments": metric(9, 4),
                "sales_appointment_to_sales": metric(10, 6),
                "sales_email_capture_rate": metric(80, 66),
            },
            {"sales_campaigns": CampaignBlock(named=[], stated_roi=13286, previous_roi=10710)},
        )
        self.assertEqual(model.bottom_line_directions(), {True})
        self.assertEqual(model.verdict(), "mixed")

    def test_an_even_funnel_split_counts_against_the_scoreboard(self) -> None:
        """Half the month moving the other way is not "across the board" either."""
        model = build({
            "sales_total_deals": metric(90, 150),
            "sales_total_leads": metric(200, 300),
            "sales_new_leads_to_appointments": metric(9, 4),
            "sales_appointment_to_sales": metric(10, 6),
        })
        self.assertEqual(model.breadth(), {"firmed": 2, "softened": 2})
        self.assertEqual(model.verdict(), "mixed")

    def test_a_unanimous_scoreboard_with_the_funnel_behind_it_still_reads_clean(self) -> None:
        model = build({
            "sales_total_deals": metric(90, 150),
            "sales_total_leads": metric(200, 300),
            "sales_new_leads_to_appointments": metric(4, 9),
            "sales_appointment_to_sales": metric(6, 10),
        })
        self.assertEqual(model.verdict(), "good")

    def test_a_month_where_only_the_campaigns_moved_is_not_quiet(self) -> None:
        model = build(
            {"sales_total_leads": metric(300, 303, band="negligible")},
            {"sales_campaigns": CampaignBlock(named=[], stated_roi=40000, previous_roi=10000)},
        )
        self.assertEqual(model.verdict(), "good")

    def test_breadth_ignores_trivial_moves(self) -> None:
        model = build({
            "sales_total_leads": metric(300, 303, band="slight"),
            "sales_total_deals": metric(150, 90),
        })
        self.assertEqual(model.breadth(), {"firmed": 0, "softened": 1})


class OneFigureIsOneStationTests(unittest.TestCase):
    """Reminders sent and customers enrolled print separately and are the same number.

    Across a real month of 23 dealerships they were identical in all 23. Two stations
    carrying one figure state the same fact twice and count it twice in breadth(),
    which then leans the verdict off a metric the report only has once.
    """

    def test_the_duplicate_station_is_dropped_when_the_figures_agree(self) -> None:
        model = build({
            "service_reminders_sent": metric(346, 288),
            "service_customers_enrolled": metric(346, 288),
        })
        keys = [s.key for s in model.service.stations]
        self.assertIn("reminders", keys)
        self.assertNotIn("enrolled", keys)

    def test_the_figure_is_counted_once_in_breadth(self) -> None:
        model = build({
            "service_reminders_sent": metric(346, 288),
            "service_customers_enrolled": metric(346, 288),
        })
        self.assertEqual(model.breadth(), {"firmed": 0, "softened": 1})

    def test_both_are_kept_when_the_report_really_does_differ(self) -> None:
        """The collapse is a comparison, not an assumption."""
        model = build({
            "service_reminders_sent": metric(346, 288),
            "service_customers_enrolled": metric(300, 210),
        })
        keys = [s.key for s in model.service.stations]
        self.assertIn("reminders", keys)
        self.assertIn("enrolled", keys)

    def test_enrolled_survives_on_its_own(self) -> None:
        model = build({"service_customers_enrolled": metric(300, 210)})
        self.assertIn("enrolled", [s.key for s in model.service.stations])


class MissingDataTests(unittest.TestCase):
    def test_stations_with_nothing_usable_are_named(self) -> None:
        model = build({"sales_total_leads": metric(300, 200)})
        self.assertNotIn("leads", model.missing())
        self.assertIn("units", model.missing())

    def test_an_empty_month_produces_a_model_rather_than_an_error(self) -> None:
        model = build()
        self.assertEqual(model.verdict(), "quiet")
        self.assertEqual(model.breadth(), {"firmed": 0, "softened": 0})


if __name__ == "__main__":
    unittest.main()
