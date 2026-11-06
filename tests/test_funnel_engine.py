"""The funnel engine as it reaches a dealership: the switch, and what survives it.

The layer tests next door pin the model and the copy in isolation. What this file pins
is the part that actually decides what a dealer reads -- whether the switch is off by
default, whether the copy the funnel engine wrote arrives intact, and whether a failure
inside it costs a month.

Each test corresponds to something that went wrong while wiring it in. The copy was
grammatical when funnel_insights produced it and ungrammatical by the time it reached
the email; the four insights were ordered by the model and re-ordered afterwards; the
SMS the model composed was thrown away and rebuilt by reading the bullets back as
prose. All three had the same cause: the packager that finishes rule-engine output was
still running over funnel output, and that packager works by re-reading rendered text.
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import dealership_insights_local as core
import settings_manager
from reporting import funnel_insights as fi
from reporting import month_model as mm
from reporting import report_breakdowns as rb

DEALERSHIP = "Demo Site 1"
MONTH = "Jul 2026"

# A month built here rather than read from a produced run. The package ships with no
# reports and no debug text, so a fixture on disk would make these tests skip in exactly
# the install where the swap most needs verifying. Every figure below is invented.
FIGURES = [
    # metric,                           previous,  current
    ("sales_total_leads",                  "331",     "91"),
    ("sales_uncontacted_leads",             "12",     "44"),
    ("sales_new_leads_to_appointments",   "4.0%",   "9.0%"),
    ("sales_new_leads_to_test_drives",    "3.0%",   "5.0%"),
    ("sales_appointment_to_sales",       "40.0%",  "31.0%"),
    ("sales_total_deals",                  "176",    "114"),
    ("sales_overdue_tasks",                 "40",     "55"),
    ("sales_email_capture_rate",         "75.0%",  "69.0%"),
    ("sales_pipeline_size",                "220",    "180"),
    ("service_reminders_sent",              "70",    "300"),
    ("service_customers_enrolled",          "40",    "150"),
    ("service_outstanding_tasks",            "6",     "22"),
    ("service_no_shows",                    "10",      "4"),
]

# Only what the breakdown parser reads. funnel_client_copy refuses to run on a month
# with no report text at all, which is itself pinned below.
SALES_TEXT = """Lead Source Performance
Source          Leads   Deals   Conversion
Website           120      70          58%
Autotrader         90      40          44%
Other             121      66          55%
Total             331     176          53%
"""
SERVICE_TEXT = "Service Campaign Performance\nNo campaigns run this month.\n"


def reports() -> tuple[str, str]:
    return SALES_TEXT, SERVICE_TEXT


def metrics_for_the_month() -> dict:
    metrics: dict = {}
    for name, previous, current in FIGURES:
        core.add_metric(metrics, name, previous, current)
    return metrics


class EngineSwitch:
    """Force the engine setting for one block, without touching the operator's file."""

    def __init__(self, engine: str) -> None:
        self.engine = engine
        self.original = settings_manager.load_settings

    def __enter__(self):
        original = self.original
        engine = self.engine

        def loader(*args, **kwargs):
            data = dict(original(*args, **kwargs))
            data["insight_engine"] = engine
            return data

        settings_manager.load_settings = loader
        return self

    def __exit__(self, *exc) -> None:
        settings_manager.load_settings = self.original


def run(engine: str):
    sales, service = reports()
    with EngineSwitch(engine):
        return core.generate_insights_from_metrics(
            metrics_for_the_month(), "success", DEALERSHIP, MONTH,
            sales_text=sales, service_text=service,
        )


def model_for() -> mm.MonthModel:
    sales, service = reports()
    return mm.build_month_model(
        dealership=DEALERSHIP, month_label=MONTH,
        metrics=metrics_for_the_month(),
        breakdowns=rb.parse_breakdowns(sales, service),
    )


class DefaultsToTheRuleEngineTests(unittest.TestCase):
    def test_the_shipped_default_is_the_rule_engine(self) -> None:
        """Nothing changes for an operator who has not asked for it."""
        self.assertEqual(settings_manager.DEFAULT_SETTINGS.get("insight_engine"), "rules")

    def test_an_unreadable_settings_file_does_not_switch_engines(self) -> None:
        original = settings_manager.load_settings
        settings_manager.load_settings = lambda *a, **k: (_ for _ in ()).throw(OSError("gone"))
        try:
            self.assertFalse(core.funnel_engine_selected())
        finally:
            settings_manager.load_settings = original

    def test_the_rule_engine_is_not_marked_as_the_funnel(self) -> None:
        result = run("rules")
        self.assertNotIn("funnel_engine", result[7])
        self.assertEqual(result[3], ["priority_insights"])


class FunnelCopyArrivesIntactTests(unittest.TestCase):
    """The packager must not run over funnel output.

    ensure_four_priority_insights re-ranks by reading the rendered bullets, and
    build_sms_from_priority_insights composes the message the same way. Both undo
    decisions the model already made from figures, which is the whole reason this
    engine exists.
    """

    def setUp(self) -> None:
        self.result = run("funnel")
        self.expected = fi.build_client_copy(model_for(), "July", "June")

    def test_the_run_is_marked_as_the_funnel_engine(self) -> None:
        self.assertIn("funnel_engine", self.result[7])
        self.assertEqual(self.result[3], ["funnel_engine"])

    def test_the_bullets_are_the_ones_the_model_chose_in_that_order(self) -> None:
        self.assertEqual(self.result[0], self.expected[0])

    def test_the_summary_is_the_one_the_model_wrote(self) -> None:
        self.assertEqual(self.result[1], self.expected[1])

    def test_the_sms_is_the_one_the_model_composed(self) -> None:
        self.assertEqual(self.result[2], self.expected[2])

    def test_the_sms_still_fits(self) -> None:
        self.assertLessEqual(len(self.result[2]), core.SMS_CHAR_LIMIT)

    def test_a_funnel_report_has_four_priority_lines(self) -> None:
        self.assertEqual(len(self.result[0]), core.PRIORITY_INSIGHT_COUNT)

    def test_sparse_but_real_metrics_are_padded_with_an_explanation(self) -> None:
        metrics = {}
        core.add_metric(metrics, "sales_total_deals", "10", "12")
        core.add_metric(metrics, "sales_total_leads", "100", "120")
        copy = core.funnel_client_copy(
            metrics, DEALERSHIP, MONTH, "sales data", "service data", breakdowns={}
        )
        self.assertIsNotNone(copy)
        bullets, _summary, sms, _verdict = copy
        self.assertEqual(len(bullets), core.PRIORITY_INSIGHT_COUNT)
        self.assertTrue(any("No further" in bullet for bullet in bullets))
        self.assertLessEqual(len(sms), core.SMS_CHAR_LIMIT)

    def test_the_month_rating_agrees_with_the_summary_verdict(self) -> None:
        """"Softer" above a summary that opens "moved in both directions" is a contradiction."""
        expected = core.FUNNEL_VERDICT_RATINGS[model_for().verdict()]
        self.assertEqual(self.result[13].final_rating_used, expected)


class AFailureCostsNothingTests(unittest.TestCase):
    def test_a_broken_model_falls_back_to_the_rule_engine(self) -> None:
        original = mm.build_month_model
        mm.build_month_model = lambda **kwargs: (_ for _ in ()).throw(ValueError("boom"))
        try:
            result = run("funnel")
        finally:
            mm.build_month_model = original
        self.assertNotIn("funnel_engine", result[7])
        self.assertEqual(result[0], run("rules")[0])

    def test_reports_that_could_not_be_read_leave_the_rule_engine_alone(self) -> None:
        self.assertIsNone(core.funnel_client_copy({}, DEALERSHIP, MONTH, "", ""))


class SummaryReadsAsProseTests(unittest.TestCase):
    """The frame is a sentence and the findings are sentences; joining them needs care.

    "July was split, with reminders sent rose from 70 to 300" shipped because the lead
    finding was spliced in after "with" while still carrying its own finite verb.
    """

    def setUp(self) -> None:
        self.bullets, self.summary, _ = fi.build_client_copy(model_for(), "July", "June")

    def test_the_frame_does_not_swallow_a_finished_sentence(self) -> None:
        for verb in (" rose ", " fell ", " reached "):
            self.assertNotIn(f", with {verb.strip()}", self.summary)
            self.assertNotIn(f"with {self.bullets[0][0].lower()}{self.bullets[0][1:].rstrip('.')}", self.summary)

    def test_the_summary_does_not_repeat_a_bullet_word_for_word(self) -> None:
        """The reader gets the bullets underneath; saying it twice wastes the email."""
        for bullet in self.bullets:
            self.assertNotIn(bullet, self.summary)
        self.assertNotIn("Elsewhere,", self.summary)

    def test_the_summary_opens_in_a_conversational_voice(self) -> None:
        self.assertTrue(self.summary.startswith("July was a softer month than June, with "), self.summary)

    def test_a_good_month_with_a_decline_is_not_ahead_across_the_main_measures(self) -> None:
        """A positive month can still contain a decline in lead volume."""
        metrics: dict = {}
        for name, previous, current in (("sales_total_deals", "20", "22"),
                                        ("sales_total_leads", "172", "128"),
                                        ("sales_new_leads_to_test_drives", "16.0%", "23.0%"),
                                        ("sales_new_leads_to_appointments", "15.0%", "21.0%")):
            core.add_metric(metrics, name, previous, current)
        model = mm.build_month_model(dealership=DEALERSHIP, month_label=MONTH, metrics=metrics, breakdowns={})
        self.assertEqual(model.verdict(), "good")
        _, summary, _ = fi.build_client_copy(model, "August", "July")
        self.assertTrue(summary.startswith("August was ahead of July overall"), summary)

    def test_every_sentence_is_closed(self) -> None:
        self.assertTrue(self.summary.endswith("."))
        self.assertNotIn("..", self.summary)


class TheSummaryClosesOnSomethingWorkingTests(unittest.TestCase):
    """The last sentence is what the reader takes away.

    Ending on the worst number of the month leaves them there, and a month almost
    always has something working. Almost -- so this finds one rather than inventing it.
    """

    def parts(self, metrics: dict):
        model = mm.build_month_model(dealership=DEALERSHIP, month_label=MONTH,
                                     metrics=metrics, breakdowns={})
        findings = fi.build_findings(model)
        return model, findings, fi.select_insights(model, findings)

    def summary_for(self, metrics: dict) -> str:
        model, findings, chosen = self.parts(metrics)
        return fi.render_summary(model, chosen, "July", "June", findings)

    def figures(self, **pairs) -> dict:
        metrics: dict = {}
        for name, (previous, current) in pairs.items():
            core.add_metric(metrics, name, previous, current)
        return metrics

    def test_a_soft_month_still_ends_on_what_worked(self) -> None:
        summary = self.summary_for(self.figures(
            sales_total_deals=("176", "114"),
            sales_total_leads=("331", "240"),
            sales_appointment_to_sales=("40.0%", "31.0%"),
            sales_new_leads_to_appointments=("4.0%", "9.0%"),
        ))
        self.assertTrue(summary.rstrip(".").endswith("rose 5 pts to 9%"), summary)

    def test_it_does_not_invent_one_when_nothing_improved(self) -> None:
        summary = self.summary_for(self.figures(
            sales_total_deals=("176", "114"),
            sales_total_leads=("331", "240"),
            sales_appointment_to_sales=("40.0%", "31.0%"),
        ))
        self.assertNotIn("up ", summary)

    def test_a_month_already_ending_well_is_left_alone(self) -> None:
        """Nothing is appended when the third sentence is already good news."""
        metrics = self.figures(
            sales_total_deals=("114", "176"),
            sales_total_leads=("240", "331"),
            sales_new_leads_to_appointments=("4.0%", "9.0%"),
        )
        model, findings, chosen = self.parts(metrics)
        self.assertEqual(
            fi.render_summary(model, chosen, "July", "June", findings),
            fi.render_summary(model, chosen, "July", "June"),
        )


class TheSummaryCarriesWhatTheBulletsCouldNotTests(unittest.TestCase):
    """Four bullets fill up on the chain; a 130% rise in outstanding tasks then had
    nowhere to go, because the summary only restated the bullets."""

    def copy(self, **extra):
        metrics: dict = {}
        for name, previous, current in (
            ("sales_total_deals", "115", "55"),
            ("sales_total_leads", "354", "127"),
            ("sales_new_leads_to_appointments", "24.0%", "6.0%"),
            ("sales_new_leads_to_test_drives", "36.0%", "31.0%"),
            ("sales_appointment_to_sales", "40.0%", "31.0%"),
            ("service_outstanding_tasks", "300", "690"),
            *((name, *pair) for name, pair in extra.items()),
        ):
            core.add_metric(metrics, name, previous, current)
        model = mm.build_month_model(dealership=DEALERSHIP, month_label=MONTH,
                                     metrics=metrics, breakdowns={})
        findings = fi.build_findings(model)
        bullets, summary, _ = fi.build_client_copy(model, "July", "June")
        return findings, bullets, summary

    def test_a_large_movement_outside_the_bullets_reaches_the_summary(self) -> None:
        _, bullets, summary = self.copy()
        self.assertFalse(any("Outstanding service tasks" in b for b in bullets))
        self.assertIn("Outstanding service tasks rose 130% to 690.", summary)

    def test_the_summary_does_not_restate_the_other_bullets(self) -> None:
        _, _, summary = self.copy()
        self.assertNotIn("Lead volume fell", summary)
        self.assertNotIn("Lead-to-appointment conversion fell", summary)

    def test_a_tiny_count_is_not_a_point_of_interest(self) -> None:
        """1 -> 4 no-shows is "up 300%" and still four people."""
        findings, _, summary = self.copy(service_no_shows=("1", "4"))
        no_shows = next((f for f in findings if f.key == "service_no_shows"), None)
        if no_shows is not None:
            self.assertTrue(no_shows.minor)
        self.assertNotIn("no-shows", summary.lower())


class TheSummaryPicksWhatMattersTests(unittest.TestCase):
    @staticmethod
    def finding(key: str, improving: bool, order: int, magnitude: float) -> fi.Finding:
        statement = f"{key} {'rose' if improving else 'fell'}"
        return fi.Finding(key=key, text=statement + ".", improving=improving, order=order,
                          magnitude=magnitude, short=statement, brief=statement, statement=statement)

    def summary(self, chosen, findings) -> str:
        model = mm.build_month_model(dealership=DEALERSHIP, month_label=MONTH, metrics={}, breakdowns={})
        return fi.render_summary(model, chosen, "July", "June", [*chosen, *findings])

    def test_the_funnel_outranks_follow_up_workload(self) -> None:
        chosen = [self.finding("units", False, fi.ORDER_UNITS, 52)]
        summary = self.summary(chosen, [
            self.finding("overdue", False, fi.ORDER_HYGIENE, 130),
            self.finding("leads", False, fi.ORDER_SALES_FUNNEL, 64),
            self.finding("close", False, fi.ORDER_SALES_FUNNEL + 4, 61),
        ])
        self.assertIn("Leads fell.", summary)
        self.assertIn("Close fell.", summary)
        self.assertNotIn("Overdue", summary)

    def test_good_news_comes_last(self) -> None:
        chosen = [self.finding("units", True, fi.ORDER_UNITS, 20)]
        summary = self.summary(chosen, [
            self.finding("bookings", True, fi.ORDER_SERVICE_FUNNEL, 80),
            self.finding("appointments", False, fi.ORDER_SALES_FUNNEL + 2, 10),
        ])
        self.assertTrue(summary.endswith("Bookings rose."), summary)

    def test_a_thin_month_borrows_from_the_bullets(self) -> None:
        chosen = [
            self.finding("units", False, fi.ORDER_UNITS, 11),
            self.finding("test_drives", True, fi.ORDER_SALES_FUNNEL + 3, 18),
            self.finding("overdue", True, fi.ORDER_HYGIENE, 50),
        ]
        summary = self.summary(chosen, [self.finding("email", True, fi.ORDER_HYGIENE, 11)])
        self.assertIn("Test_drives rose.", summary)
        self.assertIn("Email rose.", summary)

    def test_a_rate_is_sized_in_points(self) -> None:
        movement = mm.Movement(
            metric="sales_new_leads_to_appointments", label="Lead-to-appointment conversion",
            previous=2.0, current=8.0, percent_change=300.0, improving=True, usable=True,
            band="strong", unit="percentage", metric_type="rate", point_change=6.0,
        )
        self.assertEqual(fi._size(movement), 6.0)


class ARateLiftedByFewerLeadsIsNotAWinTests(unittest.TestCase):
    """Leads 331 -> 91 took lead-to-appointment from 4% to 9% while appointments fell.
    Reported bare, "conversion rose" told the dealer the opposite of what happened."""

    def copy(self, leads=("331", "91"), rate=("4.0%", "9.0%")):
        metrics: dict = {}
        core.add_metric(metrics, "sales_total_deals", "176", "114")
        core.add_metric(metrics, "sales_total_leads", *leads)
        core.add_metric(metrics, "sales_new_leads_to_appointments", *rate)
        core.add_metric(metrics, "sales_email_capture_rate", "75.0%", "69.0%")
        model = mm.build_month_model(dealership=DEALERSHIP, month_label=MONTH,
                                     metrics=metrics, breakdowns={})
        return fi.build_client_copy(model, "July", "June")

    def test_the_insight_says_the_base_shrank(self) -> None:
        bullets, _, _ = self.copy()
        line = next(b for b in bullets if b.startswith("Lead-to-appointment"))
        self.assertIn("on 72.5% fewer leads", line)

    def test_the_sms_does_not_lead_with_it(self) -> None:
        _, _, sms = self.copy()
        self.assertNotIn("appointment rate up", sms)

    def test_a_genuine_rise_is_left_alone(self) -> None:
        bullets, _, _ = self.copy(leads=("331", "240"))
        line = next(b for b in bullets if b.startswith("Lead-to-appointment"))
        self.assertNotIn("fewer leads", line)


class SmsReadsAsASentenceTests(unittest.TestCase):
    def test_the_body_starts_with_a_capital(self) -> None:
        _, _, sms = fi.build_client_copy(model_for(), "July", "June")
        self.assertTrue(sms.split("\n", 1)[1][:1].isupper(), sms)

    def test_the_body_starts_on_the_line_after_the_header(self) -> None:
        _, _, sms = fi.build_client_copy(model_for(), "July", "June")
        self.assertTrue(sms.startswith("DealerSocket July Insights:\n"), sms)
        self.assertNotIn("\n\n", sms)

    def test_a_count_shows_both_figures_and_a_rate_where_it_landed(self) -> None:
        count = mm.Movement(
            metric="sales_total_deals", label="Deal volume", previous=33.0, current=40.0,
            percent_change=21.2, improving=True, usable=True, band="strong",
            unit="count", metric_type="volume", point_change=None,
        )
        rate = mm.Movement(
            metric="sales_email_capture_rate", label="Email capture rate", previous=41.0, current=47.0,
            percent_change=14.6, improving=True, usable=True, band="strong",
            unit="percentage", metric_type="rate", point_change=6.0,
        )
        self.assertEqual(fi._compact_clause("units", count), "deals up (33→40)")
        self.assertEqual(fi._compact_clause("sales_email_capture_rate", rate), "email capture up to 47%")

    def test_sales_and_service_are_separate_sentences(self) -> None:
        def finding(key: str, compact: str, order: int) -> fi.Finding:
            return fi.Finding(key=key, text=compact, improving=True, order=order,
                              magnitude=1, short=compact, compact=compact)
        sms = fi.render_sms([
            finding("reminders", "service reminders up (178→195)", fi.ORDER_SERVICE_FUNNEL),
            finding("leads", "sales leads up (50→68)", fi.ORDER_SALES_FUNNEL),
            finding("units", "deals up (33→40)", fi.ORDER_UNITS),
        ], "August")
        self.assertEqual(sms, "DealerSocket August Insights:\n"
                              "Deals up (33→40), sales leads up (50→68). Service reminders up (178→195).")

    def test_units_are_called_deal_volume_everywhere(self) -> None:
        bullets, summary, sms = fi.build_client_copy(model_for(), "July", "June")
        self.assertNotIn("units", " ".join([*bullets, summary, sms]).lower())


class SmsUsesTheRoomItHasTests(unittest.TestCase):
    """Messages were stopping at two facts with 50+ characters to spare."""

    @staticmethod
    def finding(key: str, short: str, improving: bool, magnitude: float, hollow: bool = False) -> fi.Finding:
        return fi.Finding(key=key, text=short + ".", improving=improving, order=0,
                          magnitude=magnitude, short=short, hollow=hollow)

    def test_a_second_positive_is_added_when_it_fits(self) -> None:
        sms = fi.render_sms([
            self.finding("a", "sales campaign return up 121% to $75,334", True, 121),
            self.finding("b", "service campaign return up 83.0% to $35,120", True, 83),
            self.finding("c", "uncontacted leads up 140%", False, 140),
        ], "September")
        self.assertIn("service campaign return up 83.0%", sms)
        self.assertIn("uncontacted leads up 140%", sms)
        self.assertLessEqual(len(sms), fi.SMS_LIMIT)

    def test_a_shorter_fact_is_tried_when_the_next_one_does_not_fit(self) -> None:
        sms = fi.render_sms([
            self.finding("a", "sales campaign return up 56.8% to $20,559", True, 56.8),
            self.finding("b", "lead-to-test-drive conversion down 6 pts", False, 20),
            self.finding("c", "appointment-to-sale conversion down 3 pts across the month", False, 5),
            self.finding("d", "outstanding service tasks up 133%", False, 133),
        ], "September")
        self.assertIn("outstanding service tasks up 133%", sms)
        self.assertNotIn("appointment-to-sale", sms)
        self.assertLessEqual(len(sms), fi.SMS_LIMIT)

    def test_an_early_fact_does_not_crowd_out_a_fuller_message(self) -> None:
        sms = fi.render_sms([
            self.finding("a", "deal volume up 118% to 48", True, 118),
            self.finding("b", "lead volume up 38.3%", True, 38.3),
            self.finding("c", "uncontacted leads up 112%", False, 112),
            self.finding("d", "lead-to-appointment conversion down 2 pts", False, 5),
        ], "September")
        self.assertIn("lead-to-appointment conversion down 2 pts", sms)
        self.assertGreater(len(sms), 130)
        self.assertLessEqual(len(sms), fi.SMS_LIMIT)

    def test_each_side_leads_with_its_largest_movement(self) -> None:
        sms = fi.render_sms([
            self.finding("a", "lead-to-appointment conversion up 2 pts over the month", True, 5),
            self.finding("b", "sales campaign return down 36.5% to $25,468", False, 36.5),
            self.finding("c", "appointment-to-sale conversion up 5 pts", True, 12),
        ], "September")
        self.assertIn("ppointment-to-sale conversion up 5 pts", sms)
        self.assertNotIn("lead-to-appointment", sms.lower())

    def test_a_one_sided_month_is_not_capped_at_two_facts(self) -> None:
        sms = fi.render_sms([
            self.finding("a", "deal volume down 21.8% to 61", False, 21.8),
            self.finding("b", "sales campaign return down 96.6% to $621", False, 96.6),
            self.finding("c", "lead volume down 21.4%", False, 21.4),
            self.finding("d", "lead-to-appointment conversion up 2 pts on fewer leads", True, 50, hollow=True),
        ], "September")
        self.assertIn("lead volume down 21.4%", sms)
        self.assertNotIn("fewer leads", sms)
        self.assertLessEqual(len(sms), fi.SMS_LIMIT)


class AFigureThatCrossedZeroHasNoPercentageTests(unittest.TestCase):
    def test_a_return_that_went_negative_is_not_a_775_percent_fall(self) -> None:
        movement = mm.Movement(
            metric="sales_campaign_roi", label="Sales campaign return", previous=907.0,
            current=-6104.0, percent_change=-773.0, improving=False, usable=True,
            band="strong", unit="currency", metric_type="currency", point_change=None,
        )
        self.assertEqual(fi._movement_clause("Sales campaign return", movement, with_level=True),
                         "sales campaign return down to -$6,104")


class CampaignsCarryTheirShareTests(unittest.TestCase):
    """Naming the campaign is half the sentence; its share of the month is the half
    that saves the reader working it out."""

    def block(self, roi: int, previous: int, leader: int):
        from reporting.report_breakdowns import Campaign, CampaignBlock
        return CampaignBlock(
            named=[Campaign(name="HY NEW Prospect", count=12, roi=leader)],
            stated_roi=roi, previous_roi=previous,
        )

    def copy_for(self, block):
        model = mm.build_month_model(dealership=DEALERSHIP, month_label=MONTH, metrics={},
                                     breakdowns={"sales_campaigns": block})
        return " ".join(f.text for f in fi.build_findings(model))

    def test_the_share_of_the_month_is_stated(self) -> None:
        text = self.copy_for(self.block(roi=13286, previous=10710, leader=6877))
        self.assertIn("HY NEW Prospect returning $6,877 of it (52%)", text)

    def test_a_campaign_over_net_return_is_not_described_as_a_share(self) -> None:
        text = self.copy_for(self.block(roi=2898, previous=1836, leader=3049))
        self.assertIn("HY NEW Prospect returned $3,049", text)
        self.assertNotIn("of it", text)
        self.assertNotIn("105%", text)

    def test_a_campaign_that_returned_nothing_is_not_named_as_the_driver(self) -> None:
        """It may still be the largest of the three the report names, and naming it
        reads as an accusation while telling the reader nothing to act on."""
        text = self.copy_for(self.block(roi=0, previous=13765, leader=0))
        self.assertNotIn("HY NEW Prospect", text)
        self.assertIn("fell to $0, down from $13,765", text)

    def test_a_fall_to_zero_is_not_also_called_a_100_percent_fall(self) -> None:
        model = mm.build_month_model(dealership=DEALERSHIP, month_label=MONTH, metrics={},
                                     breakdowns={"sales_campaigns": self.block(0, 13765, 0)})
        finding = next(f for f in fi.build_findings(model) if f.key == "sales_campaign_roi")
        self.assertEqual(finding.statement, "sales campaign return fell to $0")
        self.assertNotIn("100%", finding.short)


class FiguresCarryTheirUnitTests(unittest.TestCase):
    def movement(self, **overrides) -> mm.Movement:
        record = {
            "metric": "sales_new_leads_to_appointments", "label": "Lead-to-appointment conversion",
            "previous": 4.0, "current": 9.0, "percent_change": 125.0, "improving": True,
            "usable": True, "band": "strong", "unit": "percentage", "metric_type": "rate",
            "point_change": 5.0,
        }
        record.update(overrides)
        return mm.Movement(**record)

    def station(self, movement) -> mm.Station:
        return mm.Station(key="appointments", label=movement.label, owner=mm.SALESPERSON, movement=movement)

    def test_a_rate_is_written_as_a_rate(self) -> None:
        sentence = fi._station_sentence(self.station(self.movement()))
        self.assertIn("from 4% to 9%", sentence)

    def test_a_rate_moves_in_points_not_in_per_cent_of_itself(self) -> None:
        """4% to 9% is five points. Calling it a 125% rise is true of the ratio and
        useless to a dealer, and it is the number that loses the argument in the room."""
        sentence = fi._station_sentence(self.station(self.movement()))
        self.assertIn("(5 pts)", sentence)
        self.assertNotIn("125", sentence)

    def test_one_point_is_singular(self) -> None:
        movement = self.movement(previous=8.0, current=9.0, point_change=1.0, percent_change=12.5)
        self.assertIn("(1 pt)", fi._station_sentence(self.station(movement)))

    def test_a_count_keeps_its_relative_change(self) -> None:
        movement = self.movement(label="Lead volume", unit="count", metric_type="volume",
                                 previous=331.0, current=91.0, percent_change=-72.5,
                                 point_change=None, improving=False)
        sentence = fi._station_sentence(mm.Station(key="leads", label="Lead volume",
                                                   owner=mm.SALESPERSON, movement=movement))
        self.assertIn("from 331 to 91 (72.5%)", sentence)

    def test_direction_words_describe_movement_not_welcome(self) -> None:
        """Overdue tasks rising is bad news and still a rise. Polarity lives in
        `improving`; calling the rise a fall is simply wrong."""
        movement = self.movement(label="Overdue sales tasks", unit="count", metric_type="volume",
                                 previous=40.0, current=55.0, percent_change=37.5,
                                 point_change=None, improving=False)
        station = mm.Station(key="sales_overdue_tasks", label="Overdue sales tasks",
                             owner=mm.SALESPERSON, movement=movement)
        self.assertIn("rose from", fi._station_sentence(station))
        self.assertIn("up", fi._movement_clause(station.label, movement, with_level=False))


class OrderIsGlobalTests(unittest.TestCase):
    """Findings were ordered by their index inside their own funnel, so a service
    station at index 0 tied with lead volume at index 0 and won on raw percentage."""

    def test_a_service_station_never_outranks_the_sales_chain(self) -> None:
        metrics = {
            "sales_total_leads": {"usable_for_reporting": True, "previous_value": 331.0,
                                  "current_value": 91.0, "percent_change": -72.5,
                                  "unit": "count", "metric_type": "volume",
                                  "significance_band": "strong", "higher_is_better": True},
            "service_reminders_sent": {"usable_for_reporting": True, "previous_value": 70.0,
                                       "current_value": 300.0, "percent_change": 328.6,
                                       "unit": "count", "metric_type": "volume",
                                       "significance_band": "strong", "higher_is_better": True},
        }
        model = mm.build_month_model(dealership=DEALERSHIP, month_label=MONTH,
                                     metrics=metrics, breakdowns={})
        keys = [f.key for f in fi.build_findings(model)]
        self.assertLess(keys.index("leads"), keys.index("reminders"))

    def test_the_bottom_line_leads_the_sales_chain(self) -> None:
        self.assertLess(fi.ORDER_UNITS, fi.ORDER_SALES_FUNNEL)
        self.assertLess(fi.ORDER_CAMPAIGN_RETURN, fi.ORDER_SALES_FUNNEL)
        self.assertLess(fi.ORDER_SALES_FUNNEL + len(mm.SALES_STATIONS), fi.ORDER_SERVICE_FUNNEL)
        self.assertLess(fi.ORDER_SERVICE_FUNNEL + len(mm.SERVICE_STATIONS), fi.ORDER_LOSING_CAMPAIGN)
        self.assertLess(fi.ORDER_LOSING_CAMPAIGN, fi.ORDER_HYGIENE)


if __name__ == "__main__":
    unittest.main()
