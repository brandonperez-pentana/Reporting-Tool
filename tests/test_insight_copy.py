"""Behaviour pinned by the insight copy rework.

The engine's job is to state what happened, factually, and to say why it matters
using arithmetic on figures already in the report. Each test here corresponds to a
defect that shipped: a number the engine distrusted printed anyway, a duration a
reader could not interpret, a strength field that shrugged while the data held an
improvement, or an observation where a comparison was available.
"""

from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

import dealership_insights_local as core
from reporting import metric_history


def metric(**overrides):
    record = {
        "previous": 100.0,
        "current": 90.0,
        "previous_value": 100.0,
        "current_value": 90.0,
        "delta": -10.0,
        "absolute_change": -10.0,
        "percent_change": -10.0,
        "percentage_point_change": None,
        "direction": "down",
        "unit": "count",
        "metric_type": "volume",
        "metric_label": "deal volume",
        "higher_is_better": True,
        "significance_band": "meaningful",
        "usable_for_reporting": True,
        "is_reliable": True,
        "is_suspicious": False,
    }
    record.update(overrides)
    return record


class DurationsCarryUnitsTests(unittest.TestCase):
    """"13:08" and "45:00" rendered identically and meant hours and minutes."""

    def test_hours_are_distinguishable_from_minutes(self) -> None:
        self.assertEqual(core.format_duration_value(788), "13h 08m")
        self.assertEqual(core.format_duration_value(45), "45 min")
        self.assertEqual(core.format_duration_value(0.5), "30 sec")

    def test_no_bare_colon_form_remains(self) -> None:
        for minutes in (0.5, 12.5, 45, 788, 1440):
            self.assertNotRegex(core.format_duration_value(minutes), r"^\d+:\d{2}$")


class SuspiciousMetricsAreNotPublishedTests(unittest.TestCase):
    """The engine flagged these already; only the review exports should see them."""

    def test_suspicious_metric_is_not_reportable(self) -> None:
        self.assertFalse(core.metric_usable_for_reporting(metric(is_suspicious=True)))

    def test_reliable_metric_is_reportable(self) -> None:
        self.assertTrue(core.metric_usable_for_reporting(metric()))

    def test_relationship_engine_applies_the_same_gate(self) -> None:
        from reporting.relationship_engine import _metric_usable

        self.assertFalse(_metric_usable(metric(is_suspicious=True)))
        self.assertTrue(_metric_usable(metric()))

    def test_a_zero_response_time_is_missing_data_not_an_instant_reply(self) -> None:
        """A Primary Strength read "Response times down from 7h 30m to 0 sec"."""
        record = core.build_metric_record("sales_internet_response_time", "7:30 hrs", "0:00 hrs")
        self.assertIsNotNone(record)
        self.assertFalse(core.metric_usable_for_reporting(record))
        self.assertIn("zero_response_time", record["reliability_reason"])


class SmallCountsAreNotMovementsTests(unittest.TestCase):
    """"Uncontacted leads rose from 2 to 3 (+50%)" once led a report."""

    def test_tiny_absolute_change_on_a_tiny_base_is_not_reportable(self) -> None:
        tiny = metric(
            previous=2.0, current=3.0, previous_value=2.0, current_value=3.0,
            delta=1.0, absolute_change=1.0, percent_change=50.0,
            direction="up", metric_type="count", significance_band="strong",
        )
        self.assertFalse(core.metric_usable_for_reporting(tiny))

    def test_same_percentage_on_a_real_base_still_reports(self) -> None:
        real = metric(
            previous=200.0, current=300.0, previous_value=200.0, current_value=300.0,
            delta=100.0, absolute_change=100.0, percent_change=50.0,
            direction="up", metric_type="count", significance_band="strong",
        )
        self.assertTrue(core.metric_usable_for_reporting(real))

    def test_rates_are_unaffected(self) -> None:
        rate = metric(
            previous=20.0, current=16.0, previous_value=20.0, current_value=16.0,
            delta=-4.0, absolute_change=-4.0, percent_change=-20.0,
            percentage_point_change=-4.0, metric_type="rate", significance_band="strong",
        )
        self.assertTrue(core.metric_usable_for_reporting(rate))


class ConsequenceArithmeticTests(unittest.TestCase):
    """A rate on its own asks the reader to do the multiplication."""

    def build(self):
        return {
            "sales_total_leads": metric(
                previous=1813.0, current=1643.0, previous_value=1813.0, current_value=1643.0,
                delta=-170.0, absolute_change=-170.0, percent_change=-9.38,
                metric_label="lead volume", significance_band="slight",
            ),
            "sales_total_deals": metric(
                previous=452.0, current=357.0, previous_value=452.0, current_value=357.0,
                delta=-95.0, absolute_change=-95.0, percent_change=-21.02,
            ),
            "sales_overall_conversion_rate": metric(
                previous=25.0, current=22.0, previous_value=25.0, current_value=22.0,
                delta=-3.0, absolute_change=-3.0, percent_change=-12.0,
                percentage_point_change=-3.0, metric_type="rate",
                metric_label="overall conversion", significance_band="strong",
            ),
            "sales_new_leads_to_test_drives": metric(
                previous=20.0, current=16.0, previous_value=20.0, current_value=16.0,
                delta=-4.0, absolute_change=-4.0, percent_change=-20.0,
                percentage_point_change=-4.0, metric_type="rate",
                metric_label="lead-to-test-drive conversion", significance_band="strong",
            ),
        }

    def test_rate_move_is_restated_in_units(self) -> None:
        clause = core.rate_consequence_clause(
            self.build(), "sales_new_leads_to_test_drives", "sales_total_leads", "test drives"
        )
        self.assertEqual(clause, "about 66 fewer test drives on this month's lead count")

    def test_conversion_counterfactual_names_the_deal_gap(self) -> None:
        clause = core.conversion_counterfactual_clause(self.build())
        self.assertIn("would have returned about 411 deals, not 357", clause)

    def test_counterfactual_is_withheld_when_conversion_is_not_deals_over_leads(self) -> None:
        metrics = self.build()
        metrics["sales_total_deals"]["current"] = 40.0
        metrics["sales_total_deals"]["current_value"] = 40.0
        self.assertIsNone(core.conversion_counterfactual_clause(metrics))

    def test_campaign_return_separates_volume_from_efficiency(self) -> None:
        metrics = {
            "sales_campaign_roi": metric(
                previous=46824.0, current=18886.0, previous_value=46824.0, current_value=18886.0,
                delta=-27938.0, absolute_change=-27938.0, percent_change=-59.67,
                metric_type="currency", metric_label="sales campaign ROI", significance_band="strong",
            ),
            "sales_campaign_deals": metric(
                previous=122.0, current=85.0, previous_value=122.0, current_value=85.0,
                delta=-37.0, absolute_change=-37.0, percent_change=-30.33,
                metric_label="sales campaign deals", significance_band="strong",
            ),
        }
        clause = core.campaign_return_per_deal_clause(metrics, "sales_campaign_roi", "sales_campaign_deals")
        self.assertEqual(clause, "return per attributed deal fell from $384 to $222")

    def test_clauses_avoid_phrases_the_sms_layer_keys_on(self) -> None:
        """The SMS composer reads facts back out of the prose; a clause must not
        register as a second, separate movement in the metric it is measured against."""
        clause = core.rate_consequence_clause(
            self.build(), "sales_new_leads_to_test_drives", "sales_total_leads", "test drives"
        )
        for needle in core.FAMILY_INSIGHT_NEEDLES["lead_volume"]:
            self.assertNotIn(needle, clause.lower())


class BaselineComparisonTests(unittest.TestCase):
    """An observation says what happened; a comparison says whether to care."""

    def roi_metric(self, series, current=7960.0, direction="down"):
        record = metric(
            previous=14243.0, current=current, previous_value=14243.0, current_value=current,
            delta=current - 14243.0, absolute_change=current - 14243.0, percent_change=-44.1,
            direction=direction, metric_type="currency",
            metric_label="service campaign ROI", significance_band="strong",
        )
        record["history_series"] = series
        record["history_same_month_last_year"] = None
        return {"service_campaign_roi": record}

    def test_lowest_in_the_record_is_named(self) -> None:
        series = [("Mar 2026", 12783.0), ("Apr 2026", 11879.0), ("May 2026", 9996.0), ("Jun 2026", 14243.0)]
        clause = core.baseline_clause("service_campaign_roi", self.roi_metric(series))
        self.assertEqual(clause, ("the lowest in 5 months", "record"))

    def test_no_claim_is_made_on_a_short_record(self) -> None:
        series = [("May 2026", 9996.0), ("Jun 2026", 14243.0)]
        clause = core.baseline_clause("service_campaign_roi", self.roi_metric(series))
        self.assertNotEqual((clause or ("", ""))[1], "record")

    def test_same_month_last_year_is_used_when_the_record_is_short(self) -> None:
        metrics = self.roi_metric([("Jun 2026", 14243.0)])
        metrics["service_campaign_roi"]["history_same_month_last_year"] = 20000.0
        clause = core.baseline_clause("service_campaign_roi", metrics)
        self.assertEqual(clause, ("against $20,000 the same month last year", "reference"))

    def test_a_comparison_is_dropped_where_it_could_bind_to_the_wrong_metric(self) -> None:
        mixed = "Response times slowed from 1h 40m to 2h 43m (+63%), while uncontacted leads fell from 27 to 13."
        single = "Service campaign ROI fell from $14,243 to $7,960 (-44.1%)."
        self.assertTrue(core.sentence_describes_opposite_moves(mixed))
        self.assertFalse(core.sentence_describes_opposite_moves(single))


class PrimaryStrengthTests(unittest.TestCase):
    """The field shrugged in 12 of 27 reports while the data held an improvement."""

    def test_an_improving_metric_is_found_outside_the_selected_rules(self) -> None:
        metrics = {
            "service_customers_enrolled": metric(
                previous=978.0, current=1030.0, previous_value=978.0, current_value=1030.0,
                delta=52.0, absolute_change=52.0, percent_change=5.32, direction="up",
                metric_type="count", metric_label="customers enrolled", significance_band="slight",
            )
        }
        self.assertEqual(core.best_improving_metric_phrase(metrics), "Customers enrolled up 5.3%")

    def test_a_worsening_metric_is_never_offered_as_a_strength(self) -> None:
        self.assertIsNone(core.best_improving_metric_phrase({"sales_total_deals": metric()}))

    def test_a_strong_month_still_names_its_real_decline_as_the_risk(self) -> None:
        """A strong month read "No major risk signal identified" with leads down 25.6%."""
        profile = core.InternalScoringProfile(
            sales_score=0, service_score=0, risk_score=0, overall_score=0,
            rule_based_rating="Strong", derived_month_rating="Strong", final_rating_used="Strong",
            anomaly_count=0, anomaly_types=[], unreliable_metric_count=0,
            suspicious_metric_count=0, scoring_conflict=False)
        metrics = {"sales_total_deals": metric()}
        _, _, risk, _ = core.build_executive_summary_fields(
            [], metrics, "strongly_improved", "rule-based", "success", None, profile)
        self.assertEqual(risk, core.worst_declining_metric_phrase(metrics))
        self.assertNotEqual(risk, "No major risk signal identified")

    def test_a_tiny_base_is_stated_as_figures_not_a_percentage(self) -> None:
        """Campaign deals 2 -> 10 read as "up 400%", which oversells eight deals."""
        metrics = {
            "sales_campaign_deals": metric(
                previous=2.0, current=10.0, previous_value=2.0, current_value=10.0,
                delta=8.0, absolute_change=8.0, percent_change=400.0, direction="up",
                metric_type="count", metric_label="sales campaign deals", significance_band="strong",
            )
        }
        self.assertEqual(core.best_improving_metric_phrase(metrics), "Sales campaign deals up from 2 to 10")

    def test_a_fall_to_zero_is_stated_as_figures_not_100_percent(self) -> None:
        """A strength read "uncontacted leads down 100%"."""
        metrics = {
            "sales_uncontacted_leads": metric(
                previous=51.0, current=0.0, previous_value=51.0, current_value=0.0,
                delta=-51.0, absolute_change=-51.0, percent_change=-100.0, direction="down",
                metric_type="count", metric_label="uncontacted leads", higher_is_better=False,
                significance_band="strong",
            )
        }
        self.assertEqual(core.best_improving_metric_phrase(metrics), "Uncontacted leads down from 51 to 0")

    def test_a_return_that_turned_into_a_loss_is_stated_as_figures(self) -> None:
        """"Sales campaign ROI down 109.1%" and "$-10,637" both shipped."""
        metrics = {
            "sales_campaign_roi": metric(
                previous=10000.0, current=-900.0, previous_value=10000.0, current_value=-900.0,
                delta=-10900.0, absolute_change=-10900.0, percent_change=-109.0,
                unit="currency", metric_type="currency", metric_label="sales campaign ROI",
                significance_band="strong",
            )
        }
        self.assertEqual(core.worst_declining_metric_phrase(metrics),
                         "Sales campaign ROI down from $10,000 to -$900")
        self.assertEqual(core.format_currency_value(-10637), "-$10,637")

    def test_half_a_contrast_is_not_a_headline(self) -> None:
        """A Primary Risk read "Lead volume up 45%" when conversion was not reportable."""
        from types import SimpleNamespace

        rule = SimpleNamespace(rule_id="leads_up_conversion_down", supporting_metrics=[])
        rising_leads = {"sales_total_leads": metric(
            previous=91.0, current=132.0, previous_value=91.0, current_value=132.0,
            delta=41.0, absolute_change=41.0, percent_change=45.0, direction="up",
            metric_label="lead volume", significance_band="strong")}
        self.assertIsNone(core.build_rule_headline(rule, rising_leads))

        rule = SimpleNamespace(rule_id="service_bookings_roi_mismatch", supporting_metrics=[])
        rising_bookings = {"service_bookings_from_campaigns": metric(
            previous=50.0, current=70.0, previous_value=50.0, current_value=70.0,
            delta=20.0, absolute_change=20.0, percent_change=40.0, direction="up",
            metric_label="service campaign bookings", significance_band="strong")}
        self.assertIsNone(core.build_rule_headline(rule, rising_bookings))

    def test_direction_is_read_against_what_is_good_for_the_dealership(self) -> None:
        fewer_uncontacted = metric(
            metric_type="count", metric_label="uncontacted leads",
            higher_is_better=False, direction="down",
        )
        self.assertTrue(core.metric_is_improvement(fewer_uncontacted))


class PlaceholderMonthsTests(unittest.TestCase):
    """A month whose PDFs could not be read must not assert performance."""

    def test_placeholder_copy_states_the_absence(self) -> None:
        bullets, summary, sms = core.generate_placeholder_insights()
        joined = " ".join([*bullets, summary, sms]).lower()
        for claim in ("increased", "improved", "steady overall", "performance appears steady"):
            self.assertNotIn(claim, joined)
        self.assertIn("no sales or service figures", joined)

    def test_fallback_generation_keeps_sms_explicitly_non_sendable(self) -> None:
        result = core.generate_insights_from_metrics(
            {}, "blocked", "Demo Site 1", "Jul 2026"
        )
        self.assertIn("no figures were available", result[1].lower())
        self.assertTrue(result[2].startswith("Not for sending:"), result[2])

    def test_a_placeholder_month_is_not_padded_to_four(self) -> None:
        bullets, _, _ = core.generate_placeholder_insights()
        self.assertEqual(len(core.ensure_four_priority_insights(bullets, metrics=None, pad=False)), len(bullets))

    def test_a_placeholder_does_not_publish_partial_metrics_as_findings(self) -> None:
        partial_metrics = {"sales_total_deals": metric()}
        self.assertEqual(
            core.ensure_four_priority_insights([], metrics=partial_metrics, pad=False),
            [],
        )

    def test_a_real_month_still_pads_to_four(self) -> None:
        self.assertEqual(len(core.ensure_four_priority_insights(["One fact."], metrics=None)), 4)


class RecommendedFocusIsRetiredTests(unittest.TestCase):
    def test_the_field_is_absent_from_the_report(self) -> None:
        result = core.DealershipResult(
            dealership_name="Example Motors",
            month_label="Jul 2026",
            bullet_insights=["One fact."],
            summary="A summary.",
            sms_version="An SMS.",
            report_status="success",
            parsing_status="parsed",
            confidence_score=1.0,
            insights_source="rule-based",
            top_priority_level="High",
            review_required="no",
            sentiment="mixed",
            insight_count_after_deduplication=1,
        )
        lines = "\n".join(core.build_executive_summary_lines(result))
        self.assertNotIn("Recommended Focus", lines)
        self.assertIn("Primary Risk", lines)


class MetricHistoryStoreTests(unittest.TestCase):
    """Trends used to need last month's PDFs to still be in the folder."""

    def test_a_month_round_trips(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "metric_history.json"
            store = metric_history.load_store(path)
            metric_history.record_month(store, "Example Motors", "Jun 2026", {"sales_total_deals": metric()})
            metric_history.record_month(
                store, "Example Motors", "Jul 2026",
                {"sales_total_deals": metric(current=80.0, current_value=80.0)},
            )
            self.assertTrue(metric_history.save_store(store, path))

            reloaded = metric_history.load_store(path)
            self.assertEqual(metric_history.recorded_months(reloaded, "Example Motors"), ["Jun 2026", "Jul 2026"])
            series = metric_history.metric_series(reloaded, "Example Motors", "sales_total_deals", before_month="Jul 2026")
            self.assertEqual(series, [("Jun 2026", 90.0)])

    def test_rerunning_a_month_overwrites_rather_than_duplicates(self) -> None:
        store = {"version": 1, "dealerships": {}}
        metric_history.record_month(store, "Example Motors", "Jul 2026", {"sales_total_deals": metric()})
        metric_history.record_month(
            store, "Example Motors", "Jul 2026",
            {"sales_total_deals": metric(current=80.0, current_value=80.0)},
        )
        self.assertEqual(metric_history.recorded_months(store, "Example Motors"), ["Jul 2026"])
        self.assertEqual(
            metric_history.value_at(store, "Example Motors", "sales_total_deals", "Jul 2026"), 80.0
        )

    def test_unreliable_figures_are_kept_but_excluded_from_baselines(self) -> None:
        store = {"version": 1, "dealerships": {}}
        metric_history.record_month(
            store, "Example Motors", "Jun 2026", {"sales_total_deals": metric(is_suspicious=True)}
        )
        self.assertEqual(
            metric_history.metric_series(store, "Example Motors", "sales_total_deals", before_month="Jul 2026"), []
        )

    def test_month_labels_parse(self) -> None:
        self.assertEqual(metric_history.month_key("Jul 2026"), (2026, 7))
        self.assertEqual(metric_history.same_month_last_year_label("Jul 2026"), "Jul 2025")
        self.assertIsNone(metric_history.month_key("not a month"))

    def test_an_unwritable_store_does_not_raise(self) -> None:
        """History is a nicety; failing to write it must never break the month.

        The unwritable path is a file standing where a parent folder should be, which
        fails for every account. An absent directory does not: an operator running as
        administrator can simply create it, and the test would pass by writing the file
        it was meant to prove could not be written.
        """
        with tempfile.TemporaryDirectory() as tmp:
            blocker = Path(tmp) / "not-a-folder"
            blocker.write_text("", encoding="utf-8")
            target = blocker / "x" / "y.json"
            self.assertFalse(metric_history.save_store({"version": 1, "dealerships": {}}, target))
            self.assertFalse(target.exists())


class LeadWithOneTests(unittest.TestCase):
    """An owner reads the first line."""

    def rule(self, rule_id, metric_names):
        return core.RuleInsight(
            rule_id=rule_id, message="", category="sales_overview", severity="medium",
            priority_weight=50, polarity="mixed", tone_band="slightly_softer",
            merge_key=rule_id, supporting_metrics=list(metric_names), metric_focus=list(metric_names),
        )

    def test_a_decisively_bigger_movement_is_promoted(self) -> None:
        metrics = {
            "service_outstanding_tasks": metric(
                previous=36.0, current=33.0, previous_value=36.0, current_value=33.0,
                delta=-3.0, absolute_change=-3.0, percent_change=-8.3,
                metric_type="count", metric_label="outstanding service tasks",
                higher_is_better=False, significance_band="slight",
            ),
            "sales_overall_conversion_rate": metric(
                previous=44.0, current=37.0, previous_value=44.0, current_value=37.0,
                delta=-7.0, absolute_change=-7.0, percent_change=-15.9,
                percentage_point_change=-7.0, metric_type="rate",
                metric_label="overall conversion", significance_band="strong",
            ),
        }
        selected = [
            self.rule("service_task_snapshot", ["service_outstanding_tasks"]),
            self.rule("sales_conversion_snapshot", ["sales_overall_conversion_rate"]),
        ]
        ordered = core.lead_with_most_material_insight(selected, metrics)
        self.assertEqual(ordered[0].rule_id, "sales_conversion_snapshot")

    def test_editorial_order_is_left_alone_when_the_margin_is_narrow(self) -> None:
        metrics = {
            "sales_total_deals": metric(percent_change=-20.0, significance_band="meaningful"),
            "sales_campaign_roi": metric(
                percent_change=-21.0, metric_type="currency",
                metric_label="sales campaign ROI", significance_band="meaningful",
            ),
        }
        selected = [
            self.rule("sales_volume_snapshot", ["sales_total_deals"]),
            self.rule("sales_campaign_roi_decline", ["sales_campaign_roi"]),
        ]
        self.assertEqual(core.lead_with_most_material_insight(selected, metrics)[0].rule_id, "sales_volume_snapshot")


class SmsComposedFromMetricsTests(unittest.TestCase):
    """The SMS used to derive its facts by pattern-matching the finished insight prose.

    A direction word taken from one clause and a number from another is how
    "uncontacted leads up -21.8%", "service tasks down -8.3%" and "email capture down
    from 41% to 43%" reached clients. Every case below is one of those, rebuilt from the
    metric record it should always have come from.
    """

    def rule(self, rule_id, metric_names):
        return core.RuleInsight(
            rule_id=rule_id, message="", category="sales_overview", severity="medium",
            priority_weight=50, polarity="mixed", tone_band="slightly_softer",
            merge_key=rule_id, supporting_metrics=list(metric_names), metric_focus=list(metric_names),
        )

    def test_a_rise_is_never_written_as_a_negative_percentage(self) -> None:
        """"uncontacted leads up -21.8%" — the sign came from the prose, the word did not."""
        metrics = {
            "sales_uncontacted_leads": metric(
                previous=55.0, current=43.0, previous_value=55.0, current_value=43.0,
                delta=-12.0, absolute_change=-12.0, percent_change=-21.8, direction="down",
                metric_type="count", metric_label="uncontacted leads", higher_is_better=False,
            )
        }
        phrase = core.sms_phrase_candidates("sales_uncontacted_leads", metrics)[0]
        self.assertEqual(phrase, "uncontacted leads down 21.8% to 43")
        self.assertNotIn("-", phrase)

    def test_the_direction_word_follows_the_metric(self) -> None:
        """"email capture down from 41% to 43%" — the numbers went up."""
        metrics = {
            "sales_email_capture_rate": metric(
                previous=41.0, current=43.0, previous_value=41.0, current_value=43.0,
                delta=2.0, absolute_change=2.0, percent_change=4.9,
                percentage_point_change=2.0, direction="up", metric_type="rate",
                metric_label="email capture rate", significance_band="meaningful",
            )
        }
        self.assertEqual(
            core.sms_phrase_candidates("sales_email_capture_rate", metrics)[0],
            "email capture up from 41% to 43%",
        )

    def test_a_rate_never_carries_a_relative_percentage(self) -> None:
        """"response times improved 163.4%" — a rate's relative change reads as the rate."""
        metrics = {
            "sales_overall_conversion_rate": metric(
                previous=44.0, current=37.0, previous_value=44.0, current_value=37.0,
                delta=-7.0, absolute_change=-7.0, percent_change=-15.9,
                percentage_point_change=-7.0, metric_type="rate",
                metric_label="overall conversion", significance_band="strong",
            )
        }
        for phrase in core.sms_phrase_candidates("sales_overall_conversion_rate", metrics):
            self.assertNotIn("15.9", phrase)

    def test_durations_read_as_pace_and_keep_their_units(self) -> None:
        metrics = {
            "sales_internet_response_time": metric(
                previous=788.0, current=931.0, previous_value=788.0, current_value=931.0,
                delta=143.0, absolute_change=143.0, percent_change=18.1, direction="up",
                metric_type="time", metric_label="response times",
            )
        }
        self.assertEqual(
            core.sms_phrase_candidates("sales_internet_response_time", metrics)[0],
            "response time slower at 15h 31m",
        )

    def test_a_flat_metric_produces_nothing(self) -> None:
        self.assertEqual(core.sms_phrase_candidates("sales_total_deals", {"sales_total_deals": metric(direction="flat")}), [])


class SmsTraceabilityTests(unittest.TestCase):
    """Every number in the SMS has to be one the report already showed."""

    def metrics(self):
        return {
            "sales_campaign_roi": metric(
                previous=46824.0, current=18886.0, previous_value=46824.0, current_value=18886.0,
                delta=-27938.0, absolute_change=-27938.0, percent_change=-59.67,
                metric_type="currency", metric_label="sales campaign ROI", significance_band="strong",
            )
        }

    def test_a_phrase_whose_numbers_are_in_the_report_is_used(self) -> None:
        insights = ["Sales campaign ROI fell from $46,824 to $18,886 (-59.7%)."]
        signals = core.sms_signals_from_metrics(None, self.metrics(), insights)
        self.assertEqual([s.phrase for s in signals], ["sales ROI down 59.7% to $18,886"])

    def test_a_metric_the_report_never_mentioned_is_left_out(self) -> None:
        signals = core.sms_signals_from_metrics(None, self.metrics(), ["Deal volume fell from 452 to 357 (-21%)."])
        self.assertEqual(signals, [])

    def test_composition_falls_back_to_simpler_forms_to_stay_traceable(self) -> None:
        """The report showed only the percentage, so the SMS may only use the percentage."""
        insights = ["Sales campaign ROI down 59.7%."]
        signals = core.sms_signals_from_metrics(None, self.metrics(), insights)
        self.assertEqual([s.phrase for s in signals], ["sales ROI down 59.7%"])


class SmsSelectionTests(unittest.TestCase):
    def build(self):
        return {
            "sales_campaign_roi": metric(
                previous=46824.0, current=18886.0, previous_value=46824.0, current_value=18886.0,
                delta=-27938.0, absolute_change=-27938.0, percent_change=-59.67,
                metric_type="currency", metric_label="sales campaign ROI", significance_band="strong",
            ),
            "service_outstanding_tasks": metric(
                previous=36.0, current=33.0, previous_value=36.0, current_value=33.0,
                delta=-3.0, absolute_change=-3.0, percent_change=-8.3, metric_type="count",
                metric_label="outstanding service tasks", higher_is_better=False, significance_band="slight",
            ),
            "sales_total_deals": metric(
                previous=452.0, current=357.0, previous_value=452.0, current_value=357.0,
                delta=-95.0, absolute_change=-95.0, percent_change=-21.02,
                metric_label="deal volume", significance_band="meaningful",
            ),
            "sales_total_leads": metric(
                previous=1813.0, current=1643.0, previous_value=1813.0, current_value=1643.0,
                delta=-170.0, absolute_change=-170.0, percent_change=-9.38,
                metric_label="lead volume", significance_band="slight",
            ),
        }

    def insights(self):
        return [
            "Sales campaign ROI fell from $46,824 to $18,886 (-59.7%).",
            "Outstanding service tasks fell from 36 to 33 (-8.3%).",
            "Deal volume fell from 452 to 357 (-21%).",
            "Lead volume eased from 1,813 to 1,643 (-9.4%).",
        ]

    def test_the_biggest_movement_leads(self) -> None:
        signals = core.sms_signals_from_metrics(None, self.build(), self.insights())
        self.assertTrue(signals[0].phrase.startswith("sales ROI"))

    def test_no_more_than_three_facts(self) -> None:
        self.assertLessEqual(len(core.sms_signals_from_metrics(None, self.build(), self.insights())), core.MAX_SMS_SIGNALS)

    def test_improvement_is_read_against_what_is_good_for_the_dealership(self) -> None:
        signals = core.sms_signals_from_metrics(None, self.build(), self.insights())
        tasks = [s for s in signals if s.metric == "service_outstanding_tasks"]
        for signal in tasks:
            self.assertEqual(signal.polarity, "positive")

    def test_the_finished_message_is_within_the_limit_and_valid(self) -> None:
        metrics = self.build()
        sms, segments, meta = core.build_sms_from_priority_insights(
            self.insights(), "Jul 2026", metrics=metrics, selected_rules=None
        )
        self.assertEqual(meta["sms_composition"], "metrics")
        self.assertTrue(sms.startswith("DealerSocket July Insights:"))
        self.assertLessEqual(len(sms), core.SMS_CHAR_LIMIT)
        self.assertEqual(core.validate_sms_no_invented_numbers(sms, self.insights()), [])
        self.assertTrue(meta["sms_valid"])

    def test_callers_holding_only_text_still_get_an_sms(self) -> None:
        _, _, meta = core.build_sms_from_priority_insights(self.insights(), "Jul 2026")
        self.assertEqual(meta["sms_composition"], "text")
        self.assertTrue(meta["sms_valid"])


if __name__ == "__main__":
    unittest.main()
def _usable(previous: float, current: float, band: str = "strong") -> dict:
    return {
        "usable_for_reporting": True,
        "previous_value": previous,
        "current_value": current,
        "percent_change": round((current - previous) / previous * 100, 2) if previous else 0.0,
        "direction": "up" if current > previous else "down",
        "unit": "count",
        "metric_type": "volume",
        "significance_band": band,
    }


def _unusable() -> dict:
    """A metric the report parsed but cannot use: flat at zero, both months."""
    return {
        "usable_for_reporting": False,
        "previous_value": 0.0,
        "current_value": 0.0,
        "percent_change": 0.0,
        "direction": "flat",
        "unit": "count",
        "metric_type": "volume",
        "significance_band": "none",
    }


class EmptyFamilyCoverageTests(unittest.TestCase):
    """One half of the business reporting nothing is a state the engine must recognise."""

    def test_coverage_counts_what_each_half_contributed(self) -> None:
        metrics = {
            "sales_total_leads": _usable(331, 91),
            "sales_total_deals": _usable(176, 114),
            "service_campaign_roi": _unusable(),
        }
        coverage = core.metric_family_coverage(metrics)
        self.assertEqual(coverage["sales"], {"parsed": 2, "usable": 2})
        self.assertEqual(coverage["service"], {"parsed": 1, "usable": 0})

    def test_a_half_with_nothing_usable_is_reported(self) -> None:
        metrics = {"sales_total_leads": _usable(331, 91), "service_campaign_roi": _unusable()}
        self.assertEqual(core.families_without_usable_evidence(metrics), ["service"])

    def test_both_halves_empty_is_not_an_imbalance(self) -> None:
        """That is a parsing failure, handled upstream as a placeholder report."""
        metrics = {"sales_total_leads": _unusable(), "service_campaign_roi": _unusable()}
        self.assertEqual(core.families_without_usable_evidence(metrics), [])

    def test_a_full_month_reports_no_empty_half(self) -> None:
        metrics = {"sales_total_leads": _usable(331, 91), "service_campaign_roi": _usable(10059, 44183)}
        self.assertEqual(core.families_without_usable_evidence(metrics), [])


class RebalanceWhenServiceIsEmptyTests(unittest.TestCase):
    """With no service figures every line comes from sales.

    In a soft month that reads as an unbroken run of bad news, while the same
    dealership with service data leads on a campaign ROI rise. The imbalance is in the
    ordering, not the facts — so a line already in the report moves up, and nothing is
    added, removed or reworded.
    """

    EMPTY_SERVICE = {"sales_total_leads": _usable(331, 91), "service_campaign_roi": _unusable()}
    FULL_MONTH = {"sales_total_leads": _usable(331, 91), "service_campaign_roi": _usable(10059, 44183)}

    NEGATIVE_FIRST = [
        "Lead volume eased from 331 to 91 (-72.5%).",
        "Deal volume fell from 176 to 114 (-35.2%).",
        "Lead-to-appointment conversion up from 4% to 9%.",
    ]

    def test_an_improvement_leads_when_service_reported_nothing(self) -> None:
        out = core.rebalance_when_a_family_reported_nothing(self.NEGATIVE_FIRST, self.EMPTY_SERVICE)
        self.assertEqual(out[0], "Lead-to-appointment conversion up from 4% to 9%.")

    def test_every_decline_is_still_there(self) -> None:
        out = core.rebalance_when_a_family_reported_nothing(self.NEGATIVE_FIRST, self.EMPTY_SERVICE)
        self.assertEqual(sorted(out), sorted(self.NEGATIVE_FIRST), "nothing may be dropped or reworded")

    def test_a_full_month_is_left_exactly_as_ranked(self) -> None:
        out = core.rebalance_when_a_family_reported_nothing(self.NEGATIVE_FIRST, self.FULL_MONTH)
        self.assertEqual(out, self.NEGATIVE_FIRST)

    def test_an_all_bad_month_stays_all_bad(self) -> None:
        """No positive in the report means no positive is invented to open it."""
        bullets = [
            "Lead volume eased from 331 to 91 (-72.5%).",
            "Deal volume fell from 176 to 114 (-35.2%).",
        ]
        self.assertEqual(
            core.rebalance_when_a_family_reported_nothing(bullets, self.EMPTY_SERVICE), bullets
        )

    def test_a_report_already_opening_well_is_untouched(self) -> None:
        bullets = [
            "Lead-to-appointment conversion up from 4% to 9%.",
            "Lead volume eased from 331 to 91 (-72.5%).",
        ]
        self.assertEqual(
            core.rebalance_when_a_family_reported_nothing(bullets, self.EMPTY_SERVICE), bullets
        )
def _moved(previous: float, current: float, band: str, higher_is_better: bool = True) -> dict:
    return {
        "usable_for_reporting": True,
        "previous_value": previous,
        "current_value": current,
        "percent_change": round((current - previous) / previous * 100, 2) if previous else 0.0,
        "direction": "up" if current > previous else "down",
        "unit": "count",
        "metric_type": "volume",
        "significance_band": band,
        "higher_is_better": higher_is_better,
    }


class MetricBalanceTests(unittest.TestCase):
    """Which way the month actually went, weighed across every usable metric."""

    def test_a_clearly_good_month_reads_positive(self) -> None:
        metrics = {
            "sales_total_deals": _moved(150, 190, "strong"),
            "sales_overall_conversion_rate": _moved(50, 62, "meaningful"),
            "sales_overdue_tasks": _moved(140, 138, "slight", higher_is_better=False),
        }
        self.assertEqual(core.metric_balance_direction(metrics), "positive")

    def test_a_clearly_bad_month_reads_negative(self) -> None:
        metrics = {
            "sales_total_deals": _moved(190, 120, "strong"),
            "sales_overall_conversion_rate": _moved(62, 44, "strong"),
            "sales_email_capture_rate": _moved(80, 78, "slight"),
        }
        self.assertEqual(core.metric_balance_direction(metrics), "negative")

    def test_a_genuinely_mixed_month_has_no_direction(self) -> None:
        metrics = {
            "sales_total_deals": _moved(150, 190, "strong"),
            "sales_overall_conversion_rate": _moved(62, 44, "strong"),
        }
        self.assertEqual(core.metric_balance_direction(metrics), "neutral")

    def test_many_trivial_moves_do_not_outvote_one_collapse(self) -> None:
        """Counting metrics rather than weighing them would call this a good month."""
        metrics = {
            "sales_total_deals": _moved(190, 90, "strong"),
            "sales_email_capture_rate": _moved(70, 72, "slight"),
            "sales_same_day_deal_pct": _moved(60, 62, "slight"),
        }
        self.assertNotEqual(core.metric_balance_direction(metrics), "positive")


class LaneReconciliationTests(unittest.TestCase):
    """The opening line must not claim the opposite of the month's own numbers.

    The lane is chosen from a sentiment score and from keywords in the headline
    clauses; neither consults the metrics. That let a month with ten improvements and
    five declines open "a clear step down across the main numbers" because two of the
    declines were large and service-side, and let a month with nine declines open
    "firmer service campaign returns" on the strength of one ROI rise.
    """

    GOOD_MONTH = {
        "sales_total_deals": _moved(150, 190, "strong"),
        "sales_overall_conversion_rate": _moved(50, 62, "meaningful"),
    }
    BAD_MONTH = {
        "sales_total_deals": _moved(190, 120, "strong"),
        "sales_overall_conversion_rate": _moved(62, 44, "strong"),
    }

    def test_a_soft_opener_is_dropped_when_the_month_improved(self) -> None:
        self.assertEqual(core.reconcile_lane_with_metrics("general_soft", self.GOOD_MONTH), "mixed")
        self.assertEqual(core.reconcile_lane_with_metrics("campaign_weak", self.GOOD_MONTH), "mixed")
        self.assertEqual(core.reconcile_lane_with_metrics("crm_load", self.GOOD_MONTH), "mixed")

    def test_a_strong_opener_is_dropped_when_the_month_fell(self) -> None:
        self.assertEqual(core.reconcile_lane_with_metrics("general_improved", self.BAD_MONTH), "mixed")
        self.assertEqual(core.reconcile_lane_with_metrics("service_gain", self.BAD_MONTH), "mixed")

    def test_it_never_flips_to_the_opposite_claim(self) -> None:
        """Dropping to mixed is the correction; asserting the reverse would invert the error."""
        self.assertNotEqual(core.reconcile_lane_with_metrics("general_soft", self.GOOD_MONTH), "general_improved")
        self.assertNotEqual(core.reconcile_lane_with_metrics("general_improved", self.BAD_MONTH), "general_soft")

    def test_an_opener_that_agrees_with_the_month_is_kept(self) -> None:
        self.assertEqual(core.reconcile_lane_with_metrics("general_soft", self.BAD_MONTH), "general_soft")
        self.assertEqual(core.reconcile_lane_with_metrics("general_improved", self.GOOD_MONTH), "general_improved")

    def test_a_month_with_no_clear_direction_keeps_its_lane(self) -> None:
        """Most months are genuinely mixed; the guard only fires on a real contradiction."""
        even = {
            "sales_total_deals": _moved(150, 190, "strong"),
            "sales_overall_conversion_rate": _moved(62, 44, "strong"),
        }
        self.assertEqual(core.reconcile_lane_with_metrics("general_soft", even), "general_soft")

    def test_neutral_lanes_are_left_alone(self) -> None:
        for lane in ("steady", "mixed"):
            self.assertEqual(core.reconcile_lane_with_metrics(lane, self.BAD_MONTH), lane)
            self.assertEqual(core.reconcile_lane_with_metrics(lane, self.GOOD_MONTH), lane)
class WholeMonthCoverageTests(unittest.TestCase):
    """A month that moved both ways must not be reported as if it moved one way.

    Ranking by materiality alone can fill every slot from one side. Leads down while
    campaign returns climb is an ordinary month, and an owner shown only half of it has
    been given a false picture of their own business.
    """

    MIXED = {
        "sales_total_leads": _moved(331, 91, "strong"),
        "sales_campaign_roi": _moved(10000, 44000, "strong"),
    }
    ONE_WAY = {
        "sales_total_leads": _moved(331, 91, "strong"),
        "sales_total_deals": _moved(176, 114, "strong"),
        "sales_email_capture_rate": _moved(75, 74, "slight"),
    }

    def test_a_month_that_moved_both_ways_is_recognised(self) -> None:
        self.assertTrue(core.month_moved_both_ways(self.MIXED))

    def test_a_trivial_move_does_not_make_a_month_mixed(self) -> None:
        """Otherwise a 1% drift would license a token positive in a collapse."""
        self.assertFalse(core.month_moved_both_ways(self.ONE_WAY))

    def _rule(self, rule_id: str, polarity: str) -> core.RuleInsight:
        return core.RuleInsight(
            rule_id=rule_id,
            message=f"{rule_id} message",
            category="sales",
            severity="medium",
            priority_weight=50,
            polarity=polarity,
            tone_band="neutral",
            merge_key=rule_id,
            supporting_metrics=["sales_total_leads"],
            metric_focus=["sales_total_leads"],
            has_metric_evidence=True,
        )

    def test_the_missing_direction_takes_the_last_slot(self) -> None:
        selected = [self._rule(f"neg{i}", "negative") for i in range(4)]
        ranked = selected + [self._rule("pos1", "positive")]
        out = core.ensure_both_directions_are_represented(selected, ranked, self.MIXED)
        self.assertEqual([rule.rule_id for rule in out][-1], "pos1")

    def test_the_lead_is_never_displaced(self) -> None:
        """The month's biggest movement still opens the report, whichever way it went."""
        selected = [self._rule(f"neg{i}", "negative") for i in range(4)]
        ranked = selected + [self._rule("pos1", "positive")]
        out = core.ensure_both_directions_are_represented(selected, ranked, self.MIXED)
        self.assertEqual(out[0].rule_id, "neg0")

    def test_a_one_directional_month_is_left_alone(self) -> None:
        selected = [self._rule(f"neg{i}", "negative") for i in range(4)]
        ranked = selected + [self._rule("pos1", "positive")]
        out = core.ensure_both_directions_are_represented(selected, ranked, self.ONE_WAY)
        self.assertEqual([rule.rule_id for rule in out], [rule.rule_id for rule in selected])

    def test_nothing_is_swapped_in_without_evidence(self) -> None:
        selected = [self._rule(f"neg{i}", "negative") for i in range(4)]
        weak = self._rule("pos1", "positive")
        weak.has_metric_evidence = False
        out = core.ensure_both_directions_are_represented(selected, [*selected, weak], self.MIXED)
        self.assertEqual([rule.rule_id for rule in out], [rule.rule_id for rule in selected])

    def test_an_already_balanced_selection_is_untouched(self) -> None:
        selected = [self._rule("neg0", "negative"), self._rule("pos0", "positive")]
        out = core.ensure_both_directions_are_represented(selected, selected, self.MIXED)
        self.assertEqual([rule.rule_id for rule in out], ["neg0", "pos0"])


class SmsShowsBothSidesTests(unittest.TestCase):
    """The SMS takes three signals, one per family — and they can all be one department.

    A month with deals up 26.7% and conversion up 12 points went out as "service ROI
    down 79.6%, service tasks up 266.7%, and service bookings down to 8": every word
    true, the month itself misrepresented.
    """

    def _signal(self, name: str, family: str, polarity: str) -> core.SmsSignal:
        return core.SmsSignal(
            phrase=f"{name} moved", family=family, polarity=polarity,
            metric=name, quantified_phrase=f"{name} moved",
        )

    def test_the_weakest_slot_goes_to_the_missing_side(self) -> None:
        signals = [
            self._signal("a", "roi", "negative"),
            self._signal("b", "tasks", "negative"),
            self._signal("c", "bookings", "negative"),
        ]
        scored = [(10.0, "d", self._signal("d", "deals", "positive"))]
        out = core._ensure_sms_shows_both_sides(signals, scored)
        self.assertEqual([s.polarity for s in out], ["negative", "negative", "positive"])

    def test_the_leading_signal_is_kept(self) -> None:
        signals = [self._signal("a", "roi", "negative"), self._signal("b", "tasks", "negative")]
        scored = [(10.0, "d", self._signal("d", "deals", "positive"))]
        out = core._ensure_sms_shows_both_sides(signals, scored)
        self.assertEqual(out[0].metric, "a")

    def test_nothing_changes_when_both_sides_are_already_there(self) -> None:
        signals = [self._signal("a", "roi", "negative"), self._signal("b", "deals", "positive")]
        out = core._ensure_sms_shows_both_sides(signals, [])
        self.assertEqual([s.metric for s in out], ["a", "b"])

    def test_no_candidate_means_no_change(self) -> None:
        """A month with nothing good in it still says so."""
        signals = [self._signal("a", "roi", "negative"), self._signal("b", "tasks", "negative")]
        out = core._ensure_sms_shows_both_sides(signals, [])
        self.assertEqual([s.metric for s in out], ["a", "b"])
