"""The three breakdown tables the metric parser discards.

Fixtures are written by hand rather than lifted from a report, so each one isolates a
shape that matters: a dealership running more campaigns than the report names, one
running none, one whose campaign lost money, and a pipeline row the extraction
truncated. The demo dealerships cover the ordinary case; these cover the rest.
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from reporting import report_breakdowns as rb


SALES = """\
LEADS, DEALS & CONVERSIONS
NUMBER OF LEADS NUMBER OF DEALS CONVERSION RATIOS
Jun-26 Jul-26 Jun-26 Jul-26 Jun-26 Jul-26
WALK IN 26 21 10 3 38% 14%
PHONE UP 11 9 10 6 91% 67%
INTERNET 18 27 0 5 0% 19%
OTHER 276 34 156 100 57% 294%
TOTAL 331 91 176 114 53% 125%
TARGETED MARKETING
DEALS FROM CAMPAIGNS SUCCESSFUL CAMPAIGNS DEALS ROI
DEALS ROI New Prospect Enquiry 19 $11,663
Jun-26 22 $17,942 Used Prospect Enquiry 6 $7,497
Jul-26 27 $19,113 One Year Anniversary 1 -$47
SALESPIPELINE
Unqualified Contacted Store Visit Demo Write-Up F&I Deals Lost
Jun-26 0 3 1 0 8 9 0 18
Jul-26 0 3 2 7 11 10 40
"""

SERVICE_MANY = """\
SERVICE BOOKINGS FROM
ADDITIONAL SUCCESSFUL CAMPAIGNS RO's ROI
CAMPAIGNS
RO's ROI Closed RO 28 $18,201
Jun-26 51 $30,110 NEW Closed RO 6 $4,891
Jul-26 69 $37,163 Service Reminder 7 $3,541
"""

SERVICE_NONE = """\
SERVICE BOOKINGS FROM
ADDITIONAL SUCCESSFUL CAMPAIGNS RO's ROI
CAMPAIGNS
RO's ROI n/a 0 $0
Jun-26 0 $0 n/a 0 $0
Jul-26 0 $0 n/a 0 $0
"""


class LeadSourceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.rows = {row.source: row for row in rb.parse_lead_sources(SALES)}

    def test_every_source_is_kept_including_other(self) -> None:
        """Other stands for whichever channels a dealership set up beyond the named
        three, so it varies by dealership and is not a residual to discount."""
        self.assertEqual(
            sorted(self.rows), ["Internet", "Other", "Phone Up", "Total", "Walk In"]
        )

    def test_a_channel_improving_inside_a_shrinking_total_is_visible(self) -> None:
        internet = self.rows["Internet"]
        self.assertEqual((internet.leads_previous, internet.leads_current), (18, 27))
        self.assertEqual((internet.conversion_previous, internet.conversion_current), (0.0, 19.0))
        self.assertEqual(self.rows["Total"].leads_current, 91, "while the total fell hard")

    def test_the_collapsing_channel_is_visible_too(self) -> None:
        walk_in = self.rows["Walk In"]
        self.assertEqual((walk_in.conversion_previous, walk_in.conversion_current), (38.0, 14.0))


class PipelineStageTests(unittest.TestCase):
    def setUp(self) -> None:
        self.stages = rb.parse_pipeline_stages(SALES)

    def test_the_whole_funnel_is_read(self) -> None:
        self.assertEqual(
            list(self.stages),
            ["Unqualified", "Contacted", "Store Visit", "Demo", "Write-Up", "F&I", "Deals", "Lost"],
        )

    def test_a_short_row_leaves_a_gap_rather_than_shifting_everything_left(self) -> None:
        """July lists seven numbers against eight headings. Shifting would silently
        report the Deals figure as Lost, which is worse than reporting nothing."""
        self.assertEqual(self.stages["Deals"]["Jul-26"], 40)
        self.assertIsNone(self.stages["Lost"]["Jul-26"])
        self.assertEqual(self.stages["Lost"]["Jun-26"], 18)


class CampaignBlockTests(unittest.TestCase):
    def test_the_named_block_can_fall_short_of_the_stated_total(self) -> None:
        """Three campaigns named against sixty-nine bookings. Anything written from
        this has to read as a share of the total, never as a decomposition of it."""
        block = rb.parse_campaign_block(SERVICE_MANY, kind="service")
        self.assertEqual(block.stated_count, 69)
        self.assertEqual(block.named_count, 41)
        self.assertFalse(block.accounts_for_whole_total)
        self.assertEqual(block.unnamed_count, 28)

    def test_the_month_before_is_read_too(self) -> None:
        block = rb.parse_campaign_block(SERVICE_MANY, kind="service")
        self.assertEqual((block.previous_count, block.previous_roi), (51, 30110))
        self.assertEqual((block.stated_count, block.stated_roi), (69, 37163))

    def test_a_dealership_running_no_campaigns_is_a_real_state(self) -> None:
        block = rb.parse_campaign_block(SERVICE_NONE, kind="service")
        self.assertFalse(block.runs_campaigns)
        self.assertEqual(block.named, [])
        self.assertIsNone(block.leader())

    def test_a_campaign_that_lost_money_is_read_as_negative(self) -> None:
        """The first line of the brief: campaigns which don't perform cost money."""
        block = rb.parse_campaign_block(SALES, kind="sales")
        losing = block.losing()
        self.assertEqual([c.name for c in losing], ["One Year Anniversary"])
        self.assertEqual(losing[0].roi, -47)

    def test_the_leader_is_the_biggest_return_not_the_first_listed(self) -> None:
        block = rb.parse_campaign_block(SALES, kind="sales")
        self.assertEqual(block.leader().name, "New Prospect Enquiry")
        self.assertEqual(block.leader().roi, 11663)

    def test_header_fragments_do_not_become_campaign_names(self) -> None:
        """The two columns share physical lines, so the first campaign arrives with a
        header fragment glued to the front of its name."""
        for block in (
            rb.parse_campaign_block(SALES, kind="sales"),
            rb.parse_campaign_block(SERVICE_MANY, kind="service"),
        ):
            for campaign in block.named:
                self.assertNotRegex(campaign.name, r"(?i)^(RO'?s|DEALS)\b")

    def test_a_month_row_is_never_mistaken_for_a_campaign(self) -> None:
        block = rb.parse_campaign_block(SERVICE_MANY, kind="service")
        self.assertNotIn("Jun-26", " ".join(c.name for c in block.named))
        self.assertEqual(len(block.named), 3)


class NothingIsInventedTests(unittest.TestCase):
    def test_missing_blocks_return_empty_rather_than_guessing(self) -> None:
        self.assertEqual(rb.parse_lead_sources(""), [])
        self.assertEqual(rb.parse_pipeline_stages(""), {})
        block = rb.parse_campaign_block("", kind="sales")
        self.assertFalse(block.runs_campaigns)
        self.assertIsNone(block.stated_count)


if __name__ == "__main__":
    unittest.main()
