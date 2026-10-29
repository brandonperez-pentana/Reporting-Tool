"""Month/year matching of source report filenames.

Locks in the filename spellings the pipeline must recognise, and — just as
importantly — the near-miss spellings it must NOT match. The false-positive
cases matter because real source files carry a leading send-date prefix
("2026-05-11_1027 ... - April 2026.pdf"), so a numeric year-first pattern
would silently pick up the wrong month's report.
"""

from __future__ import annotations

import unittest

import dealership_insights_local as core


def score(filename: str, month: int, year: int) -> int:
    return core.month_score(
        filename,
        core.month_variants(month, year),
        core.month_short_year_patterns(month, year),
    )


class MonthMatchingTests(unittest.TestCase):
    def test_full_year_spellings_match(self) -> None:
        for filename in (
            "{Dealer} Monthly Sales Report - July 2026",
            "{Dealer} Monthly Sales Report - Jul 2026",
            "{Dealer} Monthly Service Report - july 2026",
            "{Dealer} Sales 07 2026",
            "{Dealer} Sales 07-2026",
            "{Dealer} Sales 07_2026",
            "{Dealer} Sales 7 2026",
            "{Dealer} Sales 2026 July",
            "{Dealer} Sales 2026 Jul",
        ):
            with self.subTest(filename=filename):
                self.assertGreater(score(filename, 7, 2026), 0)

    def test_two_digit_year_spellings_match(self) -> None:
        for filename in (
            "{Dealer} Sales Jul 26",
            "{Dealer} Sales July 26",
            "{Dealer} Sales Jul-26",
            "{Dealer} Sales July_26",
        ):
            with self.subTest(filename=filename):
                self.assertGreater(score(filename, 7, 2026), 0)

    def test_full_year_outranks_two_digit_year(self) -> None:
        """So an unambiguous file wins when both spellings exist."""
        self.assertGreater(
            score("{Dealer} Sales Jul 2026", 7, 2026),
            score("{Dealer} Sales Jul 26", 7, 2026),
        )

    def test_send_date_prefix_does_not_match_wrong_month(self) -> None:
        """The April report is sent in May; targeting May must not match it."""
        april_report = "2026-05-11_1027 Your Monthly DealerSocket Snapshot - April 2026"
        self.assertEqual(score(april_report, 5, 2026), 0)
        self.assertGreater(score(april_report, 4, 2026), 0)

    def test_day_number_before_year_is_not_a_short_year(self) -> None:
        """'May 26 2025' is the 26th of May, not May 2026."""
        self.assertEqual(score("{Dealer} Sales Report May 26 2025", 5, 2026), 0)

    def test_other_months_and_years_do_not_match(self) -> None:
        self.assertEqual(score("{Dealer} Monthly Sales Report - June 2026", 7, 2026), 0)
        self.assertEqual(score("{Dealer} Monthly Sales Report - July 2025", 7, 2026), 0)
        self.assertEqual(score("{Dealer} Monthly Sales Report - July 2026", 6, 2026), 0)

    def test_full_month_name_not_read_as_short_year(self) -> None:
        """'july 2026' must not be matched by the 'july 26' short-year pattern."""
        patterns = core.month_short_year_patterns(7, 2026)
        normalized = core.normalize_text("{Dealer} Monthly Sales Report - July 2026")
        self.assertFalse(any(p.search(normalized) for p in patterns))


if __name__ == "__main__":
    unittest.main()
