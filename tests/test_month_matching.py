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


class AbbreviationTests(unittest.TestCase):
    """Exports switch abbreviations month to month: Sep 2026 one month, Sept 2026 the next."""

    def test_every_abbreviation_of_the_month_matches(self) -> None:
        for spelling in ("Sep", "Sept", "Sept.", "SEPT", "Septem", "September"):
            for year in ("2026", "26"):
                filename = f"{{Dealer}} Monthly Sales Report - {spelling} {year}"
                with self.subTest(filename=filename):
                    self.assertGreater(score(filename, 9, 2026), 0)

    def test_other_months_accept_their_longer_abbreviations(self) -> None:
        for filename, month in (("Sales Report - Janu 2026", 1), ("Sales Report - Febr 2026", 2),
                                ("Sales Report - Augu 2026", 8), ("Sales Report - Novem 2026", 11)):
            with self.subTest(filename=filename):
                self.assertGreater(score(filename, month, 2026), 0)

    def test_an_abbreviation_never_matches_a_neighbouring_month(self) -> None:
        for month in (8, 10):
            self.assertEqual(score("{Dealer} Monthly Sales Report - Sept 2026", month, 2026), 0)
        self.assertEqual(score("Sales Report - June 2026", 7, 2026), 0)
        self.assertEqual(score("Sales Report - Mar 2026", 5, 2026), 0)

    def test_two_letters_is_not_a_month(self) -> None:
        self.assertIsNone(core.month_from_name("Se"))
        self.assertEqual(core.month_from_name("Sept."), 9)

    def test_the_month_is_stripped_from_the_dealership_name(self) -> None:
        name = core.extract_dealership_name_from_filename(
            "{Demo Site 1} Monthly Sales Report - Sept 2026.pdf", "Fallback")
        self.assertEqual(name, "Demo Site 1")

    def test_a_typed_month_label_accepts_the_abbreviation(self) -> None:
        self.assertEqual(core.parse_month_label("Sept 2026"), (9, 2026, "Sep 2026"))
        self.assertEqual(core.parse_month_label("Sep 2026"), (9, 2026, "Sep 2026"))


if __name__ == "__main__":
    unittest.main()
