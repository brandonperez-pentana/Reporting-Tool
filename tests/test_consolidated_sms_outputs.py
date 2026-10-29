import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
import sys

sys.path.insert(0, str(ROOT))

from tools.generate_consolidated_sms_outputs import (
    DEFAULT_OUTPUT_BASENAME,
    SmsRow,
    build_consolidated_rows,
    format_sms_block,
    group_rows_by_month,
    monthly_output_paths,
    safe_month_filename,
    validate_written_exports,
    write_markdown_output,
    write_monthly_outputs,
    write_text_output,
    ValidationReport,
)


class ConsolidatedSmsExportTests(unittest.TestCase):
    def test_safe_month_filename(self) -> None:
        self.assertEqual(safe_month_filename("Jun 2026"), "June_2026")
        self.assertEqual(safe_month_filename("Mar 2026"), "March_2026")

    def test_format_sms_block_preserves_line_breaks(self) -> None:
        row = SmsRow(
            dealership="Demo Site 1",
            month_label="Jun 2026",
            sms_text="DealerSocket June Insights:\nquantified summary here.",
            char_count=52,
            source_file="C:\\source.txt",
            over_limit=False,
        )
        md_block = "\n".join(format_sms_block(row, block_format="md"))
        self.assertIn("### Demo Site 1", md_block)
        self.assertIn("DealerSocket June Insights:\nquantified summary here.", md_block)
        self.assertIn("Character Count: 52", md_block)
        self.assertIn("---", md_block)

        txt_block = "\n".join(format_sms_block(row, block_format="txt"))
        self.assertIn("Demo Site 1", txt_block)
        self.assertIn("DealerSocket June Insights:\nquantified summary here.", txt_block)
        self.assertIn("------------------------------------------------------------", txt_block)

    def test_monthly_and_full_exports_preserve_sms_text(self) -> None:
        rows = [
            SmsRow(
                dealership="Dealer A",
                month_label="Jun 2026",
                sms_text="DealerSocket June Insights:\nleads rose 10%.",
                char_count=len("DealerSocket June Insights:\nleads rose 10%."),
                source_file="C:\\Insights Output\\Dealer A - Jun 2026.txt",
                over_limit=False,
            ),
            SmsRow(
                dealership="Dealer B",
                month_label="May 2026",
                sms_text="DealerSocket May Insights:\nservice ROI improved 5%.",
                char_count=len("DealerSocket May Insights:\nservice ROI improved 5%."),
                source_file="C:\\Insights Output\\Dealer B - May 2026.txt",
                over_limit=False,
            ),
        ]
        rows_by_month = group_rows_by_month(rows)

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            output_md = root / f"{DEFAULT_OUTPUT_BASENAME}.md"
            output_txt = root / f"{DEFAULT_OUTPUT_BASENAME}.txt"

            write_markdown_output(rows_by_month, output_md, title="Consolidated SMS Outputs")
            write_text_output(rows_by_month, output_txt, title="Consolidated SMS Outputs")
            report = ValidationReport(consolidation_complete=True)
            write_monthly_outputs(rows_by_month, root, report)

            validate_written_exports(rows, rows_by_month, output_md, output_txt, report)

            self.assertEqual(report.spacing_separator_validation, "PASS")
            self.assertEqual(report.sms_text_preservation_failures, [])
            self.assertEqual(report.character_count_failures, [])
            june_paths = monthly_output_paths(root, "Jun 2026")
            self.assertTrue(june_paths["md"].exists())
            self.assertTrue(june_paths["txt"].exists())
            self.assertTrue(june_paths["docx"].exists())
            self.assertIn("DealerSocket June Insights:\nleads rose 10%.", june_paths["md"].read_text(encoding="utf-8"))

    def test_build_consolidated_rows_is_read_only(self) -> None:
        report = ValidationReport()
        rows = build_consolidated_rows([], report)
        self.assertEqual(rows, [])
        self.assertFalse(report.production_reports_modified)


if __name__ == "__main__":
    unittest.main()
