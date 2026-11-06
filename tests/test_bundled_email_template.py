"""The web app always uses the email template and branding shipped with the package."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

import email_assembler
import report_engine
from webapp import api


class BundledEmailTemplateTests(unittest.TestCase):
    def test_bundled_template_and_brand_assets_are_present(self) -> None:
        self.assertTrue(email_assembler.HTML_TEMPLATE_PATH.is_file())
        self.assertTrue(email_assembler.TEXT_TEMPLATE_PATH.is_file())
        assets = email_assembler.PACKAGE_DIR / "email_assets"
        for filename in ("banner_top.png", "banner_hero.png", "button_help.png", "footer_logo.png"):
            with self.subTest(filename=filename):
                self.assertTrue((assets / filename).is_file())

    def test_need_help_button_uses_the_customer_service_portal(self) -> None:
        assembled = email_assembler.assemble_dealership_email(
            dealership_name="Demo Site 1",
            month_label="Jul 2026",
            greeting_name="Alex",
            priority_insights=["Insight one", "Insight two", "Insight three", "Insight four"],
            summary_resolved="Summary sentence.",
            to_recipients="",
            cc_recipients="",
            pdf_attachments=[],
            report_reference={},
            assets_dir=email_assembler.PACKAGE_DIR / "email_assets",
        )
        self.assertIn(
            'href="https://www.pentanasolutions.com/customer-service-portal"',
            assembled.html_body,
        )

    def test_settings_and_ui_do_not_offer_template_selection(self) -> None:
        obsolete_keys = {"template_path", "email_template_assets_folder", "email_assembly_mode"}
        self.assertTrue(obsolete_keys.isdisjoint(api.SETTING_KEYS))

        root = Path(__file__).resolve().parents[1]
        app_js = (root / "webapp" / "static" / "app.js").read_text(encoding="utf-8")
        page = (root / "webapp" / "static" / "index.html").read_text(encoding="utf-8")
        self.assertNotIn("Template path", app_js)
        self.assertNotIn("email_template_assets_folder", app_js)
        self.assertNotIn("assembly-mode", app_js + page)
        self.assertNotIn("legacy_msg", app_js + page)
        self.assertFalse(hasattr(api, "validate_template"))

    def test_reporting_ignores_a_legacy_user_template_path(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            accounts = root / "accounts"
            output = root / "output"
            accounts.mkdir()
            output.mkdir()
            artifacts = SimpleNamespace(results=[])

            with (
                mock.patch.object(report_engine.reporting_core, "run_reporting_pipeline", return_value=artifacts),
                mock.patch.object(report_engine, "_set_active_session"),
                mock.patch.object(report_engine, "append_workflow_evidence_record"),
                mock.patch.object(report_engine.email_builder, "load_recipient_mappings", return_value={}),
                mock.patch.object(
                    report_engine.email_builder,
                    "validate_template_file",
                    side_effect=AssertionError("user template should never be validated"),
                ),
            ):
                session = report_engine.run_reporting(
                    accounts_folder=accounts,
                    output_folder=output,
                    month="Jul 2026",
                    template_path=root / "missing.msg",
                )

        self.assertIsNone(session.template_path)
        self.assertEqual(session.email_assembly_mode, "assembler")
        self.assertEqual(session.email_template_assets_folder, email_assembler.PACKAGE_DIR / "email_assets")


if __name__ == "__main__":
    unittest.main()