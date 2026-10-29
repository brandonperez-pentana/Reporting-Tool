"""Local validation for email workflow upgrade (Jinja, contacts, SMS exclusion, context)."""

from __future__ import annotations

import base64
import json
import logging
import tempfile
import unittest
from pathlib import Path
import email_assembler
import email_builder

FOUR_INSIGHTS = ["Insight A", "Insight B", "Insight C", "Insight D"]


class OutlookHtmlDocumentTests(unittest.TestCase):
    def test_ensure_outlook_html_wraps_fragment(self) -> None:
        out = email_builder.ensure_outlook_html_document('<p style="margin:0">Hi</p>')
        low = out.lower()
        self.assertIn("<html", low)
        self.assertIn("charset", low)
        self.assertIn('<p style="margin:0">Hi</p>', out)

    def test_ensure_outlook_html_preserves_full_document(self) -> None:
        doc = "<!DOCTYPE html><html><head><meta charset=\"utf-8\"></head><body><table><tr><td>x</td></tr></table></body></html>"
        out = email_builder.ensure_outlook_html_document(doc)
        self.assertIn("x", out)
        self.assertIn("<table", out.lower())


class EmailTemplateTests(unittest.TestCase):
    def test_jinja_renders_required_variables(self) -> None:
        ctx = email_builder.build_email_context_object(
            dealership_name="ABC Motors",
            greeting_name="Sam",
            month_label="Apr 2026",
            priority_insights=FOUR_INSIGHTS,
            summary="Overall good month.",
            sender_name="Sam Taylor",
            sender_title="CRM Consultant",
        )
        for key in email_builder.PREFERRED_EMAIL_REQUIRED_VARS:
            self.assertIn(key, ctx)
            self.assertTrue(str(ctx[key]).strip(), f"missing value for {key}")
        rendered = email_builder._render_jinja_template(email_builder.DEFAULT_BODY_TEMPLATE, {**ctx, **{k.upper(): v for k, v in ctx.items()}})
        self.assertIn("Sam", rendered)
        self.assertIn("Apr 2026", rendered)
        self.assertNotIn("SMS", rendered.lower())

    def test_missing_jinja_variable_raises(self) -> None:
        with self.assertRaises(email_builder.TemplateValidationError):
            email_builder._render_jinja_template("Hello {{ unknown_var }}", {"name": "x"})

    def test_sms_omitted_from_email_render(self) -> None:
        out = email_builder.render_email(
            template_path=None,
            dealership_name="Test Dealer",
            month_label="Jan 2026",
            executive_summary="x",
            priority_insights=["p1", "p2", "p3", "p4"],
            summary="sum",
            sms_version="THIS SMS MUST NOT APPEAR IN BODY",
            sentiment="pos",
            review_reason_tags=[],
            greeting_name="Alex",
            omit_sms_from_email=True,
        )
        self.assertNotIn("THIS SMS", out.preview_text)
        self.assertNotIn("THIS SMS", out.html_body)


class ContactsTests(unittest.TestCase):
    def test_contacts_csv_load_and_match(self) -> None:
        sample = (
            "dealership,contact_name,email,cc,active,notes,last_updated\n"
            "Foo Inc,Jane,j@foo.com,cc@foo.com,yes,,\n"
        )
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "c.csv"
            p.write_text(sample, encoding="utf-8")
            m = email_builder.load_recipient_mappings(p)
            r = email_builder.resolve_contact_for_dealership("Foo Inc", m)
            self.assertTrue(r.matched)
            self.assertTrue(r.ready)
            self.assertEqual(r.to, "j@foo.com")
            self.assertEqual(r.cc, "cc@foo.com")
            self.assertEqual(r.contact_name, "Jane")
            self.assertEqual(r.greeting_name, "Jane")

    def test_case_insensitive_dealership_match(self) -> None:
        sample = "dealership,contact_name,email,cc,active,notes,last_updated\nFoo Inc,Jane,j@foo.com,,yes,,\n"
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "c.csv"
            p.write_text(sample, encoding="utf-8")
            m = email_builder.load_recipient_mappings(p)
            r = email_builder.resolve_contact_for_dealership("FOO INC", m)
            self.assertTrue(r.matched)
            self.assertEqual(r.greeting_name, "Jane")

    def test_contact_name_in_jinja_context(self) -> None:
        ctx = email_builder.build_email_context_object(
            dealership_name="Foo Inc",
            greeting_name="Jane",
            month_label="Jan 2026",
            priority_insights=FOUR_INSIGHTS,
            summary="s",
            email="j@foo.com",
            cc="c@foo.com",
        )
        self.assertEqual(ctx["name"], "Jane")
        self.assertEqual(ctx["dealership"], "Foo Inc")
        self.assertEqual(ctx["email"], "j@foo.com")
        self.assertEqual(ctx["cc"], "c@foo.com")

    def test_manual_override_keeps_contact_greeting(self) -> None:
        sample = "dealership,contact_name,email,cc,active,notes,last_updated\nFoo Inc,Jane,j@foo.com,,yes,,\n"
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "c.csv"
            p.write_text(sample, encoding="utf-8")
            m = email_builder.load_recipient_mappings(p)
            r = email_builder.resolve_contact_for_dealership(
                "Foo Inc",
                m,
                manual_to="other@example.com",
                manual_cc="",
            )
            self.assertEqual(r.to, "other@example.com")
            self.assertEqual(r.greeting_name, "Jane")
            self.assertEqual(r.contact_name, "Jane")

    def test_inactive_contact_uses_dealership_greeting_not_ready(self) -> None:
        sample = (
            "dealership,contact_name,email,cc,active,notes,last_updated\n"
            "Foo Inc,Bob,b@foo.com,,no,,\n"
        )
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "c.csv"
            p.write_text(sample, encoding="utf-8")
            m = email_builder.load_recipient_mappings(p)
            r = email_builder.resolve_contact_for_dealership("Foo Inc", m)
            self.assertTrue(r.matched)
            self.assertFalse(r.ready)
            self.assertEqual(r.greeting_name, "Foo Inc")
            self.assertEqual(r.to, "")

    def test_missing_contact_not_ready(self) -> None:
        r = email_builder.resolve_contact_for_dealership("Unknown Dealer", {})
        self.assertFalse(r.matched)
        self.assertFalse(r.ready)

    def test_invalid_email_inactive_allowed(self) -> None:
        row = {
            "dealership": "X",
            "contact_name": "Y",
            "email": "not-an-email",
            "cc": "",
            "active": "no",
            "notes": "",
            "last_updated": "",
        }
        errs = email_builder.validate_contact_row_for_save(row)
        self.assertEqual(errs, [])


class EmailContextTests(unittest.TestCase):
    def test_email_context_insight_slots_and_summary(self) -> None:
        ctx = email_builder.build_email_context_object(
            dealership_name="D",
            greeting_name="G",
            month_label="M",
            priority_insights=FOUR_INSIGHTS,
            summary="S",
        )
        self.assertEqual(ctx["insight_1"], "Insight A")
        self.assertEqual(ctx["point_4"], "Insight D")

    def test_normalize_priority_insights_variants(self) -> None:
        bullet = "- First\n• Second"
        lines = email_builder.extract_priority_insight_lines(bullet)
        self.assertEqual(lines, ["First", "Second"])
        d = {"2": "B", "1": "A", "3": "C", "4": "D"}
        i1, i2, i3, i4 = email_builder.normalize_priority_insights_to_four(d)
        self.assertEqual(i1, "A")
        self.assertEqual(i2, "B")
        self.assertEqual(i4, "D")

    def test_resolve_summary_report_reference_summary_key(self) -> None:
        s = email_builder.resolve_summary_for_email(
            summary="",
            report_summary="",
            executive_summary="",
            report_reference={"Summary": "From reference."},
        )
        self.assertEqual(s, "From reference.")

    def test_missing_contact_still_renders_body(self) -> None:
        out = email_builder.render_email(
            template_path=None,
            dealership_name="Lonely Motors",
            month_label="Jan 2026",
            executive_summary="Exec",
            priority_insights=FOUR_INSIGHTS,
            summary="Month was steady.",
            sms_version="SMS HIDDEN",
            sentiment="pos",
            review_reason_tags=[],
            greeting_name="",
            omit_sms_from_email=True,
        )
        self.assertIn("Lonely Motors", out.preview_text)
        self.assertIn("Insight A", out.preview_text)
        self.assertIn("Month was steady.", out.preview_text)
        self.assertNotIn("SMS HIDDEN", out.preview_text)
        self.assertIn("Insight A", out.html_body)
        self.assertIn("Month was steady.", out.html_body)


class EmailAssemblerTests(unittest.TestCase):
    def test_assembler_html_and_plain_contain_insights_and_summary(self) -> None:
        assembled = email_assembler.assemble_dealership_email(
            dealership_name="Test Dealer",
            month_label="Mar 2026",
            greeting_name="Pat",
            priority_insights=FOUR_INSIGHTS,
            summary_resolved="Month summary sentence.",
            to_recipients="",
            cc_recipients="",
            pdf_attachments=[],
            report_reference={"summary": "x"},
            assets_dir=None,
        )
        self.assertIn("Insight A", assembled.html_body)
        self.assertIn("Insight B", assembled.html_body)
        self.assertIn("Insight D", assembled.html_body)
        self.assertIn("Month summary sentence.", assembled.html_body)
        self.assertIn("Insight A", assembled.plain_text_body)
        self.assertIn("Insight B", assembled.plain_text_body)
        self.assertIn("Month summary sentence.", assembled.plain_text_body)

    def test_no_sign_off_at_all_when_the_operator_has_not_set_one(self) -> None:
        """A copy handed to someone else must not sign their mail with a stranger's name.

        Blank is the shipped default, so the whole block goes — not "Best regards," over
        an empty line where a name would have been.
        """
        assembled = email_assembler.assemble_dealership_email(
            dealership_name="Test Dealer",
            month_label="Mar 2026",
            greeting_name="Pat",
            priority_insights=FOUR_INSIGHTS,
            summary_resolved="Month summary sentence.",
            to_recipients="",
            cc_recipients="",
            pdf_attachments=[],
            report_reference={},
            assets_dir=None,
        )
        self.assertNotIn("Best regards", assembled.html_body)
        self.assertNotIn("Best regards", assembled.plain_text_body)
        self.assertFalse(assembled.plain_text_body.rstrip().endswith(","))

    def test_a_sign_off_appears_once_the_operator_sets_one(self) -> None:
        assembled = email_assembler.assemble_dealership_email(
            dealership_name="Test Dealer",
            month_label="Mar 2026",
            greeting_name="Pat",
            priority_insights=FOUR_INSIGHTS,
            summary_resolved="Month summary sentence.",
            to_recipients="",
            cc_recipients="",
            pdf_attachments=[],
            report_reference={},
            assets_dir=None,
            sender_name="Sam Taylor",
            sender_title="CRM Consultant",
        )
        for body in (assembled.html_body, assembled.plain_text_body):
            self.assertIn("Best regards", body)
            self.assertIn("Sam Taylor", body)
            self.assertIn("CRM Consultant", body)

    def test_assembler_resolves_local_asset_uri(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "logo.png").write_bytes(b"\x89PNG\r\n\x1a\n")
            (root / "email_asset_manifest.json").write_text(
                json.dumps({"banner_top": "logo.png"}),
                encoding="utf-8",
            )
            assembled = email_assembler.assemble_dealership_email(
                dealership_name="D",
                month_label="M",
                greeting_name="G",
                priority_insights=FOUR_INSIGHTS,
                summary_resolved="S",
                to_recipients="",
                cc_recipients="",
                pdf_attachments=[],
                report_reference={},
                assets_dir=root,
            )
            # Preview HTML embeds the asset's real bytes as a data: URI, so the
            # preview renders standalone rather than depending on the preview
            # widget resolving file:// or cid: references.
            expected_data_uri = "data:image/png;base64," + base64.b64encode(
                (root / "logo.png").read_bytes()
            ).decode("ascii")
            self.assertIn(expected_data_uri, assembled.html_body)
            self.assertNotIn("cid:", assembled.html_body)
            # Outlook still uses cid: plus a real inline attachment.
            self.assertEqual(len(assembled.inline_assets), 1)
            self.assertEqual(assembled.inline_assets[0][1].name, "logo.png")
            self.assertIn("cid:banner_top", assembled.html_body_outlook)

    def test_assembler_greeting_falls_back_to_dealership(self) -> None:
        assembled = email_assembler.assemble_dealership_email(
            dealership_name="Solo Motors",
            month_label="Jan 2026",
            greeting_name="Solo Motors",
            priority_insights=FOUR_INSIGHTS,
            summary_resolved="Sum.",
            to_recipients="t@solo.example",
            cc_recipients="c@solo.example",
            pdf_attachments=[],
            report_reference={},
            assets_dir=None,
        )
        self.assertIn("Hi Solo Motors", assembled.plain_text_body)
        self.assertEqual(assembled.email_context.get("name"), "Solo Motors")
        self.assertEqual(assembled.email_context.get("email"), "t@solo.example")
        self.assertEqual(assembled.email_context.get("cc"), "c@solo.example")

    def test_assembler_contact_name_in_greeting(self) -> None:
        assembled = email_assembler.assemble_dealership_email(
            dealership_name="Demo Site 1",
            month_label="Jan 2026",
            greeting_name="Alex Smith",
            priority_insights=FOUR_INSIGHTS,
            summary_resolved="Sum.",
            to_recipients="alex@example.com",
            cc_recipients="",
            pdf_attachments=[],
            report_reference={},
            assets_dir=None,
        )
        self.assertIn("Hi Alex Smith", assembled.plain_text_body)
        self.assertEqual(assembled.email_context.get("name"), "Alex Smith")

    def test_assembler_sms_never_in_body(self) -> None:
        assembled = email_assembler.assemble_dealership_email(
            dealership_name="D",
            month_label="M",
            greeting_name="G",
            priority_insights=FOUR_INSIGHTS,
            summary_resolved="Y",
            to_recipients="",
            cc_recipients="",
            pdf_attachments=[],
            report_reference={},
            assets_dir=None,
        )
        self.assertTrue(assembled.assembler_metadata.get("sms_excluded"))
        self.assertNotIn("Conversion up to", assembled.html_body)

    def test_legacy_injection_emits_warning(self) -> None:
        with self.assertLogs(logging.getLogger(), level="WARNING") as cm:
            email_builder.inject_insights_and_summary_into_html_if_missing(
                "<html><body><p><b>Key Insights:</b></p><p>After</p></body></html>",
                insight_1="Filled by legacy",
                insight_2="",
                insight_3="",
                insight_4="",
                summary="",
            )
        self.assertTrue(any("legacy injection fallback" in m for m in cm.output))


class WordHtmlInjectionTests(unittest.TestCase):
    """Regression: Outlook .msg HTML often has static 'Key Insights:' with no {{ insight_* }} placeholders."""

    def test_inject_insights_into_html_without_jinja_slots(self) -> None:
        html = (
            "<p>Hi X</p><p><strong><span>Key Insights:</span></strong></p>"
            "<p><span><br><br>If you have questions</span></p>"
        )
        out = email_builder.inject_insights_and_summary_into_html_if_missing(
            html,
            insight_1="First insight line",
            insight_2="Second insight",
            insight_3=email_builder.INSIGHT_SLOT_PAD,
            insight_4=email_builder.INSIGHT_SLOT_PAD,
            summary="Closing summary sentence.",
        )
        self.assertIn("First insight line", out)
        self.assertIn("Second insight", out)
        self.assertIn("Closing summary sentence.", out)

    def test_no_double_inject_when_first_insight_already_in_body(self) -> None:
        html = "<p>Key Insights:</p><ul><li>First insight line</li></ul><p>Rest</p>"
        out = email_builder.inject_insights_and_summary_into_html_if_missing(
            html,
            insight_1="First insight line",
            insight_2="Second",
            insight_3="",
            insight_4="",
            summary="",
        )
        self.assertEqual(out.count("<ul>"), 1)

    def test_render_html_file_without_insight_jinja_inserts_content(self) -> None:
        tpl = (
            "<html><body><p>Hi {{ name }},</p>"
            "<p><strong>Key Insights:</strong></p>"
            "<p><br/>Tail after insights</p></body></html>"
        )
        with tempfile.NamedTemporaryFile(mode="w", suffix=".html", delete=False, encoding="utf-8") as handle:
            handle.write(tpl)
            path = handle.name
        try:
            out = email_builder.render_email(
                template_path=path,
                dealership_name="D",
                month_label="Mar 2026",
                executive_summary="e",
                priority_insights=FOUR_INSIGHTS,
                summary="Gamma summary for the month.",
                sms_version="SMS MUST NOT APPEAR",
                sentiment="x",
                review_reason_tags=[],
                greeting_name="Pat",
            )
            self.assertIn("Insight A", out.preview_text)
            self.assertIn("Insight B", out.preview_text)
            self.assertIn("Gamma summary for the month.", out.preview_text)
            self.assertNotIn("SMS MUST NOT APPEAR", out.preview_text)
            self.assertNotIn("SMS MUST NOT APPEAR", out.html_body)
        finally:
            Path(path).unlink(missing_ok=True)


if __name__ == "__main__":
    unittest.main()
