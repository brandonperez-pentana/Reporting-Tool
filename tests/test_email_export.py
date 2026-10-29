"""The manual send route: an email written to disk as .eml.

This exists for the machine where Outlook automation will not run, so the tests avoid
Outlook entirely and check the file itself. What matters is that the file Outlook opens
is the same email it would have drafted — the branding images resolve, both source PDFs
are attached, and it opens ready to send rather than as a message that has arrived.
"""

from __future__ import annotations

import email
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import email_export


def _fixture_files(folder: Path) -> tuple[Path, Path, Path]:
    image = folder / "banner_top.png"
    # Smallest thing that is unambiguously a PNG by signature and by extension.
    image.write_bytes(bytes.fromhex("89504e470d0a1a0a") + b"\x00" * 16)
    sales = folder / "Demo Site 1 - Sales Report - July 26.pdf"
    sales.write_bytes(b"%PDF-1.4\nsales\n")
    service = folder / "Demo Site 1 - Service Report - July 26.pdf"
    service.write_bytes(b"%PDF-1.4\nservice\n")
    return image, sales, service


class BuildMessageTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="eml-test-"))
        self.image, self.sales, self.service = _fixture_files(self.tmp)

    def _message(self, **overrides):
        kwargs = dict(
            subject="Your Monthly DealerSocket Snapshot - July 2026",
            to_recipients="sam@example.com",
            cc_recipients="lead@example.com",
            body_text="Hi Sam,\n\nAttached are your reports.",
            html_body='<html><body><img src="cid:banner_top"><p>Hi Sam,</p></body></html>',
            inline_assets=[("banner_top", self.image)],
            attachments=[self.sales, self.service],
        )
        kwargs.update(overrides)
        message, attached, inline = email_export.build_message(**kwargs)
        return message, attached, inline

    def test_opens_as_a_draft_not_as_received_mail(self) -> None:
        """X-Unsent is the whole point: without it Outlook opens a read-only message."""
        message, _, _ = self._message()
        self.assertEqual(message["X-Unsent"], "1")

    def test_headers_carry_the_recipients(self) -> None:
        message, _, _ = self._message()
        self.assertEqual(message["To"], "sam@example.com")
        self.assertEqual(message["Cc"], "lead@example.com")
        self.assertIn("July 2026", message["Subject"])

    def test_no_from_header_when_no_sender_given(self) -> None:
        """Outlook fills the account in itself; a guessed From would send as the wrong one."""
        message, _, _ = self._message()
        self.assertIsNone(message["From"])

    def test_inline_image_is_related_to_the_html_and_keeps_its_cid(self) -> None:
        message, _, inline = self._message()
        self.assertEqual(inline, 1)
        cids = [part.get("Content-ID") for part in message.walk() if part.get("Content-ID")]
        # The html says src="cid:banner_top", so the part has to answer to <banner_top>.
        self.assertIn("<banner_top>", cids)
        related = [p for p in message.walk() if p.get_content_type() == "multipart/related"]
        self.assertTrue(related, "inline images must sit in a multipart/related with the html")

    def test_both_pdfs_are_attached_under_their_own_names(self) -> None:
        message, attached, _ = self._message()
        self.assertEqual(
            attached,
            ["Demo Site 1 - Sales Report - July 26.pdf", "Demo Site 1 - Service Report - July 26.pdf"],
        )
        pdfs = [p for p in message.walk() if (p.get_filename() or "").endswith(".pdf")]
        self.assertEqual(len(pdfs), 2)
        self.assertTrue(all(p.get_content_type() == "application/pdf" for p in pdfs))

    def test_a_missing_attachment_is_reported_rather_than_faked(self) -> None:
        """A PDF deleted since the run must not silently vanish from the operator's view."""
        _, attached, _ = self._message(attachments=[self.sales, self.tmp / "gone.pdf"])
        self.assertEqual(attached, ["Demo Site 1 - Sales Report - July 26.pdf"])

    def test_plain_text_alternative_is_present(self) -> None:
        message, _, _ = self._message()
        plain = [p for p in message.walk() if p.get_content_type() == "text/plain"]
        self.assertTrue(plain)
        self.assertIn("Hi Sam", plain[0].get_payload(decode=True).decode("utf-8"))


class FilenameTests(unittest.TestCase):
    def test_characters_windows_refuses_are_replaced(self) -> None:
        stem = email_export.safe_filename_stem('A/B:C*D?"E<F>G|H')
        for bad in '/\\:*?"<>|':
            self.assertNotIn(bad, stem)

    def test_empty_name_falls_back_rather_than_producing_a_dotfile(self) -> None:
        self.assertEqual(email_export.safe_filename_stem("   ."), "email")

    def test_exporting_twice_never_overwrites(self) -> None:
        folder = Path(tempfile.mkdtemp(prefix="eml-unique-"))
        first = email_export.unique_path(folder, "Demo Site 1 - Jul 2026")
        first.write_text("one", encoding="utf-8")
        second = email_export.unique_path(folder, "Demo Site 1 - Jul 2026")
        self.assertNotEqual(first, second)
        self.assertTrue(second.name.endswith("(2).eml"))


class WriteEmlTests(unittest.TestCase):
    def test_file_round_trips_through_a_mail_parser(self) -> None:
        tmp = Path(tempfile.mkdtemp(prefix="eml-write-"))
        image, sales, service = _fixture_files(tmp)
        out = tmp / "out"
        exported = email_export.write_eml(
            out,
            "Demo Site 1 - Jul 2026",
            subject="Snapshot",
            to_recipients="sam@example.com",
            body_text="Hi Sam,",
            html_body='<html><body><img src="cid:banner_top"></body></html>',
            inline_assets=[("banner_top", image)],
            attachments=[sales, service],
        )
        self.assertTrue(exported.path.is_file())
        self.assertEqual(exported.path.suffix, ".eml")
        self.assertEqual(exported.inline_asset_count, 1)
        self.assertEqual(len(exported.attachment_names), 2)

        parsed = email.message_from_bytes(exported.path.read_bytes())
        self.assertEqual(parsed["X-Unsent"], "1")
        filenames = {p.get_filename() for p in parsed.walk() if p.get_filename()}
        self.assertIn("Demo Site 1 - Sales Report - July 26.pdf", filenames)
        self.assertIn("Demo Site 1 - Service Report - July 26.pdf", filenames)

    def test_folder_is_created_if_it_does_not_exist(self) -> None:
        tmp = Path(tempfile.mkdtemp(prefix="eml-mkdir-"))
        target = tmp / "does" / "not" / "exist"
        exported = email_export.write_eml(
            target, "X", subject="s", to_recipients="a@example.com", body_text="b"
        )
        self.assertTrue(exported.path.is_file())


if __name__ == "__main__":
    unittest.main()
