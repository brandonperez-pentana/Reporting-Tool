"""Contacts that resolve to the right people, and BCC reaching the message.

Two failures motivated these. A second contacts row for a dealership used to overwrite
the first outright, so adding a colleague silently removed whoever was already there —
no error, and only visible when the email went to one person instead of two. And BCC
did not exist at all, so a copy to a manager had to go in CC where the dealership saw it.
"""

from __future__ import annotations

import email
import sys
import tempfile
import unittest
from unittest import mock
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import email_builder
import email_export


def _contacts_file(body: str) -> Path:
    folder = Path(tempfile.mkdtemp(prefix="contacts-"))
    path = folder / "contacts.csv"
    path.write_text(body, encoding="utf-8")
    return path


HEADER = "dealership,contact_name,email,cc,bcc,active,notes,last_updated\n"


class MergedContactRowTests(unittest.TestCase):
    def _resolve(self, body: str, dealership: str = "Demo Site 1"):
        mapping = email_builder.load_recipient_mappings(_contacts_file(body))
        return email_builder.resolve_contact_for_dealership(dealership, mapping)

    def test_two_rows_for_one_dealership_both_receive_the_email(self) -> None:
        resolved = self._resolve(
            HEADER
            + "Demo Site 1,Sam,sam@example.com,,,yes,,2026-01-01\n"
            + "Demo Site 1,Pat,pat@example.com,,,yes,,2026-01-01\n"
        )
        self.assertIn("sam@example.com", resolved.to)
        self.assertIn("pat@example.com", resolved.to, "the second row must not overwrite the first")

    def test_the_greeting_names_everyone_it_goes_to(self) -> None:
        resolved = self._resolve(
            HEADER
            + "Demo Site 1,Sam,sam@example.com,,,yes,,2026-01-01\n"
            + "Demo Site 1,Pat,pat@example.com,,,yes,,2026-01-01\n"
        )
        self.assertEqual(resolved.greeting_name, "Sam and Pat")

    def test_cc_and_bcc_are_combined_across_rows(self) -> None:
        resolved = self._resolve(
            HEADER
            + "Demo Site 1,Sam,sam@example.com,cc1@example.com,bcc1@example.com,yes,,2026-01-01\n"
            + "Demo Site 1,Pat,pat@example.com,cc2@example.com,bcc2@example.com,yes,,2026-01-01\n"
        )
        for address in ("cc1@example.com", "cc2@example.com"):
            self.assertIn(address, resolved.cc)
        for address in ("bcc1@example.com", "bcc2@example.com"):
            self.assertIn(address, resolved.bcc)

    def test_the_same_address_twice_is_not_sent_twice(self) -> None:
        resolved = self._resolve(
            HEADER
            + "Demo Site 1,Sam,sam@example.com,shared@example.com,,yes,,2026-01-01\n"
            + "Demo Site 1,Pat,pat@example.com,SHARED@example.com,,yes,,2026-01-01\n"
        )
        self.assertEqual(resolved.cc.lower().count("shared@example.com"), 1)

    def test_a_single_row_is_unchanged(self) -> None:
        """The merge must not disturb the ordinary case."""
        resolved = self._resolve(HEADER + "Demo Site 1,Sam,sam@example.com,,,yes,,2026-01-01\n")
        self.assertEqual(resolved.to, "sam@example.com")
        self.assertEqual(resolved.greeting_name, "Sam")
        self.assertEqual(resolved.cc, "")

    def test_a_file_written_before_bcc_existed_still_loads(self) -> None:
        older = "dealership,contact_name,email,cc,active,notes,last_updated\n"
        resolved = self._resolve(older + "Demo Site 1,Sam,sam@example.com,,yes,,2026-01-01\n")
        self.assertEqual(resolved.to, "sam@example.com")
        self.assertEqual(resolved.bcc, "")


class ContactFileShapeTests(unittest.TestCase):
    def test_bcc_sits_beside_cc_in_the_saved_file(self) -> None:
        self.assertIn("bcc", email_builder.CONTACT_CSV_FIELDNAMES)
        fields = email_builder.CONTACT_CSV_FIELDNAMES
        self.assertEqual(fields[fields.index("cc") + 1], "bcc")

    def test_hidden_columns_survive_a_round_trip(self) -> None:
        """The app stopped showing notes and active; it must not therefore delete them."""
        folder = Path(tempfile.mkdtemp(prefix="contacts-rt-"))
        path = folder / "contacts.csv"
        email_builder.save_contact_rows(
            path,
            [{
                "dealership": "Demo Site 1", "contact_name": "Sam",
                "email": "sam@example.com", "cc": "", "bcc": "boss@example.com",
                "active": "no", "notes": "left the business", "last_updated": "2026-01-01",
            }],
        )
        rows = email_builder.load_contact_rows(path)
        self.assertEqual(rows[0]["notes"], "left the business")
        self.assertEqual(rows[0]["active"], "no")
        self.assertEqual(rows[0]["bcc"], "boss@example.com")


    def test_no_contacts_file_ships_inside_the_package(self) -> None:
        """A shipped contacts.csv turns installing an update into data loss.

        Extracting a new build over a copy in use is how an operator installs one, and a
        header-only file in the archive replaced every contact they had entered with
        nothing, without saying so.
        """
        script = Path(__file__).resolve().parent.parent / "tools" / "build_web_package.py"
        if not script.is_file():
            # The build script is not shipped inside the package it builds.
            self.skipTest("build script not present in this copy")
        source = script.read_text(encoding="utf-8")
        self.assertNotIn("EXAMPLE_CONTACTS", source)
        self.assertIn("data/contacts.csv must not ship", source)

    def test_the_contacts_file_is_created_on_demand_with_the_right_header(self) -> None:
        from webapp import api

        folder = Path(tempfile.mkdtemp(prefix="contacts-"))
        created = api.ensure_contacts_file(folder / "contacts.csv")
        self.assertTrue(created.is_file())
        header = created.read_text(encoding="utf-8").splitlines()[0].split(",")
        self.assertEqual(header, list(email_builder.CONTACT_CSV_FIELDNAMES))

    def test_creating_it_never_touches_a_file_already_there(self) -> None:
        """The whole point: a reinstall must not empty the operator's contacts."""
        from webapp import api

        folder = Path(tempfile.mkdtemp(prefix="contacts-"))
        existing = folder / "contacts.csv"
        email_builder.save_contact_rows(existing, [
            {"dealership": "Demo Site 1", "contact_name": "Sam",
             "email": "sam@example.com", "active": "yes"},
        ])
        before = existing.read_bytes()
        api.ensure_contacts_file(existing)
        self.assertEqual(existing.read_bytes(), before)
        self.assertEqual(len(email_builder.load_contact_rows(existing)), 1)


class BccOnTheWireTests(unittest.TestCase):
    def test_bcc_reaches_the_exported_message(self) -> None:
        exported = email_export.write_eml(
            Path(tempfile.mkdtemp(prefix="bcc-")),
            "Demo Site 1 - Aug 2026",
            subject="Snapshot",
            to_recipients="sam@example.com",
            cc_recipients="cc@example.com",
            bcc_recipients="hidden@example.com",
            body_text="Hi Sam,",
        )
        parsed = email.message_from_bytes(exported.path.read_bytes())
        self.assertEqual(parsed["Bcc"], "hidden@example.com")
        self.assertEqual(parsed["Cc"], "cc@example.com")

    def test_no_bcc_header_when_there_is_no_bcc(self) -> None:
        """An empty Bcc: header would be noise in every message that does not use one."""
        exported = email_export.write_eml(
            Path(tempfile.mkdtemp(prefix="bcc-none-")),
            "Demo Site 1 - Aug 2026",
            subject="Snapshot",
            to_recipients="sam@example.com",
            body_text="Hi Sam,",
        )
        parsed = email.message_from_bytes(exported.path.read_bytes())
        self.assertIsNone(parsed["Bcc"])


class DefaultBccTests(unittest.TestCase):
    """A standing BCC list applied to every email, on top of the contact's own."""

    def _with_default(self, default: str, row_bcc: str) -> str:
        import report_engine
        with mock.patch.object(report_engine, "load_settings", return_value={"default_bcc": default}):
            return report_engine._bcc_with_defaults(row_bcc)

    def test_the_default_list_is_added_to_a_contact_bcc(self) -> None:
        merged = self._with_default("a@example.com; b@example.com", "own@example.com")
        for address in ("own@example.com", "a@example.com", "b@example.com"):
            self.assertIn(address, merged)

    def test_an_address_in_both_is_not_bcc_d_twice(self) -> None:
        merged = self._with_default("shared@example.com", "SHARED@example.com")
        self.assertEqual(merged.lower().count("shared@example.com"), 1)

    def test_no_default_leaves_the_contact_untouched(self) -> None:
        self.assertEqual(self._with_default("", "own@example.com"), "own@example.com")

    def test_the_default_applies_even_with_no_contact_bcc(self) -> None:
        self.assertEqual(self._with_default("a@example.com", ""), "a@example.com")

    def test_the_shipped_default_is_empty(self) -> None:
        """These are the operator's own colleagues; they must not travel in a package."""
        import settings_manager

        self.assertEqual(settings_manager.DEFAULT_SETTINGS["default_bcc"], "")


class ContactsPanelTests(unittest.TestCase):
    """The dealership picker must not rebuild the panel from inside its own handler.

    Chrome fires `change` on every arrow-key move through an open dropdown. Re-rendering
    on the first one destroyed the <select> being used, focus fell to <body>, and every
    later keystroke went nowhere — so the name that stuck was whichever option the first
    event landed on, not the one chosen. With 23 dealerships that reads as the wrong
    contact details attached to the wrong dealership.
    """

    def test_saving_returns_the_file_so_the_panel_can_show_what_landed(self) -> None:
        """Closing the panel on save hid the outcome.

        An entry that never reached the file looked saved, because the only feedback was
        a count in a status bar the operator had already stopped reading. The response
        now carries the saved contacts and the dealership names, and the panel re-renders
        from them without closing.
        """
        from webapp import api

        folder = Path(tempfile.mkdtemp(prefix="save-"))
        path = folder / "contacts.csv"
        path.write_text(HEADER, encoding="utf-8")
        result = api.save_contacts({
            "path": str(path),
            "rows": [
                {"dealership": "A Motors", "contact_name": "Sam", "email": "sam@example.com"},
                {"dealership": "B Motors", "contact_name": "Pat", "email": "pat@example.com"},
            ],
        })
        self.assertEqual(result["rows"], 2)
        self.assertEqual(result["dealerships"], ["A Motors", "B Motors"])
        self.assertEqual(len(result["contacts"]["rows"]), 2, "both rows must reach the file")

    def test_the_change_handler_does_not_re_render(self) -> None:
        script = (Path(__file__).resolve().parent.parent / "webapp" / "static" / "app.js").read_text(encoding="utf-8")
        start = script.index("select.addEventListener('change'")
        handler = script[start:script.index("return select;", start)]
        self.assertNotIn("renderContacts()", handler)
        self.assertIn("updateContactsNote()", handler)


if __name__ == "__main__":
    unittest.main()
