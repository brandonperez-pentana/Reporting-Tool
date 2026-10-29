"""Short forms in PDF filenames come from data/dealership_aliases.json, not code."""

from __future__ import annotations

import json
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import dealership_insights_local as core


class DealershipAliasTests(unittest.TestCase):
    def setUp(self) -> None:
        self.folder = Path(tempfile.mkdtemp(prefix="aliases-"))
        self.addCleanup(shutil.rmtree, self.folder, True)
        self.path = self.folder / "dealership_aliases.json"

    def _write(self, payload: object) -> None:
        self.path.write_text(json.dumps(payload), encoding="utf-8")
        core._ALIAS_CACHE.pop(str(self.path), None)

    def _name(self, file_name: str) -> str:
        with mock.patch.object(core, "DEALERSHIP_ALIASES_PATH", self.path):
            return core.extract_dealership_name_from_filename(file_name, "Fallback")

    def test_a_short_form_expands_to_the_full_name(self) -> None:
        self._write({"aliases": [{"name": "Demo Site 1 North", "matches": ["Demo Site 1 Nth", "DS1 N"]}]})
        self.assertEqual(self._name("{Demo Site 1 Nth} Monthly Sales Report - July 2026.pdf"), "Demo Site 1 North")
        self.assertEqual(self._name("ds1 (n) Monthly Service Report - July 2026.pdf"), "Demo Site 1 North")

    def test_only_whole_words_match(self) -> None:
        self._write({"aliases": [{"name": "Demo Site 1 North", "matches": ["Nth"]}]})
        self.assertEqual(self._name("Nthside Motors - Sales Report - July 2026.pdf"), "Nthside Motors")

    def test_the_expanded_name_is_not_expanded_again(self) -> None:
        self._write({"aliases": [{"name": "Demo Site 2 Eastside", "matches": ["Demo Site 2 East"]}]})
        self.assertEqual(self._name("Demo Site 2 Eastside - Sales Report - July 2026.pdf"), "Demo Site 2 Eastside")

    def test_no_file_means_names_pass_through(self) -> None:
        self.assertEqual(self._name("Demo Site 3 - Sales Report - July 2026.pdf"), "Demo Site 3")

    def test_an_unreadable_file_is_ignored(self) -> None:
        self.path.write_text("{not json", encoding="utf-8")
        core._ALIAS_CACHE.pop(str(self.path), None)
        with self.assertLogs(level="WARNING"):
            self.assertEqual(self._name("Demo Site 3 - Sales Report - July 2026.pdf"), "Demo Site 3")

    def test_malformed_entries_are_skipped(self) -> None:
        self._write({"aliases": ["text", {"name": "", "matches": ["x"]}, {"name": "Y", "matches": "x"},
                                 {"name": "Demo Site 4", "matches": ["DS4"]}]})
        self.assertEqual(self._name("DS4 - Sales Report - July 2026.pdf"), "Demo Site 4")

    def test_the_shipped_file_is_valid(self) -> None:
        shipped = core.DEALERSHIP_ALIASES_PATH
        if not shipped.is_file():
            self.skipTest("no aliases file in this copy")
        payload = json.loads(shipped.read_text(encoding="utf-8"))
        for entry in payload["aliases"]:
            self.assertTrue(str(entry["name"]).strip())
            self.assertIsInstance(entry["matches"], list)


if __name__ == "__main__":
    unittest.main()
