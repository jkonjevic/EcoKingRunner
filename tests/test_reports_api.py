import os
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from ecoking import webapp


class HideReportTests(unittest.TestCase):
    """The Izvještaji list is a directory listing, so hiding is remembered."""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp, True)
        # ECOKING_MODE=cloud keeps reports_dir() inside DATA_DIR instead of
        # reaching for the real Desktop.
        patch = mock.patch.dict(
            os.environ, {"DATA_DIR": str(self.tmp), "ECOKING_MODE": "cloud"}
        )
        patch.start()
        self.addCleanup(patch.stop)

    def make_report(self, date: str) -> Path:
        path = webapp.report_path(date)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"report")
        return path

    def dates(self) -> list[str]:
        return [item["date"] for item in webapp.list_reports()]

    def test_hidden_row_leaves_the_list_but_the_file_stays(self):
        kept = self.make_report("2026-07-30")
        hidden = self.make_report("2026-07-31")
        webapp.hide_report("2026-07-31")
        self.assertEqual(self.dates(), ["2026-07-30"])
        self.assertTrue(hidden.exists(), "hiding must never delete the workbook")
        self.assertTrue(kept.exists())

    def test_hiding_survives_a_restart(self):
        self.make_report("2026-07-31")
        webapp.hide_report("2026-07-31")
        self.assertEqual(webapp.load_hidden_reports(), {"2026-07-31"})
        self.assertEqual(self.dates(), [])

    def test_restore_brings_every_row_back(self):
        self.make_report("2026-07-30")
        self.make_report("2026-07-31")
        webapp.hide_report("2026-07-30")
        webapp.hide_report("2026-07-31")
        self.assertEqual(webapp.reports_payload()["hiddenCount"], 2)
        webapp.unhide_reports()
        self.assertEqual(self.dates(), ["2026-07-31", "2026-07-30"])

    def test_regenerating_a_hidden_report_shows_it_again(self):
        self.make_report("2026-07-31")
        webapp.hide_report("2026-07-31")
        webapp.unhide_report("2026-07-31")
        self.assertEqual(self.dates(), ["2026-07-31"])

    def test_dates_whose_file_is_gone_stop_being_tracked(self):
        self.make_report("2026-07-30")
        self.make_report("2026-07-31")
        webapp.hide_report("2026-07-30")
        webapp.report_path("2026-07-30").unlink()
        webapp.hide_report("2026-07-31")
        self.assertEqual(webapp.load_hidden_reports(), {"2026-07-31"})

    def test_missing_report_is_rejected(self):
        with self.assertRaises(ValueError):
            webapp.hide_report("2026-07-31")

    def test_corrupt_state_file_does_not_break_the_list(self):
        self.make_report("2026-07-31")
        webapp.hidden_reports_path().write_text("{not json", encoding="utf-8")
        self.assertEqual(self.dates(), ["2026-07-31"])


if __name__ == "__main__":
    unittest.main()
