import os
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from ecoking import webapp


class DeleteReportTests(unittest.TestCase):
    """Reports are removed from disk, so the target must be pinned down."""

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

    def test_deletes_only_the_named_report(self):
        keep = self.make_report("2026-07-30")
        drop = self.make_report("2026-07-31")
        webapp.delete_report("2026-07-31")
        self.assertFalse(drop.exists())
        self.assertTrue(keep.exists())

    def test_missing_report_is_rejected(self):
        with self.assertRaises(ValueError):
            webapp.delete_report("2026-07-31")

    def test_locked_file_explains_itself(self):
        # Excel holds the workbook open, so unlink fails on Windows.
        path = self.make_report("2026-07-31")
        with mock.patch.object(Path, "unlink", side_effect=OSError("in use")):
            with self.assertRaises(RuntimeError) as caught:
                webapp.delete_report("2026-07-31")
        self.assertIn("Excel", str(caught.exception))
        self.assertTrue(path.exists())

    def test_listing_reflects_the_deletion(self):
        self.make_report("2026-07-30")
        self.make_report("2026-07-31")
        webapp.delete_report("2026-07-30")
        self.assertEqual([item["date"] for item in webapp.list_reports()], ["2026-07-31"])


if __name__ == "__main__":
    unittest.main()
