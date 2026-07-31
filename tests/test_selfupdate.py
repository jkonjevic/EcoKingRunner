import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch
from urllib.error import URLError

from ecoking import selfupdate


class ApplyUpdateTests(unittest.TestCase):
    """apply_update() only touches the paths it's told to, nothing else."""

    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.source = Path(self.tmp.name) / "downloaded"
        self.root = Path(self.tmp.name) / "installed"

        (self.source / "ecoking").mkdir(parents=True)
        (self.source / "ecoking" / "webapp.py").write_text("# new version", encoding="utf-8")
        (self.source / "ecoking_daily.py").write_text("# new scraper", encoding="utf-8")
        (self.source / "requirements.txt").write_text("openpyxl>=3.1.5\n", encoding="utf-8")

        self.root.mkdir(parents=True)
        (self.root / "ecoking").mkdir()
        (self.root / "ecoking" / "webapp.py").write_text("# old version", encoding="utf-8")
        (self.root / "ecoking_daily.py").write_text("# old scraper", encoding="utf-8")

    def test_a_synced_file_is_overwritten(self) -> None:
        selfupdate.apply_update(self.source, self.root, synced_paths=("ecoking_daily.py",))
        self.assertEqual((self.root / "ecoking_daily.py").read_text(encoding="utf-8"), "# new scraper")

    def test_a_synced_directory_is_replaced_wholesale(self) -> None:
        (self.root / "ecoking" / "stale_module.py").write_text("# should disappear", encoding="utf-8")
        selfupdate.apply_update(self.source, self.root, synced_paths=("ecoking",))
        self.assertFalse((self.root / "ecoking" / "stale_module.py").exists())
        self.assertEqual((self.root / "ecoking" / "webapp.py").read_text(encoding="utf-8"), "# new version")

    def test_local_state_outside_the_synced_paths_is_untouched(self) -> None:
        (self.root / ".env").write_text("password=real-secret\n", encoding="utf-8")
        (self.root / "stations.json").write_text('{"stations": []}', encoding="utf-8")
        selfupdate.apply_update(self.source, self.root, synced_paths=("ecoking", "ecoking_daily.py"))
        self.assertEqual((self.root / ".env").read_text(encoding="utf-8"), "password=real-secret\n")
        self.assertEqual((self.root / "stations.json").read_text(encoding="utf-8"), '{"stations": []}')

    def test_a_path_missing_from_the_download_is_left_alone(self) -> None:
        selfupdate.apply_update(self.source, self.root, synced_paths=("does_not_exist.py",))
        self.assertEqual((self.root / "ecoking_daily.py").read_text(encoding="utf-8"), "# old scraper")

    def test_returns_only_the_paths_it_actually_updated(self) -> None:
        updated = selfupdate.apply_update(
            self.source, self.root, synced_paths=("ecoking_daily.py", "missing.py")
        )
        self.assertEqual(updated, ["ecoking_daily.py"])

    def test_a_new_file_not_previously_installed_is_created(self) -> None:
        selfupdate.apply_update(self.source, self.root, synced_paths=("requirements.txt",))
        self.assertEqual((self.root / "requirements.txt").read_text(encoding="utf-8"), "openpyxl>=3.1.5\n")


class SyncFromGithubTests(unittest.TestCase):
    """No live network here -- _download is mocked so this stays hermetic."""

    def test_a_download_failure_returns_false_instead_of_raising(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, patch.object(
            selfupdate, "_download", side_effect=URLError("no route to host")
        ):
            result = selfupdate.sync_from_github(Path(tmp))
        self.assertFalse(result)

    def test_a_successful_download_applies_the_update(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "installed"
            root.mkdir()

            with patch.object(selfupdate, "_download", side_effect=_fake_complete_download):
                result = selfupdate.sync_from_github(root, synced_paths=("ecoking_daily.py",))

        self.assertTrue(result)

    def test_a_download_missing_a_required_file_is_rejected_untouched(self) -> None:
        """Regression test: this exact scenario deleted a live install's own
        ecoking/selfupdate.py by syncing in a snapshot from before that file
        existed, bricking every update after it -- not just this one."""
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "installed"
            (root / "ecoking").mkdir(parents=True)
            (root / "ecoking" / "selfupdate.py").write_text("# the running version", encoding="utf-8")

            def fake_download_missing_selfupdate(url: str, destination: Path) -> None:
                with zipfile.ZipFile(destination, "w") as archive:
                    archive.writestr("EcoKingRunner-main/ecoking/webapp.py", "# newer webapp")
                    archive.writestr("EcoKingRunner-main/ecoking_daily.py", "# newer scraper")
                    # No ecoking/selfupdate.py in this snapshot.

            with patch.object(selfupdate, "_download", side_effect=fake_download_missing_selfupdate):
                result = selfupdate.sync_from_github(root)

            self.assertFalse(result)
            self.assertEqual(
                (root / "ecoking" / "selfupdate.py").read_text(encoding="utf-8"), "# the running version"
            )

    def test_archive_url_targets_the_requested_branch(self) -> None:
        url = selfupdate.archive_url("jkonjevic", "EcoKingRunner", "main")
        self.assertEqual(url, "https://github.com/jkonjevic/EcoKingRunner/archive/refs/heads/main.zip")


def _fake_complete_download(url: str, destination: Path) -> None:
    with zipfile.ZipFile(destination, "w") as archive:
        archive.writestr("EcoKingRunner-main/ecoking/selfupdate.py", "# self")
        archive.writestr("EcoKingRunner-main/ecoking/webapp.py", "# webapp")
        archive.writestr("EcoKingRunner-main/ecoking_daily.py", "# from github")


if __name__ == "__main__":
    unittest.main()
