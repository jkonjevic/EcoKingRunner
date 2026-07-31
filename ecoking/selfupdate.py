"""Refresh the frozen desktop build's application code from GitHub.

The frozen build ships the app's Python/HTML/JS source as loose files next to
the exe (not baked into the PyInstaller archive -- see EcoKingWebRunner.spec's
``excludes``) specifically so this can overwrite them safely. Only code paths
are touched; local state (``.env``, ``stations.json``, the report template,
logs, reports) is never part of the sync set, so it survives every update.

A network hiccup or a private/renamed repo must never stop the app from
starting, so every failure here is caught and logged, not raised.
"""

from __future__ import annotations

import logging
import shutil
import tempfile
import urllib.request
from pathlib import Path
from zipfile import ZipFile

DEFAULT_OWNER = "jkonjevic"
DEFAULT_REPO = "EcoKingRunner"
DEFAULT_BRANCH = "main"
TIMEOUT_SECONDS = 20

#: Code the frozen build should stay in sync with. Anything else already on
#: disk next to the exe (.env, stations.json, the workbook, logs/, reports)
#: is left completely alone.
SYNCED_PATHS = (
    "ecoking",
    "web",
    "ecoking_daily.py",
    "ecoking_web_launcher.py",
    "requirements.txt",
)

#: A download missing any of these gets rejected before anything on disk is
#: touched. Without this check, a stale or partial branch snapshot could
#: overwrite the running app's own ecoking/selfupdate.py with a copy that
#: predates its existence, breaking every future update, not just this one.
REQUIRED_FILES = ("ecoking/selfupdate.py", "ecoking/webapp.py", "ecoking_daily.py")


def archive_url(owner: str, repo: str, branch: str) -> str:
    return f"https://github.com/{owner}/{repo}/archive/refs/heads/{branch}.zip"


def sync_from_github(
    root: Path,
    owner: str = DEFAULT_OWNER,
    repo: str = DEFAULT_REPO,
    branch: str = DEFAULT_BRANCH,
    synced_paths: tuple[str, ...] = SYNCED_PATHS,
) -> bool:
    """Best-effort refresh of ``root`` from the GitHub repo's default branch.

    Returns whether it actually updated anything; never raises.
    """
    url = archive_url(owner, repo, branch)
    try:
        with tempfile.TemporaryDirectory(prefix="ecoking-update-") as tmp:
            tmp_path = Path(tmp)
            archive_path = tmp_path / "source.zip"
            _download(url, archive_path)
            extract_dir = tmp_path / "extracted"
            with ZipFile(archive_path) as archive:
                archive.extractall(extract_dir)
            # GitHub's zip wraps everything in a single "<repo>-<branch>/" folder.
            source_root = next(extract_dir.iterdir())
            _verify_download_is_safe(source_root)
            apply_update(source_root, root, synced_paths)
        logging.info("Updated application code from %s/%s@%s.", owner, repo, branch)
        return True
    except Exception as exc:
        logging.warning("Could not check for updates (%s). Continuing with the current version.", exc)
        return False


def _download(url: str, destination: Path) -> None:
    request = urllib.request.Request(url, headers={"User-Agent": "EcoKingRunner-selfupdate"})
    with urllib.request.urlopen(request, timeout=TIMEOUT_SECONDS) as response:
        destination.write_bytes(response.read())


def _verify_download_is_safe(source_root: Path) -> None:
    missing = [path for path in REQUIRED_FILES if not (source_root / path).exists()]
    if missing:
        raise RuntimeError(f"Downloaded update is missing required file(s): {missing}")


def apply_update(source_root: Path, root: Path, synced_paths: tuple[str, ...] = SYNCED_PATHS) -> list[str]:
    """Copy only ``synced_paths`` from ``source_root`` onto ``root``.

    Downloads and extracts fully before touching ``root`` (in the caller), so
    a network failure never leaves a half-applied update; a failure partway
    through the copy itself is a narrower, already-rare risk left as-is for
    an internal tool.
    """
    updated: list[str] = []
    for relative in synced_paths:
        source = source_root / relative
        if not source.exists():
            continue
        destination = root / relative
        if source.is_dir():
            if destination.exists():
                shutil.rmtree(destination)
            shutil.copytree(source, destination)
        else:
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, destination)
        updated.append(relative)
    return updated
