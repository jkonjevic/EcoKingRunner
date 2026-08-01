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

import json
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
    "telemetry_list.py",
    "telemetry_mapping.json",
    "requirements.txt",
)

#: A download missing any of these gets rejected before anything on disk is
#: touched. Without this check, a stale or partial branch snapshot could
#: overwrite the running app's own ecoking/selfupdate.py with a copy that
#: predates its existence, breaking every future update, not just this one.
REQUIRED_FILES = ("ecoking/selfupdate.py", "ecoking/webapp.py", "ecoking_daily.py")


def archive_url(owner: str, repo: str, branch: str) -> str:
    return f"https://github.com/{owner}/{repo}/archive/refs/heads/{branch}.zip"


def commit_api_url(owner: str, repo: str, branch: str) -> str:
    return f"https://api.github.com/repos/{owner}/{repo}/commits/{branch}"


def sync_from_github(
    root: Path,
    owner: str = DEFAULT_OWNER,
    repo: str = DEFAULT_REPO,
    branch: str = DEFAULT_BRANCH,
    synced_paths: tuple[str, ...] = SYNCED_PATHS,
) -> bool:
    """Best-effort refresh of ``root`` from the GitHub repo's default branch.

    Returns whether it actually updated anything; never raises. Prints to the
    console (in addition to logging) so a non-technical user watching the
    window can see which commit it pulled without digging into log level
    configuration.
    """
    print(f"Checking for updates ({owner}/{repo}@{branch})...", flush=True)
    try:
        sha, message, date = _fetch_latest_commit_info(owner, repo, branch)
        print(f"Latest commit on GitHub: {sha[:7]} ({date}) {message}", flush=True)
    except Exception as exc:
        sha = None
        print(f"Could not fetch latest commit info ({exc}); trying the update anyway.", flush=True)

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
        commit_note = f" (commit {sha[:7]})" if sha else ""
        print(f"Updated application code from {owner}/{repo}@{branch}{commit_note}.", flush=True)
        logging.info("Updated application code from %s/%s@%s.", owner, repo, branch)
        return True
    except Exception as exc:
        print(f"Could not check for updates ({exc}). Continuing with the current version.", flush=True)
        logging.warning("Could not check for updates (%s). Continuing with the current version.", exc)
        return False


def _download(url: str, destination: Path) -> None:
    request = urllib.request.Request(url, headers={"User-Agent": "EcoKingRunner-selfupdate"})
    with urllib.request.urlopen(request, timeout=TIMEOUT_SECONDS) as response:
        destination.write_bytes(response.read())


def _fetch_latest_commit_info(owner: str, repo: str, branch: str) -> tuple[str, str, str]:
    """Return (short description of the latest commit on ``branch``).

    Separate request from the archive download itself (GitHub's zip has no
    commit metadata in it) -- purely informational, so any failure here must
    not block the actual update.
    """
    url = commit_api_url(owner, repo, branch)
    request = urllib.request.Request(
        url,
        headers={"User-Agent": "EcoKingRunner-selfupdate", "Accept": "application/vnd.github+json"},
    )
    with urllib.request.urlopen(request, timeout=TIMEOUT_SECONDS) as response:
        data = json.loads(response.read())
    sha = data["sha"]
    message = data["commit"]["message"].splitlines()[0]
    date = data["commit"]["committer"]["date"]
    return sha, message, date


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
