"""Entry point for the web UI: ``python ecoking_web_launcher.py``, and the
frozen desktop build (see EcoKingWebRunner.spec / build_windows_web.bat).

In the frozen build, ``ecoking/``, ``ecoking_daily.py`` and ``web/`` ship as
loose files next to the exe instead of being baked into the PyInstaller
archive, so :mod:`ecoking.selfupdate` can refresh them from GitHub between
runs. That means this file must not import them at module load time -- Python
would either fail to find the loose copies or import a stale bundled one.
Every import of them below is deliberately deferred inside a function so
``app_root()`` lands on ``sys.path`` first.
"""

from __future__ import annotations

import sys
from pathlib import Path


def app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _ensure_app_root_on_path() -> None:
    root = str(app_root())
    if root not in sys.path:
        sys.path.insert(0, root)


def run_scraper_in_process() -> int:
    """Handle a ``--run-scraper`` re-launch.

    A frozen exe can't be told "run this other .py file" the way a real
    python.exe can, so the web UI launches this same exe with this flag and
    ecoking_daily.main() runs in-process instead of as a subprocess.
    """
    args = [arg for arg in sys.argv[1:] if arg != "--run-scraper"]
    sys.argv = ["ecoking_daily.py", *args]
    from ecoking_daily import main as scraper_main

    return scraper_main()


def main() -> None:
    _ensure_app_root_on_path()

    if "--run-scraper" in sys.argv:
        raise SystemExit(run_scraper_in_process())

    if getattr(sys, "frozen", False):
        from ecoking.selfupdate import sync_from_github

        sync_from_github(app_root())

    from ecoking.webapp import main as run_webapp

    run_webapp()


if __name__ == "__main__":
    main()
