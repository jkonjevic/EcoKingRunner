"""Check the station registry against the Excel template.

``python -m ecoking.check`` exits non-zero on errors so CI catches a broken
station list before it reaches a run. Warnings are printed but do not fail.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from ecoking import stations as registry

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_TEMPLATE = ROOT / "ECO KING BLANKO TABLICA.xlsx"


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate stations.json against the report template.")
    parser.add_argument("--stations", default=None, help="Station registry path.")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="Report template path.")
    parser.add_argument("--strict", action="store_true", help="Fail on warnings too.")
    args = parser.parse_args()

    template = Path(args.template)
    if not template.exists():
        print(f"error: template not found: {template}")
        return 1

    path = registry.resolve_stations_path(args.stations, ROOT)
    if not path.exists():
        print(f"error: station registry not found: {path}")
        return 1

    stations = registry.load_stations(path)
    rows = registry.load_excel_rows(template)
    issues = registry.validate(stations, rows) + registry.ambiguous_device_labels(stations)

    errors = [issue for issue in issues if issue.severity == "error"]
    warnings = [issue for issue in issues if issue.severity == "warning"]

    print(f"{path.name}: {len(stations)} stations, {template.name}: {len(rows)} rows")
    for issue in errors + warnings:
        print(f"  [{issue.severity}] {issue.station}: {issue.message}")

    if errors:
        print(f"\n{len(errors)} error(s).")
        return 1
    if warnings and args.strict:
        print(f"\n{len(warnings)} warning(s) with --strict.")
        return 1
    print(f"\nOK ({len(warnings)} warning(s)).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
