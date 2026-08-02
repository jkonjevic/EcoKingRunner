"""Web UI for the EcoKing daily report.

The same server runs on a laptop and on a hosted container. The only
differences are where the report lands and whether it can be opened in a local
Excel, both decided by :func:`is_cloud`.
"""

from __future__ import annotations

import json
import mimetypes
import os
import re
import secrets
import shutil
import subprocess
import sys
import threading
from dataclasses import dataclass, field, replace
from datetime import datetime, timedelta
from http import HTTPStatus
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Callable
from urllib.parse import parse_qs, urlparse

from ecoking import logtext
from ecoking import stations as registry
from ecoking.stations import ExcelRow, Station


def _app_root() -> Path:
    # A frozen build's __file__ points inside the bundle, not next to the
    # exe where .env / stations.json / the template actually live.
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


ROOT = _app_root()
WEB_DIR = ROOT / "web"
TEMPLATE_PATH = ROOT / "ECO KING BLANKO TABLICA.xlsx"
SCRAPER = ROOT / "ecoking_daily.py"

ISO_DATE = "%Y-%m-%d"
_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_STATION_LINE_RE = re.compile(r"\[(\d+)/(\d+)\]")


# --------------------------------------------------------------------------- #
# Environment
# --------------------------------------------------------------------------- #


def is_cloud() -> bool:
    """True when there is no user desktop to write to."""
    if os.getenv("ECOKING_MODE"):
        return os.getenv("ECOKING_MODE", "").lower() == "cloud"
    return bool(os.getenv("SPACE_ID") or os.getenv("RENDER") or os.getenv("KOYEB_APP_NAME"))


def data_dir() -> Path:
    """Writable directory for reports, logs and the device cache."""
    configured = os.getenv("DATA_DIR")
    base = Path(configured) if configured else (ROOT / "data" if is_cloud() else ROOT)
    base.mkdir(parents=True, exist_ok=True)
    return base


def desktop_directory() -> Path:
    candidates = [Path.home() / "Desktop"]
    if os.environ.get("OneDrive"):
        candidates.insert(0, Path(os.environ["OneDrive"]) / "Desktop")
    return next((path for path in candidates if path.is_dir()), ROOT)


def reports_dir() -> Path:
    directory = data_dir() / "reports" if is_cloud() else desktop_directory()
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def log_dir() -> Path:
    # Locally this stays ROOT/logs so both UIs write to the same folder.
    directory = (data_dir() / "logs") if is_cloud() else (ROOT / "logs")
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def stations_path() -> Path:
    override = os.getenv("STATIONS_PATH")
    if override:
        return Path(override)
    if is_cloud():
        # A hosted container has an ephemeral filesystem, so edits are kept in
        # the data volume and fall back to the file shipped in the image.
        editable = data_dir() / registry.DEFAULT_STATIONS_FILE
        if not editable.exists() and (ROOT / registry.DEFAULT_STATIONS_FILE).exists():
            shutil.copy2(ROOT / registry.DEFAULT_STATIONS_FILE, editable)
        return editable
    return registry.resolve_stations_path(None, ROOT)


def report_path(selected_date: str) -> Path:
    return reports_dir() / f"EcoKing_Report_{selected_date}.xlsx"


def yesterday() -> str:
    return (datetime.now() - timedelta(days=1)).strftime(ISO_DATE)


# --------------------------------------------------------------------------- #
# Run state
# --------------------------------------------------------------------------- #


#: One report per day, so a run over several days is a queue of the runs that
#: already worked -- the scraper still sees exactly one --selected-date.
MAX_BATCH_DAYS = 14

#: Written between days so the log stays one readable stream and the UI can
#: still tell which day a line belongs to. Single-day runs emit no separator,
#: which keeps their log byte-identical to what it has always been.
DAY_SEPARATOR = "───────── DAN {index}/{total} · {date} ─────────"


@dataclass
class DayRun:
    """One date in a batch: its own progress, outcome and list of problems."""

    date: str
    status: str = "pending"  # pending | running | ok | failed | stopped | skipped
    return_code: int | None = None
    started_at: str | None = None
    finished_at: str | None = None
    done: int = 0
    total: int = 0
    current: str = ""
    errors: int = 0
    warnings: int = 0
    failures: list[logtext.Failure] = field(default_factory=list)

    def to_json(self) -> dict[str, Any]:
        return {
            "date": self.date,
            "status": self.status,
            "returnCode": self.return_code,
            "startedAt": self.started_at,
            "finishedAt": self.finished_at,
            "done": self.done,
            "total": self.total,
            "current": self.current,
            "errors": self.errors,
            "warnings": self.warnings,
            "failures": [failure.to_json() for failure in self.failures],
            "reportReady": report_path(self.date).exists(),
        }


@dataclass
class RunState:
    process: subprocess.Popen[str] | None = None
    running: bool = False
    return_code: int | None = None
    log_path: Path | None = None
    lines: list[str] = field(default_factory=list)
    started_at: str | None = None
    days: list[DayRun] = field(default_factory=list)
    index: int = 0
    #: Set by "Zaustavi": the running day is killed and the rest are skipped.
    cancelled: bool = False
    #: Set by "Preskoči dan": only the running day is killed.
    stop_current: bool = False

    @property
    def active(self) -> DayRun | None:
        """The day being run, or the last one touched once the batch is over."""
        if not self.days:
            return None
        return self.days[min(self.index, len(self.days) - 1)]

    # The single-day fields the UI has always polled. They now describe the
    # active day, so a batch of one behaves exactly as it did before.
    @property
    def selected_date(self) -> str | None:
        day = self.active
        return day.date if day else None

    @property
    def done(self) -> int:
        day = self.active
        return day.done if day else 0

    @property
    def total(self) -> int:
        day = self.active
        return day.total if day else 0

    @property
    def current(self) -> str:
        day = self.active
        return day.current if day else ""


STATE = RunState()
STATE_LOCK = threading.Lock()


def append_log(line: str) -> None:
    translated = logtext.translate(line)
    if not translated:
        return
    # Severity and failures are read off the scraper's own English wording,
    # which is stable, rather than off the translation shown in the console.
    severity = logtext.classify(line)
    failure = logtext.parse_failure(line)
    with STATE_LOCK:
        STATE.lines.append(translated)
        day = STATE.active if STATE.running else None
        if day:
            if severity == "error":
                day.errors += 1
            elif severity == "warning":
                day.warnings += 1
            if failure:
                day.failures.append(failure)
        progress = _STATION_LINE_RE.search(translated)
        if progress and day:
            day.done = int(progress.group(1))
            day.total = int(progress.group(2))
            # Two passes report progress: EcoKing stations ("Stanica=") and the
            # telemetry levels ("Lokacija="). The counter restarts for the
            # second one, which is what the log lines say too.
            marker = next((name for name in ("Stanica=", "Lokacija=") if name in translated), None)
            if marker:
                label = translated.split(marker, 1)[-1].split(", Excel red=", 1)[0]
                day.current = label.strip()
        log_path = STATE.log_path
    if log_path:
        with log_path.open("a", encoding="utf-8") as handle:
            handle.write(translated + "\n")


def start_batch(dates: list[str], config: dict[str, Any]) -> str:
    """Reset the log and run one scraper process per date, in order."""
    with STATE_LOCK:
        if STATE.running:
            raise RuntimeError("Obračun je već u toku.")
        STATE.lines.clear()
        STATE.return_code = None
        STATE.log_path = log_dir() / f"ecoking-{datetime.now().strftime('%Y%m%d-%H%M%S')}.log"
        STATE.started_at = datetime.now().strftime("%d.%m.%Y. %H:%M:%S")
        STATE.running = True
        STATE.cancelled = False
        STATE.stop_current = False
        STATE.days = [DayRun(date=date) for date in dates]
        STATE.index = 0
        log_path = STATE.log_path

    threading.Thread(target=_pump_batch, args=(config,), daemon=True).start()
    return str(log_path.name)


def _pump_batch(config: dict[str, Any]) -> None:
    """Walk the queue. A day that fails is recorded, not allowed to end it."""
    try:
        with STATE_LOCK:
            total = len(STATE.days)
        for index in range(total):
            with STATE_LOCK:
                if STATE.cancelled:
                    for pending in STATE.days[index:]:
                        pending.status = "skipped"
                    break
                STATE.index = index
                day = STATE.days[index]
                day.status = "running"
                day.started_at = datetime.now().strftime("%d.%m.%Y. %H:%M:%S")
            if total > 1:
                append_log("")
                append_log(DAY_SEPARATOR.format(index=index + 1, total=total, date=day.date))
            _announce_day(day.date, bool(config.get("onlyTelemetry")))

            cmd, environment, _ = build_run_command({**config, "selectedDate": day.date})
            return_code = _pump_process(cmd, environment)

            with STATE_LOCK:
                stopped = STATE.stop_current or STATE.cancelled
                STATE.stop_current = False
                day.return_code = return_code
                day.finished_at = datetime.now().strftime("%d.%m.%Y. %H:%M:%S")
                day.status = "ok" if return_code == 0 else "stopped" if stopped else "failed"
            append_log(_day_verdict(day))
            if return_code == 0:
                unhide_report(day.date)
    finally:
        with STATE_LOCK:
            STATE.running = False
            STATE.process = None
            failed = [day for day in STATE.days if day.status not in {"ok"}]
            STATE.return_code = 0 if not failed else 1
            summary = _batch_summary(STATE.days)
        if summary:
            append_log(summary)


def _announce_day(selected_date: str, only_telemetry: bool) -> None:
    append_log(
        f"Očitavanje nivoa rezervoara u 17h za datum {selected_date}."
        if only_telemetry
        else f"Generisanje izvještaja za datum {selected_date}."
    )
    append_log(f"Izlazni fajl: {report_path(selected_date).name}")


def _day_verdict(day: DayRun) -> str:
    if day.status == "ok":
        return f"Obračun za {day.date} je završen uspješno."
    if day.status == "stopped":
        return f"Obračun za {day.date} je zaustavljen."
    return f"Obračun za {day.date} je završen sa greškom. Kod izlaza: {day.return_code}"


def _batch_summary(days: list[DayRun]) -> str:
    """One closing line for a multi-day run; single days already said it all."""
    if len(days) < 2:
        return ""
    ok = sum(1 for day in days if day.status == "ok")
    failed = [day.date for day in days if day.status == "failed"]
    skipped = [day.date for day in days if day.status in {"skipped", "stopped"}]
    parts = [f"ZBIRNO: {ok} od {len(days)} dana uspješno."]
    if failed:
        parts.append(f"Sa greškom: {', '.join(failed)}.")
    if skipped:
        parts.append(f"Nije pokrenuto: {', '.join(skipped)}.")
    return " ".join(parts)


def _pump_process(cmd: list[str], environment: dict[str, str]) -> int:
    """Spawn one scraper process and stream its output into the state."""
    return_code = 1
    try:
        process = subprocess.Popen(
            cmd,
            cwd=str(ROOT),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            env=environment,
        )
        with STATE_LOCK:
            STATE.process = process
        assert process.stdout is not None
        for raw_line in process.stdout:
            append_log(raw_line)
        return_code = process.wait()
    except Exception as exc:  # pragma: no cover - depends on the OS
        append_log(f"GREŠKA: Pokretanje procesa nije uspjelo: {exc}")
    with STATE_LOCK:
        STATE.process = None
    return return_code


def selected_dates(payload: dict[str, Any]) -> list[str]:
    """The dates a run covers: deduplicated, ascending, validated as a whole.

    A batch is checked before any of it starts, because discovering a bad date
    on day three of three -- twenty minutes in -- helps nobody.
    """
    raw = payload.get("selectedDates")
    if isinstance(raw, list) and raw:
        candidates = [str(item).strip() for item in raw if str(item).strip()]
    else:
        candidates = [str(payload.get("selectedDate") or yesterday()).strip()]

    ordered = sorted(dict.fromkeys(candidates))
    if not ordered:
        raise ValueError("Izaberi bar jedan datum.")
    if len(ordered) > MAX_BATCH_DAYS:
        raise ValueError(f"Najviše {MAX_BATCH_DAYS} dana odjednom; izabrano je {len(ordered)}.")

    today = datetime.now().date()
    problems = []
    for date in ordered:
        if not _DATE_RE.match(date):
            problems.append(f"{date}: datum mora biti u formatu YYYY-MM-DD.")
            continue
        try:
            parsed = datetime.strptime(date, ISO_DATE).date()
        except ValueError:
            problems.append(f"{date}: datum ne postoji.")
            continue
        if parsed > today:
            problems.append(f"{date}: datum ne može biti u budućnosti.")
    if problems:
        raise ValueError("\n".join(problems))
    return ordered


def build_run_command(config: dict[str, Any]) -> tuple[list[str], dict[str, str], str]:
    selected_date = str(config.get("selectedDate") or yesterday())
    output = report_path(selected_date)
    # A frozen exe is not a general-purpose interpreter -- it can't be told
    # to "run ecoking_daily.py". It re-launches itself with a flag it
    # recognises instead (see ecoking_web_launcher.py's --run-scraper).
    cmd = [sys.executable, "--run-scraper"] if getattr(sys, "frozen", False) else [sys.executable, str(SCRAPER)]
    cmd += [
        "--output",
        str(output),
        "--template",
        str(TEMPLATE_PATH),
        "--stations",
        str(stations_path()),
        "--selected-date",
        selected_date,
        "--workers",
        str(_clamp(config.get("workers"), 1, 1, 8)),
        "--slow-mo-ms",
        str(_clamp(config.get("slowMo"), 0, 0, 5000)),
    ]
    cmd.append("--headed" if bool(config.get("browserVisible")) and not is_cloud() else "--headless")
    if bool(config.get("verbose", True)):
        cmd.append("--verbose")
    limit = str(config.get("limit") or "").strip()
    if limit:
        cmd.extend(["--limit", str(_clamp(limit, 1, 1, 500))])

    # "Samo nivoi" re-runs the second pass over a report that already exists;
    # otherwise the checkbox decides whether it follows the EcoKing scrape.
    if bool(config.get("onlyTelemetry")):
        cmd.append("--only-telemetry")
    elif not bool(config.get("withTelemetry", True)):
        cmd.append("--skip-telemetry")

    environment = os.environ.copy()
    environment["NO_COLOR"] = "1"
    environment["PYTHONIOENCODING"] = "utf-8"
    environment["PYTHONUNBUFFERED"] = "1"
    environment["CHART_WAIT_MS"] = str(_clamp(config.get("chartWait"), 5000, 1000, 60000))
    environment["SEARCH_RESULTS_WAIT_MS"] = str(_clamp(config.get("searchWait"), 2000, 500, 30000))
    environment["COPY_REPORT_TO_DESKTOP"] = "0" if is_cloud() else "1"
    # The telemetry pass has its own browser settings -- the ones above and
    # --headed/--headless are aimed at the EcoKing scrape.
    environment["TELEMETRY_WAIT_MS"] = str(_clamp(config.get("telemetryWait"), 10000, 2000, 120000))
    telemetry_visible = bool(config.get("telemetryVisible")) and not is_cloud()
    environment["TELEMETRY_HEADLESS"] = "0" if telemetry_visible else "1"
    return cmd, environment, selected_date


def _clamp(value: Any, default: int, low: int, high: int) -> int:
    try:
        number = int(str(value).strip())
    except (TypeError, ValueError):
        return default
    return max(low, min(high, number))


# --------------------------------------------------------------------------- #
# Station registry helpers
# --------------------------------------------------------------------------- #


def load_template_rows() -> list[ExcelRow]:
    if not TEMPLATE_PATH.exists():
        return []
    return registry.load_excel_rows(TEMPLATE_PATH)


def stations_payload() -> dict[str, Any]:
    rows = load_template_rows()
    path = stations_path()
    stations = registry.load_stations(path) if path.exists() else []
    issues = registry.validate(stations, rows) + registry.ambiguous_device_labels(stations)
    return {
        "path": str(path),
        "template": TEMPLATE_PATH.name,
        "stations": [station.to_json() for station in stations],
        "rows": [
            {"row": row.row, "lokacija": row.lokacija, "vodomjer": row.vodomjer, "label": row.label}
            for row in rows
        ],
        "issues": [issue.to_json() for issue in issues],
    }


def save_stations_payload(payload: dict[str, Any]) -> dict[str, Any]:
    incoming = payload.get("stations")
    if not isinstance(incoming, list):
        raise ValueError("Očekivana je lista stanica.")
    # Pasting a dropdown entry straight off the site is the obvious thing to do
    # when a name is ambiguous, so accept it and store the short form.
    stations = [
        replace(Station.from_json(item), uredjaj=registry.device_label(str(item.get("uredjaj") or "").strip()))
        for item in incoming
        if isinstance(item, dict)
    ]
    rows = load_template_rows()
    issues = registry.validate(stations, rows)
    blocking = [issue for issue in issues if issue.severity == "error"]
    if blocking:
        raise ValueError("\n".join(f"{issue.station}: {issue.message}" for issue in blocking))
    registry.save_stations(stations_path(), stations)
    return stations_payload()


def repair_stations() -> dict[str, Any]:
    """Re-point every station at the closest free template row."""
    rows = load_template_rows()
    path = stations_path()
    stations = registry.load_stations(path) if path.exists() else []
    registry.save_stations(path, registry.reconcile_with_template(stations, rows))
    return stations_payload()


#: Dates the user took out of the Izvještaji list. The list itself is just a
#: directory listing, so hiding a row has to be remembered somewhere; this
#: keeps it out of the reports folder so it never looks like a report.
HIDDEN_REPORTS_FILE = "hidden_reports.json"


def hidden_reports_path() -> Path:
    return data_dir() / HIDDEN_REPORTS_FILE


def load_hidden_reports() -> set[str]:
    path = hidden_reports_path()
    if not path.exists():
        return set()
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        # A corrupt file must not take the reports list down with it.
        return set()
    dates = raw.get("hidden") if isinstance(raw, dict) else raw
    return {str(date) for date in dates or [] if _DATE_RE.match(str(date))}


def save_hidden_reports(dates: set[str]) -> None:
    path = hidden_reports_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps({"hidden": sorted(dates, reverse=True)}, indent=2) + "\n", encoding="utf-8"
    )


def all_report_dates() -> set[str]:
    return {
        item.stem.replace("EcoKing_Report_", "")
        for item in reports_dir().glob("EcoKing_Report_*.xlsx")
    }


def list_reports() -> list[dict[str, Any]]:
    hidden = load_hidden_reports()
    reports = []
    for item in sorted(reports_dir().glob("EcoKing_Report_*.xlsx"), reverse=True):
        date = item.stem.replace("EcoKing_Report_", "")
        if date in hidden:
            continue
        try:
            stat = item.stat()
        except OSError:
            continue
        reports.append(
            {
                "date": date,
                "name": item.name,
                "size": stat.st_size,
                "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%d.%m.%Y. %H:%M"),
            }
        )
    return reports[:30]


def reports_payload() -> dict[str, Any]:
    """The list plus how many rows are hidden, so the UI can offer them back."""
    return {"reports": list_reports(), "hiddenCount": len(load_hidden_reports() & all_report_dates())}


def hide_report(selected_date: str) -> None:
    """Take one row out of the Izvještaji list. The .xlsx stays on disk."""
    if selected_date not in all_report_dates():
        raise ValueError("Izvještaj za taj datum ne postoji.")
    hidden = load_hidden_reports()
    hidden.add(selected_date)
    # Dates whose file is gone would pile up forever otherwise.
    save_hidden_reports(hidden & all_report_dates())
    append_log(f"Izvještaj {report_path(selected_date).name} je uklonjen iz liste (fajl je ostao na disku).")


def unhide_reports() -> None:
    save_hidden_reports(set())


def unhide_report(selected_date: str) -> None:
    """A freshly generated report must never stay hidden."""
    hidden = load_hidden_reports()
    if selected_date in hidden:
        save_hidden_reports((hidden - {selected_date}) & all_report_dates())


def _spawn_first_available(commands: list[list[str]]) -> bool:
    """Try each desktop opener in turn; the first one installed wins."""
    for command in commands:
        if not shutil.which(command[0]):
            continue
        try:
            subprocess.Popen(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
        except OSError:
            continue
    return False


def open_in_excel(path: Path) -> bool:
    """Open a finished report in the desktop spreadsheet app. Local runs only."""
    if is_cloud() or not path.exists():
        return False
    if sys.platform.startswith("win"):
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
            return True
        except OSError as exc:
            append_log(f"UPOZORENJE: Nije moguće otvoriti Excel: {exc}")
            return False

    commands = [["open", str(path)]] if sys.platform == "darwin" else [
        ["xdg-open", str(path)],
        ["gio", "open", str(path)],
        ["libreoffice", "--calc", str(path)],
    ]
    if _spawn_first_available(commands):
        return True
    append_log(f"UPOZORENJE: Excel nije otvoren automatski. Fajl je na: {path}")
    return False


def open_reports_folder(selected_date: str | None = None) -> bool:
    """Show the reports folder in the file manager, one report highlighted.

    A hosted container has no file manager, and the browser download is the
    only way to a report there, so this is a local-only convenience.
    """
    if is_cloud():
        return False
    directory = reports_dir()
    target = report_path(selected_date) if selected_date else None
    if target and not target.exists():
        target = None

    if sys.platform.startswith("win"):
        try:
            if target:
                # Passing the argument as one string keeps Explorer's
                # "/select,<path>" syntax intact -- a list would quote the
                # comma-joined pair apart and open the wrong thing.
                subprocess.Popen(f'explorer /select,"{target}"')
                return True
            os.startfile(str(directory))  # type: ignore[attr-defined]
            return True
        except OSError as exc:
            append_log(f"UPOZORENJE: Nije moguće otvoriti folder: {exc}")
            return False

    if sys.platform == "darwin":
        commands = [["open", "-R", str(target)]] if target else [["open", str(directory)]]
    else:
        commands = [["xdg-open", str(directory)], ["gio", "open", str(directory)]]
    if _spawn_first_available(commands):
        return True
    append_log(f"UPOZORENJE: Folder nije otvoren automatski. Putanja je: {directory}")
    return False


# --------------------------------------------------------------------------- #
# Sessions
# --------------------------------------------------------------------------- #

SESSIONS: set[str] = set()
SESSION_COOKIE = "ecoking_session"


def app_password() -> str:
    return os.getenv("APP_PASSWORD", "").strip()


def authenticate(password: str) -> str | None:
    expected = app_password()
    if not expected or not secrets.compare_digest(password, expected):
        return None
    token = secrets.token_urlsafe(32)
    SESSIONS.add(token)
    return token


# --------------------------------------------------------------------------- #
# HTTP
# --------------------------------------------------------------------------- #


class Handler(BaseHTTPRequestHandler):
    server_version = "EcoKing"

    def log_message(self, format: str, *args: object) -> None:
        return

    # -- helpers ---------------------------------------------------------- #

    def send_json(self, payload: object, status: int = 200, cookie: str | None = None) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        if cookie:
            self.send_header("Set-Cookie", f"{SESSION_COOKIE}={cookie}; Path=/; HttpOnly; SameSite=Lax")
        self.end_headers()
        self.wfile.write(body)

    def send_file(self, path: Path, download_name: str | None = None) -> None:
        if not path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        body = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        if download_name:
            self.send_header("Content-Disposition", f'attachment; filename="{download_name}"')
        else:
            self.send_header("Cache-Control", "no-cache")
        self.end_headers()
        self.wfile.write(body)

    def read_json(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length", "0") or 0)
        if not length:
            return {}
        raw = self.rfile.read(length).decode("utf-8")
        try:
            payload = json.loads(raw or "{}")
        except json.JSONDecodeError as exc:
            raise ValueError(f"Neispravan JSON: {exc}") from exc
        return payload if isinstance(payload, dict) else {}

    def is_authorised(self) -> bool:
        if not app_password():
            return True
        cookie = SimpleCookie(self.headers.get("Cookie", ""))
        token = cookie[SESSION_COOKIE].value if SESSION_COOKIE in cookie else ""
        return token in SESSIONS

    def guard(self) -> bool:
        if self.is_authorised():
            return True
        self.send_json({"error": "Potrebna je prijava.", "authRequired": True}, status=HTTPStatus.UNAUTHORIZED)
        return False

    # -- routes ----------------------------------------------------------- #

    def do_GET(self) -> None:
        route = urlparse(self.path)
        query = parse_qs(route.query)

        if route.path in {"", "/", "/index.html"}:
            self.send_file(WEB_DIR / "index.html")
            return

        if route.path.startswith("/static/"):
            name = route.path[len("/static/") :]
            candidate = (WEB_DIR / name).resolve()
            if WEB_DIR.resolve() in candidate.parents:
                self.send_file(candidate)
            else:
                self.send_error(HTTPStatus.FORBIDDEN)
            return

        if route.path == "/api/session":
            self.send_json({"authRequired": bool(app_password()), "authenticated": self.is_authorised()})
            return

        if not route.path.startswith("/api/"):
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        if not self.guard():
            return

        handlers: dict[str, Callable[[], None]] = {
            "/api/bootstrap": lambda: self.send_json(self._bootstrap()),
            "/api/state": lambda: self.send_json(self._run_state(query)),
            "/api/stations": lambda: self.send_json(stations_payload()),
            "/api/reports": lambda: self.send_json(reports_payload()),
        }
        handler = handlers.get(route.path)
        if handler:
            handler()
            return

        if route.path == "/api/report":
            selected_date = str(query.get("date", [yesterday()])[0])
            if not _DATE_RE.match(selected_date):
                self.send_json({"error": "Neispravan datum."}, status=HTTPStatus.BAD_REQUEST)
                return
            path = report_path(selected_date)
            if not path.exists():
                self.send_json({"error": "Izvještaj za taj datum ne postoji."}, status=HTTPStatus.NOT_FOUND)
                return
            self.send_file(path, download_name=path.name)
            return

        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        route = urlparse(self.path)

        if route.path == "/api/login":
            try:
                payload = self.read_json()
            except ValueError as exc:
                self.send_json({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
                return
            token = authenticate(str(payload.get("password") or ""))
            if not token:
                self.send_json({"error": "Pogrešna lozinka."}, status=HTTPStatus.UNAUTHORIZED)
                return
            self.send_json({"ok": True}, cookie=token)
            return

        if not route.path.startswith("/api/"):
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        if not self.guard():
            return

        try:
            payload = self.read_json()
        except ValueError as exc:
            self.send_json({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return

        try:
            if route.path == "/api/run":
                self._start_run(payload)
                return
            if route.path == "/api/stop":
                self._stop(payload)
                return
            if route.path == "/api/stations":
                self.send_json(save_stations_payload(payload))
                return
            if route.path == "/api/stations/repair":
                self.send_json(repair_stations())
                return
            if route.path == "/api/hide-report":
                selected_date = str(payload.get("selectedDate") or "")
                if not _DATE_RE.match(selected_date):
                    self.send_json({"error": "Neispravan datum."}, status=HTTPStatus.BAD_REQUEST)
                    return
                hide_report(selected_date)
                self.send_json({"ok": True, **reports_payload()})
                return
            if route.path == "/api/unhide-reports":
                unhide_reports()
                self.send_json({"ok": True, **reports_payload()})
                return
            if route.path == "/api/open-folder":
                selected_date = str(payload.get("selectedDate") or "")
                if selected_date and not _DATE_RE.match(selected_date):
                    self.send_json({"error": "Neispravan datum."}, status=HTTPStatus.BAD_REQUEST)
                    return
                if open_reports_folder(selected_date or None):
                    self.send_json({"ok": True})
                else:
                    self.send_json(
                        {"error": f"Folder nije otvoren. Putanja je: {reports_dir()}"},
                        status=HTTPStatus.INTERNAL_SERVER_ERROR,
                    )
                return
            if route.path == "/api/open-report":
                selected_date = str(payload.get("selectedDate") or yesterday())
                if not _DATE_RE.match(selected_date):
                    self.send_json({"error": "Neispravan datum."}, status=HTTPStatus.BAD_REQUEST)
                    return
                if open_in_excel(report_path(selected_date)):
                    self.send_json({"ok": True})
                else:
                    self.send_json(
                        {"error": "Excel nije otvoren. Koristi dugme za preuzimanje."},
                        status=HTTPStatus.INTERNAL_SERVER_ERROR,
                    )
                return
        except (ValueError, RuntimeError) as exc:
            self.send_json({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return

        self.send_error(HTTPStatus.NOT_FOUND)

    # -- route bodies ------------------------------------------------------ #

    def _bootstrap(self) -> dict[str, Any]:
        rows = load_template_rows()
        path = stations_path()
        stations = registry.load_stations(path) if path.exists() else []
        return {
            "cloud": is_cloud(),
            "today": datetime.now().strftime(ISO_DATE),
            "defaultDate": yesterday(),
            "template": TEMPLATE_PATH.name,
            "templateExists": TEMPLATE_PATH.exists(),
            "stationCount": sum(1 for station in stations if station.enabled),
            "rowCount": len(rows),
            "reportsLocation": str(reports_dir()),
            "canOpenLocally": not is_cloud(),
        }

    def _run_state(self, query: dict[str, list[str]]) -> dict[str, Any]:
        cursor = _clamp(query.get("cursor", ["0"])[0], 0, 0, 10_000_000)
        with STATE_LOCK:
            lines = STATE.lines[cursor:]
            days = [day.to_json() for day in STATE.days]
            payload = {
                "running": STATE.running,
                "returnCode": STATE.return_code,
                "cursor": len(STATE.lines),
                "lines": lines,
                "logFile": STATE.log_path.name if STATE.log_path else "",
                "startedAt": STATE.started_at,
                # The single-day keys below describe the active day, so the
                # stats tiles and progress bar keep working unchanged.
                "selectedDate": STATE.selected_date,
                "done": STATE.done,
                "total": STATE.total,
                "current": STATE.current,
                "days": days,
                "batch": {
                    "total": len(days),
                    "index": min(STATE.index, max(0, len(days) - 1)),
                    "ok": sum(1 for day in days if day["status"] == "ok"),
                    "failed": sum(1 for day in days if day["status"] == "failed"),
                    "problems": sum(len(day["failures"]) for day in days),
                },
            }
        if not payload["running"] and payload["returnCode"] == 0 and payload["selectedDate"]:
            payload["reportReady"] = report_path(str(payload["selectedDate"])).exists()
        return payload

    def _start_run(self, payload: dict[str, Any]) -> None:
        dates = selected_dates(payload)
        if not TEMPLATE_PATH.exists():
            raise ValueError(f"Nedostaje master template: {TEMPLATE_PATH.name}")

        if bool(payload.get("onlyTelemetry")):
            # This pass only writes into a workbook a full run already made,
            # so the station list is irrelevant -- but every report must exist.
            missing = [date for date in dates if not report_path(date).exists()]
            if missing:
                raise ValueError(
                    f"Nema izvještaja za: {', '.join(missing)}. Prvo pokreni obračun, pa onda nivoe."
                )
        else:
            rows = load_template_rows()
            path = stations_path()
            stations = registry.load_stations(path) if path.exists() else []
            blocking = [issue for issue in registry.validate(stations, rows) if issue.severity == "error"]
            if blocking:
                raise ValueError(
                    "Lista stanica ima greške:\n" + "\n".join(f"{i.station}: {i.message}" for i in blocking)
                )

        for date in dates:
            unhide_report(date)
        log_name = start_batch(dates, payload)
        if len(dates) > 1:
            append_log(f"Obračun za {len(dates)} dana: {', '.join(dates)}.")
        self.send_json({"ok": True, "logFile": log_name, "dates": dates})

    def _stop(self, payload: dict[str, Any]) -> None:
        # "batch" (the default, and what a single-day run has always done)
        # ends the whole queue; "day" kills the running day and moves on.
        scope = str(payload.get("scope") or "batch")
        with STATE_LOCK:
            process = STATE.process
            remaining = len(STATE.days) - STATE.index - 1
            if scope == "day":
                STATE.stop_current = True
            else:
                STATE.cancelled = True
        if scope == "day":
            append_log("Preskakanje tekućeg dana je zatraženo.")
        elif remaining > 0:
            append_log(f"Zaustavljanje je zatraženo; preostalih {remaining} dana se preskače.")
        else:
            append_log("Zaustavljanje procesa je zatraženo.")
        if process and process.poll() is None:
            process.terminate()
            try:
                process.wait(timeout=8)
            except subprocess.TimeoutExpired:
                process.kill()
        self.send_json({"ok": True})


def serve(host: str, port: int, open_browser: bool = False) -> None:
    server = ThreadingHTTPServer((host, port), Handler)
    mode = "cloud" if is_cloud() else "local"
    print(f"EcoKing web UI ({mode}): http://{host}:{port}")
    if app_password():
        print("Password protection is on (APP_PASSWORD).")
    elif is_cloud():
        print("WARNING: APP_PASSWORD is not set, so anyone with the URL can start a run.")

    if open_browser:
        import threading
        import webbrowser

        # The server isn't listening yet at this point, so the tab is opened
        # from a short-lived timer instead of blocking serve_forever below.
        browse_host = "127.0.0.1" if host in {"0.0.0.0", "::"} else host
        threading.Timer(0.6, lambda: webbrowser.open(f"http://{browse_host}:{port}")).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


def main() -> None:
    import argparse

    from dotenv import load_dotenv

    load_dotenv(dotenv_path=ROOT / ".env")
    parser = argparse.ArgumentParser(description="Run the EcoKing web UI.")
    parser.add_argument("--host", default=os.getenv("HOST", "127.0.0.1"))
    parser.add_argument("--port", type=int, default=int(os.getenv("PORT", "8765")))
    parser.add_argument(
        "--no-browser", action="store_true", help="Don't open a browser tab automatically."
    )
    args = parser.parse_args()
    serve(args.host, args.port, open_browser=not args.no_browser and not is_cloud())


if __name__ == "__main__":
    main()
