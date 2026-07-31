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
from dataclasses import dataclass, field
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

ROOT = Path(__file__).resolve().parent.parent
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


@dataclass
class RunState:
    process: subprocess.Popen[str] | None = None
    running: bool = False
    return_code: int | None = None
    log_path: Path | None = None
    lines: list[str] = field(default_factory=list)
    started_at: str | None = None
    selected_date: str | None = None
    done: int = 0
    total: int = 0
    current: str = ""


STATE = RunState()
STATE_LOCK = threading.Lock()


def append_log(line: str) -> None:
    translated = logtext.translate(line)
    if not translated:
        return
    with STATE_LOCK:
        STATE.lines.append(translated)
        progress = _STATION_LINE_RE.search(translated)
        if progress:
            STATE.done = int(progress.group(1))
            STATE.total = int(progress.group(2))
            label = translated.split("Stanica=", 1)[-1].split(", Excel red=", 1)[0]
            STATE.current = label.strip()
        log_path = STATE.log_path
    if log_path:
        with log_path.open("a", encoding="utf-8") as handle:
            handle.write(translated + "\n")


def start_process(cmd: list[str], environment: dict[str, str], selected_date: str | None) -> str:
    """Reset the log, spawn the scraper, and stream its output into the state."""
    with STATE_LOCK:
        if STATE.running:
            raise RuntimeError("Obračun je već u toku.")
        STATE.lines.clear()
        STATE.return_code = None
        STATE.log_path = log_dir() / f"ecoking-{datetime.now().strftime('%Y%m%d-%H%M%S')}.log"
        STATE.started_at = datetime.now().strftime("%d.%m.%Y. %H:%M:%S")
        STATE.running = True
        STATE.done = 0
        STATE.total = 0
        STATE.current = ""
        STATE.selected_date = selected_date
        log_path = STATE.log_path

    threading.Thread(target=_pump_process, args=(cmd, environment), daemon=True).start()
    return str(log_path.name)


def _pump_process(cmd: list[str], environment: dict[str, str]) -> None:
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
        STATE.running = False
        STATE.process = None
        STATE.return_code = return_code
    append_log(
        "Obračun je završen uspješno."
        if return_code == 0
        else f"Obračun je završen sa greškom. Kod izlaza: {return_code}"
    )


def build_run_command(config: dict[str, Any]) -> tuple[list[str], dict[str, str], str]:
    selected_date = str(config.get("selectedDate") or yesterday())
    output = report_path(selected_date)
    cmd = [
        sys.executable,
        str(SCRAPER),
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

    environment = os.environ.copy()
    environment["NO_COLOR"] = "1"
    environment["PYTHONIOENCODING"] = "utf-8"
    environment["PYTHONUNBUFFERED"] = "1"
    environment["CHART_WAIT_MS"] = str(_clamp(config.get("chartWait"), 5000, 1000, 60000))
    environment["SEARCH_RESULTS_WAIT_MS"] = str(_clamp(config.get("searchWait"), 2000, 500, 30000))
    environment["COPY_REPORT_TO_DESKTOP"] = "0" if is_cloud() else "1"
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
    stations = [Station.from_json(item) for item in incoming if isinstance(item, dict)]
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


def list_reports() -> list[dict[str, Any]]:
    reports = []
    for item in sorted(reports_dir().glob("EcoKing_Report_*.xlsx"), reverse=True):
        try:
            stat = item.stat()
        except OSError:
            continue
        reports.append(
            {
                "date": item.stem.replace("EcoKing_Report_", ""),
                "name": item.name,
                "size": stat.st_size,
                "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%d.%m.%Y. %H:%M"),
            }
        )
    return reports[:30]


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
    for command in commands:
        if not shutil.which(command[0]):
            continue
        try:
            subprocess.Popen(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
        except OSError:
            continue
    append_log(f"UPOZORENJE: Excel nije otvoren automatski. Fajl je na: {path}")
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
            "/api/reports": lambda: self.send_json({"reports": list_reports()}),
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
                self._stop()
                return
            if route.path == "/api/stations":
                self.send_json(save_stations_payload(payload))
                return
            if route.path == "/api/stations/repair":
                self.send_json(repair_stations())
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
            payload = {
                "running": STATE.running,
                "returnCode": STATE.return_code,
                "cursor": len(STATE.lines),
                "lines": lines,
                "logFile": STATE.log_path.name if STATE.log_path else "",
                "startedAt": STATE.started_at,
                "selectedDate": STATE.selected_date,
                "done": STATE.done,
                "total": STATE.total,
                "current": STATE.current,
            }
        if not payload["running"] and payload["returnCode"] == 0 and payload["selectedDate"]:
            payload["reportReady"] = report_path(str(payload["selectedDate"])).exists()
        return payload

    def _start_run(self, payload: dict[str, Any]) -> None:
        selected_date = str(payload.get("selectedDate") or yesterday())
        if not _DATE_RE.match(selected_date):
            raise ValueError("Datum mora biti u formatu YYYY-MM-DD.")
        if datetime.strptime(selected_date, ISO_DATE).date() > datetime.now().date():
            raise ValueError("Datum ne može biti u budućnosti.")
        if not TEMPLATE_PATH.exists():
            raise ValueError(f"Nedostaje master template: {TEMPLATE_PATH.name}")

        rows = load_template_rows()
        path = stations_path()
        stations = registry.load_stations(path) if path.exists() else []
        blocking = [issue for issue in registry.validate(stations, rows) if issue.severity == "error"]
        if blocking:
            raise ValueError(
                "Lista stanica ima greške:\n" + "\n".join(f"{i.station}: {i.message}" for i in blocking)
            )

        cmd, environment, selected_date = build_run_command({**payload, "selectedDate": selected_date})
        log_name = start_process(cmd, environment, selected_date=selected_date)
        append_log(f"Generisanje izvještaja za datum {selected_date}.")
        append_log(f"Izlazni fajl: {report_path(selected_date).name}")
        self.send_json({"ok": True, "logFile": log_name})

    def _stop(self) -> None:
        with STATE_LOCK:
            process = STATE.process
        if process and process.poll() is None:
            append_log("Zaustavljanje procesa je zatraženo.")
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
