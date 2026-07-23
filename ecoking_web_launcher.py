from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse


ROOT = Path(__file__).resolve().parent
DEFAULT_WORKBOOK = ROOT / "EcoKing - tabela potrošnje - Jul 2026. TEST.xlsx"
DEFAULT_STATIONS = ROOT / "herceg_novi_stations.json"
LOG_DIR = ROOT / "logs"
ANSI_RE = re.compile(r"\x1b\[[0-9;]*m")


@dataclass
class RunState:
    process: subprocess.Popen[str] | None = None
    running: bool = False
    return_code: int | None = None
    log_path: Path | None = None
    lines: list[str] = field(default_factory=list)
    started_at: str | None = None


STATE = RunState()
STATE_LOCK = threading.Lock()


def strip_ansi(text: str) -> str:
    return ANSI_RE.sub("", text).rstrip("\r\n")


def short_path(path: str | Path) -> str:
    try:
        return str(Path(path).resolve().relative_to(ROOT))
    except Exception:
        return str(path)


def translate_log_line(line: str) -> str:
    clean = strip_ansi(line)
    if not clean:
        return ""

    prefix = ""
    body = clean
    match = re.match(r"^(\d{2}:\d{2}:\d{2})\s+([A-Z]+)\s+(.*)$", clean)
    if match:
        time_part, level, body = match.groups()
        level_sr = {
            "DEBUG": "DETALJ",
            "INFO": "INFO",
            "WARNING": "UPOZORENJE",
            "ERROR": "GREŠKA",
            "CRITICAL": "KRITIČNO",
        }.get(level, level)
        prefix = f"{time_part} {level_sr:<10} "

    station_match = re.match(
        r"^\[(\d+)/(\d+)\] Station key=(.+), Excel row=(.+), Excel LOKACIJA=(.+), VODOMJER=(.+), browser search=(.+)$",
        body,
    )
    if station_match:
        idx, total, key, row, location, meter, search = station_match.groups()
        return f"{prefix}[{idx}/{total}] Stanica={key}, Excel red={row}, Lokacija={location}, Vodomjer={meter}, pretraga={search}"

    replacements: list[tuple[re.Pattern[str], str]] = [
        (re.compile(r"^Loaded (\d+) workbook rows from (.+)$"), r"Učitano je \1 redova iz Excel fajla \2"),
        (re.compile(r"^Loaded (\d+) location mappings from (.+)$"), r"Učitano je \1 mapiranja lokacija iz \2"),
        (re.compile(r"^Built (\d+) station-driven scrape jobs from (\d+) mappings$"), r"Pripremljeno je \1 zadataka iz \2 mapiranja stanica"),
        (re.compile(r"^Launching browser: headless=(\w+) slow_mo_ms=(\d+)$"), r"Pokrećem browser: skriven=\1, usporenje=\2 ms"),
        (re.compile(r"^Running (\d+) jobs across (\d+) browser workers\.$"), r"Pokrećem \1 zadataka kroz \2 paralelna browser procesa."),
        (re.compile(r"^Opening (.+)$"), r"Otvaram stranicu \1"),
        (re.compile(r"^Login flow submitted$"), "Prijava je poslata"),
        (re.compile(r"^Searching location: (.+)$"), r"Tražim lokaciju: \1"),
        (re.compile(r"^Clicked device dropdown using precise top-left selector$"), "Otvoren je padajući meni za izbor uređaja"),
        (re.compile(r"^Filling location search with visible dropdown input (.+)$"), r"Unosim vrijednost u polje pretrage (\1)"),
        (re.compile(r"^Retyped location search with keyboard into (.+)$"), r"Ponovo unosim pretragu preko tastature (\1)"),
        (re.compile(r"^No dropdown results for (.+) after DOM fill\. Retrying with keyboard typing\.$"), r"Nema rezultata za \1 nakon prvog unosa. Pokušavam ponovo unosom preko tastature."),
        (re.compile(r"^Filtered dropdown options for (.+): (.+)$"), r"Rezultati u padajućem meniju za \1: \2"),
        (re.compile(r"^Choosing dropdown result: (.+)$"), r"Biranje pronađenog uređaja: \1"),
        (re.compile(r"^Choosing only visible dropdown result: (.+)$"), r"Biranje jedinog vidljivog uređaja: \1"),
        (re.compile(r"^Selected (.+) by fallback serial search (.+)$"), r"Lokacija \1 je izabrana preko rezervne pretrage serijskog broja \2"),
        (re.compile(r"^Selecting interval: (.+)$"), r"Biranje intervala: \1"),
        (re.compile(r"^Clicked interval option '(.+)'$"), r"Kliknut je interval '\1'"),
        (re.compile(r"^Clicking interval button with selector (.+)$"), r"Klik na dugme intervala preko selektora \1"),
        (re.compile(r"^Read (.+): daily=(.+) m3, max=(.+) m3, min=(.+) m3$"), r"Očitano za \1: dnevno=\2 m3, maksimum=\3 m3, minimum=\4 m3"),
        (re.compile(r'^FOUND: MIN: (.+) MAX: (.+) DAILY: (.+) for "(.+)"$'), r'PRONAĐENO: MIN=\1, MAX=\2, DNEVNO=\3 za "\4"'),
        (re.compile(r"^Failed to scrape station (.+) for Excel row (.+)$"), r"Neuspjelo očitavanje stanice \1 za Excel red \2"),
        (re.compile(r"^Search query (.+) failed for (.+)\. Trying fallback query\.$"), r"Pretraga \1 nije uspjela za \2. Pokušavam rezervnu pretragu."),
        (re.compile(r"^Saved debug artifacts: (.+)\.png and (.+)\.html$"), r"Sačuvani su debug fajlovi: \1.png i \2.html"),
        (re.compile(r"^Created sheet (.+) in (.+) with (\d+) scraped rows$"), r"Kreiran je list \1 u fajlu \2 sa \3 očitanih redova"),
        (re.compile(r"^========== EXECUTION REPORT ==========$"), "========== IZVJEŠTAJ IZVRŠENJA =========="),
        (re.compile(r"^SUCCESSFUL: (\d+)$"), r"USPJEŠNO: \1"),
        (re.compile(r"^NO DATA / NO ENTRIES: (\d+)$"), r"BEZ PODATAKA / BEZ UNOSA: \1"),
        (re.compile(r"^FAILED: (\d+)$"), r"NEUSPJELO: \1"),
        (re.compile(r"^======================================$"), "======================================"),
        (re.compile(r"^Run failed: (.+)$"), r"Pokretanje nije uspjelo: \1"),
    ]
    for pattern, replacement in replacements:
        if pattern.search(body):
            return prefix + pattern.sub(replacement, body)
    return prefix + body


def append_log(line: str) -> None:
    translated = translate_log_line(line)
    if not translated:
        return
    with STATE_LOCK:
        STATE.lines.append(translated)
        log_path = STATE.log_path
    if log_path:
        with log_path.open("a", encoding="utf-8") as handle:
            handle.write(translated + "\n")


def open_local_file(path: Path) -> bool:
    path = path.resolve()
    if not path.exists():
        append_log(f"UPOZORENJE: Excel fajl ne postoji i ne može se otvoriti: {path}")
        return False

    if sys.platform.startswith("win"):
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
            append_log(f"Otvaram ažurirani Excel fajl: {path}")
            return True
        except Exception as exc:
            append_log(f"UPOZORENJE: Nije moguće automatski otvoriti Excel fajl: {exc}")
            return False

    if sys.platform == "darwin":
        candidates = [["open", str(path)]]
    else:
        candidates = [
            ["gio", "open", str(path)],
            ["xdg-open", str(path)],
            ["libreoffice", "--calc", str(path)],
            ["localc", str(path)],
            ["soffice", "--calc", str(path)],
        ]

    errors: list[str] = []
    for command in candidates:
        executable = command[0]
        if not shutil.which(executable):
            errors.append(f"{executable}: nije instaliran")
            continue
        try:
            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            try:
                stdout, stderr = process.communicate(timeout=2)
            except subprocess.TimeoutExpired:
                append_log(f"Otvaram ažurirani Excel fajl preko komande: {' '.join(command[:2])}")
                return True

            if process.returncode == 0:
                append_log(f"Otvaram ažurirani Excel fajl preko komande: {' '.join(command[:2])}")
                return True

            details = (stderr or stdout or "").strip()
            errors.append(f"{' '.join(command[:2])}: {details or f'kod izlaza {process.returncode}'}")
        except Exception as exc:
            errors.append(f"{executable}: {exc}")

    append_log("UPOZORENJE: Nije moguće automatski otvoriti Excel fajl.")
    for error in errors[:5]:
        append_log(f"UPOZORENJE: {error}")
    append_log(f"Excel fajl možeš otvoriti ručno: {path}")
    return False


def choose_workbook_with_native_dialog(initial_path: Path) -> tuple[Path | None, str | None]:
    initial_path = initial_path.expanduser()
    initial_dir = initial_path.parent if initial_path.suffix else initial_path
    if not initial_dir.exists():
        initial_dir = ROOT
    initial_dir = initial_dir.resolve()

    if sys.platform.startswith("win"):
        return None, "Native file dialog for web preview is only implemented for Linux/macOS. Use the Windows desktop launcher for native Windows file picking."

    candidates: list[list[str]]
    if sys.platform == "darwin":
        script = (
            'POSIX path of (choose file with prompt "Izaberi Excel fajl" '
            'of type {"org.openxmlformats.spreadsheetml.sheet", "com.microsoft.excel.xls"})'
        )
        candidates = [["osascript", "-e", script]]
    else:
        candidates = [
            [
                "zenity",
                "--file-selection",
                "--title=Izaberi Excel fajl",
                f"--filename={initial_dir}/",
                "--file-filter=Excel fajlovi | *.xlsx *.xlsm",
                "--file-filter=Svi fajlovi | *",
            ],
            [
                "kdialog",
                "--title",
                "Izaberi Excel fajl",
                "--getopenfilename",
                str(initial_dir),
                "*.xlsx *.xlsm|Excel fajlovi",
            ],
        ]

    missing: list[str] = []
    for command in candidates:
        if not shutil.which(command[0]):
            missing.append(command[0])
            continue
        try:
            result = subprocess.run(command, capture_output=True, text=True, timeout=120)
        except subprocess.TimeoutExpired:
            return None, "File dialog je istekao bez izbora fajla."
        except Exception as exc:
            return None, f"File dialog nije moguće otvoriti: {exc}"

        if result.returncode == 0:
            selected = result.stdout.strip().splitlines()[0] if result.stdout.strip() else ""
            if not selected:
                return None, "Nije izabran fajl."
            return Path(selected), None
        if result.returncode in {1, 5}:
            return None, "Izbor fajla je otkazan."
        details = (result.stderr or result.stdout or "").strip()
        return None, details or f"{command[0]} je završio sa kodom {result.returncode}."

    return None, f"Nije pronađen sistemski file dialog alat ({', '.join(missing)}). Instaliraj zenity ili kdialog, ili koristi ugrađeni web locator."


def build_command(config: dict[str, object]) -> tuple[list[str], dict[str, str]]:
    cmd = [
        sys.executable,
        str(ROOT / "ecoking_daily.py"),
        "--workbook",
        str(config.get("workbook") or DEFAULT_WORKBOOK),
        "--workers",
        str(max(1, int(config.get("workers") or 1))),
        "--slow-mo-ms",
        str(max(0, int(config.get("slowMo") or 0))),
    ]
    cmd.append("--headed" if bool(config.get("browserVisible")) else "--headless")
    if bool(config.get("verbose", True)):
        cmd.append("--verbose")
    if bool(config.get("keepOpen")):
        cmd.append("--keep-browser-open")
    limit = str(config.get("limit") or "").strip()
    if limit:
        cmd.extend(["--limit", limit])
    env = os.environ.copy()
    env["NO_COLOR"] = "1"
    env["LOCATION_MAP_PATH"] = str(DEFAULT_STATIONS)
    env["CHART_WAIT_MS"] = str(max(1000, int(config.get("chartWait") or 5000)))
    env["SEARCH_RESULTS_WAIT_MS"] = str(max(500, int(config.get("searchWait") or 2000)))
    env["PYTHONIOENCODING"] = "utf-8"
    return cmd, env


def validate_config(config: dict[str, object]) -> list[str]:
    errors: list[str] = []
    for label, key, default in [
        ("Excel ulaz", "workbook", DEFAULT_WORKBOOK),
        ("Mapiranje stanica", "stationMap", DEFAULT_STATIONS),
        ("Skripta", "script", ROOT / "ecoking_daily.py"),
    ]:
        path = Path(str(config.get(key) or default))
        if not path.exists():
            errors.append(f"{label} ne postoji: {path}")
    try:
        if str(config.get("limit") or "").strip() and int(str(config.get("limit"))) < 1:
            errors.append("Limit mora biti pozitivan broj.")
        for key in ["workers", "slowMo", "chartWait", "searchWait"]:
            int(config.get(key) or 0)
    except ValueError:
        errors.append("Provjerite numerička polja i datum. Datum mora biti dd.mm.yyyy.")
    return errors


def run_process(cmd: list[str], env: dict[str, str], workbook_path: Path, open_after: bool) -> None:
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
            env=env,
        )
        with STATE_LOCK:
            STATE.process = process
        assert process.stdout is not None
        for raw_line in process.stdout:
            append_log(raw_line)
        return_code = process.wait()
    except Exception as exc:
        append_log(f"GREŠKA: Pokretanje procesa nije uspjelo: {exc}")
        return_code = 1

    with STATE_LOCK:
        STATE.running = False
        STATE.process = None
        STATE.return_code = return_code
    if return_code == 0:
        append_log("Obračun je završen uspješno.")
        if open_after:
            open_local_file(workbook_path)
    else:
        append_log(f"Obračun je završen sa greškom. Kod izlaza: {return_code}")


HTML = r"""<!doctype html>
<html lang="sr">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>EcoKing dnevni obračun</title>
  <style>
    :root {
      --blue: #0e6ea8;
      --blue-dark: #0a527e;
      --ink: #17202a;
      --muted: #5d6b7a;
      --line: #d8e1ea;
      --bg: #eef3f8;
      --panel: #ffffff;
      --soft: #f7fafc;
      --shade: rgba(15, 23, 42, .52);
      --green: #147a50;
      --red: #c24141;
      --yellow: #9a6700;
      --log-bg: #0d1620;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: var(--bg);
      color: var(--ink);
      font: 14px/1.45 Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }
    .app-header {
      background: linear-gradient(180deg, var(--blue), var(--blue-dark));
      color: white;
      padding: 18px 22px;
      border-bottom: 1px solid #083e61;
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 18px;
    }
    .brand { display: flex; align-items: center; gap: 12px; }
    .brand-mark {
      width: 40px;
      height: 40px;
      border-radius: 10px;
      display: grid;
      place-items: center;
      background: rgba(255,255,255,.14);
      border: 1px solid rgba(255,255,255,.22);
      font-weight: 800;
      letter-spacing: .02em;
    }
    .app-header h1 { margin: 0; font-size: 20px; line-height: 1.2; }
    .app-header p { margin: 3px 0 0; color: #d9edf7; font-size: 13px; }
    .header-meta {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      justify-content: flex-end;
    }
    .pill {
      border: 1px solid rgba(255,255,255,.25);
      background: rgba(255,255,255,.12);
      color: white;
      border-radius: 999px;
      padding: 6px 10px;
      font-size: 12px;
      font-weight: 700;
      white-space: nowrap;
    }
    main {
      display: grid;
      grid-template-columns: 380px minmax(0, 1fr);
      gap: 18px;
      padding: 18px;
      min-height: calc(100vh - 77px);
    }
    .panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 16px;
      margin-bottom: 12px;
      box-shadow: 0 1px 2px rgba(15, 23, 42, .04);
    }
    .panel-title {
      display: flex;
      align-items: baseline;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 12px;
    }
    h2 { margin: 0; font-size: 15px; }
    .section-note { margin: 0; color: var(--muted); font-size: 12px; }
    label { display: block; font-weight: 700; margin: 8px 0 5px; }
    input[type="text"], input[type="number"] {
      width: 100%;
      min-height: 38px;
      padding: 9px 10px;
      border: 1px solid #c6d4e1;
      border-radius: 6px;
      font: inherit;
      background: white;
      color: var(--ink);
      outline: none;
    }
    input[type="text"]:focus, input[type="number"]:focus {
      border-color: var(--blue);
      box-shadow: 0 0 0 3px rgba(14, 110, 168, .12);
    }
    .grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
    .field-help { color: var(--muted); font-size: 12px; margin: 4px 0 0; }
    .check {
      display: grid;
      grid-template-columns: 34px minmax(0, 1fr);
      gap: 10px;
      align-items: start;
      padding: 10px;
      border: 1px solid var(--line);
      border-radius: 8px;
      margin: 8px 0;
      background: var(--soft);
      font-weight: 500;
    }
    .check input { width: 18px; height: 18px; margin: 2px 0 0; accent-color: var(--blue); }
    .check strong { display: block; font-size: 13px; }
    .check span { display: block; color: var(--muted); font-size: 12px; margin-top: 2px; }
    .buttons { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
    button {
      border: 0;
      border-radius: 6px;
      min-height: 38px;
      padding: 10px 14px;
      font: inherit;
      font-weight: 700;
      cursor: pointer;
      background: #d7e2eb;
      color: var(--ink);
      transition: transform .08s ease, box-shadow .12s ease, background .12s ease;
    }
    button:hover { box-shadow: 0 2px 8px rgba(15, 23, 42, .12); }
    button:active { transform: translateY(1px); }
    button.primary { background: var(--blue); color: white; }
    button.danger { background: var(--red); color: white; }
    button:disabled { opacity: .55; cursor: not-allowed; }
    button.secondary { background: #e6eef5; color: #24445c; }
    .status {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      margin-bottom: 10px;
    }
    .badge {
      border-radius: 999px;
      padding: 5px 10px;
      background: #dce8f2;
      color: #174b70;
      font-weight: 700;
      white-space: nowrap;
    }
    .badge.running { background: #fff0c2; color: var(--yellow); }
    .badge.done { background: #dff4e8; color: var(--green); }
    .badge.error { background: #ffe0e0; color: var(--red); }
    .log-shell {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 14px;
      box-shadow: 0 1px 2px rgba(15, 23, 42, .04);
      min-height: 100%;
      display: grid;
      grid-template-rows: auto auto minmax(0, 1fr);
    }
    .run-metrics {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 10px;
      margin-bottom: 12px;
    }
    .metric {
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--soft);
      padding: 10px;
    }
    .metric span { display: block; color: var(--muted); font-size: 11px; font-weight: 700; text-transform: uppercase; }
    .metric strong { display: block; font-size: 14px; margin-top: 3px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .progress {
      height: 8px;
      background: #dfe8f0;
      border-radius: 999px;
      overflow: hidden;
      margin-bottom: 12px;
    }
    .progress > div {
      height: 100%;
      width: 0%;
      background: linear-gradient(90deg, var(--blue), #22a6b3);
      transition: width .2s ease;
    }
    pre {
      height: calc(100vh - 270px);
      min-height: 420px;
      margin: 0;
      padding: 14px 16px;
      overflow: auto;
      background: var(--log-bg);
      color: #edf3f7;
      border-radius: 8px;
      border: 1px solid #22313f;
      white-space: pre-wrap;
      word-break: break-word;
      font: 13px/1.45 Consolas, "Liberation Mono", monospace;
    }
    .hint { color: var(--muted); font-size: 12px; margin-top: 6px; }
    .command {
      color: var(--muted);
      font-size: 12px;
      word-break: break-all;
      margin-top: 8px;
      display: none;
    }
    .command.open { display: block; }
    .inline-actions {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      margin-top: 10px;
    }
    .path-row {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 8px;
      align-items: center;
    }
    .modal {
      position: fixed;
      inset: 0;
      display: none;
      align-items: center;
      justify-content: center;
      background: var(--shade);
      padding: 24px;
      z-index: 20;
    }
    .modal.open { display: flex; }
    .dialog {
      width: min(760px, 96vw);
      max-height: 84vh;
      overflow: hidden;
      background: white;
      border-radius: 8px;
      border: 1px solid var(--line);
      box-shadow: 0 20px 50px rgba(0,0,0,.22);
      display: grid;
      grid-template-rows: auto auto minmax(260px, 1fr) auto;
    }
    .dialog header {
      background: white;
      color: var(--ink);
      padding: 14px;
      border-bottom: 1px solid var(--line);
    }
    .dialog header h2 { margin: 0; }
    .dialog-path {
      padding: 10px 14px;
      border-bottom: 1px solid var(--line);
      color: var(--muted);
      word-break: break-all;
      font-size: 13px;
    }
    .file-list {
      overflow: auto;
      padding: 8px;
    }
    .file-item {
      width: 100%;
      text-align: left;
      border-radius: 6px;
      background: white;
      font-weight: 500;
      padding: 9px 10px;
    }
    .file-item:hover { background: #edf4fa; }
    .dialog-actions {
      display: flex;
      justify-content: flex-end;
      gap: 8px;
      padding: 12px 14px;
      border-top: 1px solid var(--line);
    }
    @media (max-width: 980px) {
      main { grid-template-columns: 1fr; }
      pre { height: 480px; }
      .app-header { align-items: flex-start; flex-direction: column; }
    }
  </style>
</head>
<body>
  <header class="app-header">
    <div class="brand">
      <div class="brand-mark">EK</div>
      <div>
        <h1>EcoKing dnevni obračun</h1>
        <p>Izaberi Excel, pokreni očitavanje i prati tok izvršenja.</p>
      </div>
    </div>
    <div class="header-meta">
      <span class="pill">Upis u izabrani Excel</span>
      <span class="pill">List: jučerašnji datum</span>
    </div>
  </header>
  <main>
    <section>
      <div class="panel">
        <div class="panel-title">
          <h2>Excel fajl</h2>
          <p class="section-note">Obavezan ulaz</p>
        </div>
        <label for="workbook">Radna tabela</label>
        <div class="path-row">
          <input id="workbook" type="text">
          <button id="browse" type="button" class="secondary">File explorer</button>
        </div>
        <div class="hint">Otvara sistemski file explorer. Ako Linux nema <code>zenity</code> ili <code>kdialog</code>, koristi se web locator kao fallback.</div>
      </div>
      <div class="panel">
        <div class="panel-title">
          <h2>Pokretanje</h2>
          <p class="section-note">Najčešće ne mijenjati</p>
        </div>
        <label class="check">
          <input id="browserVisible" type="checkbox">
          <span><strong>Prikaži browser</strong><span>Uključi kada želiš da vidiš šta automatizacija klikće. Isključi za brži rad.</span></span>
        </label>
        <label class="check">
          <input id="verbose" type="checkbox" checked>
          <span><strong>Detaljni logovi</strong><span>Prikazuje svaki korak: izbor stanice, intervale, očitane vrijednosti i greške.</span></span>
        </label>
        <label class="check">
          <input id="openAfter" type="checkbox" checked>
          <span><strong>Otvori Excel nakon završetka</strong><span>Po uspješnom završetku otvara ažurirani Excel fajl.</span></span>
        </label>
        <label class="check">
          <input id="keepOpen" type="checkbox">
          <span><strong>Ostavi browser otvoren</strong><span>Samo za provjeru problema. Ručno zatvori terminal ili proces kada završiš pregled.</span></span>
        </label>
        <div class="grid2">
          <div>
            <label for="workers">Radnici</label>
            <input id="workers" type="number" min="1" max="8" value="1">
            <p class="field-help">Broj paralelnih browsera. Za vidljiv browser koristi 1.</p>
          </div>
          <div>
            <label for="limit">Limit</label>
            <input id="limit" type="number" min="1" placeholder="sve stanice">
            <p class="field-help">Ograniči broj stanica za probni run.</p>
          </div>
          <div>
            <label for="slowMo">Usporenje ms</label>
            <input id="slowMo" type="number" min="0" value="0">
            <p class="field-help">Pauza između browser akcija. Korisno samo kada pratiš klikove.</p>
          </div>
          <div>
            <label for="chartWait">Čekanje grafa ms</label>
            <input id="chartWait" type="number" min="1000" value="5000">
            <p class="field-help">Koliko dugo se čeka da se graf učita prije greške.</p>
          </div>
          <div>
            <label for="searchWait">Čekanje pretrage ms</label>
            <input id="searchWait" type="number" min="500" value="2000">
            <p class="field-help">Koliko dugo se čekaju rezultati u padajućem meniju.</p>
          </div>
        </div>
      </div>
      <div class="panel">
        <div class="panel-title">
          <h2>Kontrola</h2>
        </div>
        <div class="buttons">
          <button id="run" class="primary">Pokreni obračun</button>
          <button id="stop" class="danger" disabled>Zaustavi</button>
        </div>
        <div id="message" class="hint"></div>
        <div class="inline-actions">
          <button id="openWorkbook" type="button" class="secondary">Otvori Excel</button>
          <button id="toggleCommand" type="button" class="secondary">Prikaži komandu</button>
        </div>
        <div id="command" class="command"></div>
      </div>
    </section>
    <section class="log-shell">
      <div class="status">
        <h2>Log izvršenja</h2>
        <span id="badge" class="badge">Spremno</span>
      </div>
      <div class="run-metrics">
        <div class="metric"><span>Napredak</span><strong id="progressText">0 / 0</strong></div>
        <div class="metric"><span>Trenutna stanica</span><strong id="currentStation">-</strong></div>
        <div class="metric"><span>Log fajl</span><strong id="logPath">-</strong></div>
      </div>
      <div class="progress"><div id="progressBar"></div></div>
      <pre id="log"></pre>
    </section>
  </main>
  <div id="fileModal" class="modal">
    <div class="dialog">
      <header><h2>Izaberi Excel fajl</h2></header>
      <div id="dialogPath" class="dialog-path"></div>
      <div id="fileList" class="file-list"></div>
      <div class="dialog-actions">
        <button id="closeDialog" type="button">Zatvori</button>
      </div>
    </div>
  </div>
  <script>
    const defaults = {
      workbook: "__WORKBOOK__",
      browserVisible: true,
      openAfter: true,
      workers: 1,
      slowMo: 0,
      chartWait: 5000,
      searchWait: 2000,
      verbose: true
    };
    const ids = ["workbook", "workers", "limit", "slowMo", "chartWait", "searchWait"];
    for (const id of ids) if (defaults[id] !== undefined) document.getElementById(id).value = defaults[id];
    for (const id of ["browserVisible", "verbose", "keepOpen", "openAfter"]) if (defaults[id] !== undefined) document.getElementById(id).checked = defaults[id];

    let cursor = 0;
    const log = document.getElementById("log");
    const badge = document.getElementById("badge");
    const message = document.getElementById("message");
    const runBtn = document.getElementById("run");
    const stopBtn = document.getElementById("stop");
    const progressText = document.getElementById("progressText");
    const progressBar = document.getElementById("progressBar");
    const currentStation = document.getElementById("currentStation");
    const logPath = document.getElementById("logPath");
    const command = document.getElementById("command");
    const toggleCommand = document.getElementById("toggleCommand");
    const openWorkbook = document.getElementById("openWorkbook");

    function config() {
      return {
        workbook: document.getElementById("workbook").value,
        workers: Number(document.getElementById("workers").value || 1),
        limit: document.getElementById("limit").value,
        slowMo: Number(document.getElementById("slowMo").value || 0),
        chartWait: Number(document.getElementById("chartWait").value || 5000),
        searchWait: Number(document.getElementById("searchWait").value || 2000),
        browserVisible: document.getElementById("browserVisible").checked,
        verbose: document.getElementById("verbose").checked,
        keepOpen: document.getElementById("keepOpen").checked,
        openAfter: document.getElementById("openAfter").checked
      };
    }

    async function start() {
      message.textContent = "";
      const response = await fetch("/api/start", {method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(config())});
      const payload = await response.json();
      if (!response.ok) {
        message.textContent = payload.error || "Pokretanje nije uspjelo.";
        return;
      }
      cursor = 0;
      log.textContent = "";
      progressText.textContent = "0 / 0";
      progressBar.style.width = "0%";
      currentStation.textContent = "-";
      command.textContent = payload.command || "";
      poll();
    }

    async function stop() {
      await fetch("/api/stop", {method: "POST"});
      poll();
    }

    async function poll() {
      const response = await fetch(`/api/state?cursor=${cursor}`);
      const payload = await response.json();
      cursor = payload.cursor;
      if (payload.lines.length) {
        log.textContent += payload.lines.join("\n") + "\n";
        updateProgress(payload.lines);
        log.scrollTop = log.scrollHeight;
      }
      runBtn.disabled = payload.running;
      stopBtn.disabled = !payload.running;
      badge.className = "badge" + (payload.running ? " running" : payload.returnCode === 0 ? " done" : payload.returnCode ? " error" : "");
      badge.textContent = payload.running ? "U toku" : payload.returnCode === 0 ? "Završeno" : payload.returnCode ? "Greška" : "Spremno";
      if (payload.logPath) {
        message.textContent = `Log fajl: ${payload.logPath}`;
        logPath.textContent = payload.logPath;
      }
    }

    runBtn.addEventListener("click", start);
    stopBtn.addEventListener("click", stop);
    toggleCommand.addEventListener("click", () => {
      command.classList.toggle("open");
      toggleCommand.textContent = command.classList.contains("open") ? "Sakrij komandu" : "Prikaži komandu";
    });
    openWorkbook.addEventListener("click", async () => {
      message.textContent = "";
      const response = await fetch("/api/open-workbook", {method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(config())});
      const payload = await response.json();
      if (!response.ok) message.textContent = payload.error || "Excel nije otvoren.";
      poll();
    });
    document.getElementById("browse").addEventListener("click", chooseWorkbook);
    document.getElementById("closeDialog").addEventListener("click", () => document.getElementById("fileModal").classList.remove("open"));

    async function chooseWorkbook() {
      message.textContent = "Otvaram sistemski file explorer...";
      try {
        const response = await fetch("/api/browse-workbook", {method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(config())});
        const payload = await response.json();
        if (response.ok && payload.path) {
          document.getElementById("workbook").value = payload.path;
          message.textContent = "Excel fajl je izabran.";
          return;
        }
        message.textContent = payload.error || "Sistemski file explorer nije dostupan. Otvaram web locator.";
      } catch (error) {
        message.textContent = "Sistemski file explorer nije dostupan. Otvaram web locator.";
      }
      await openBrowser(document.getElementById("workbook").value);
    }

    function updateProgress(lines) {
      for (const line of lines) {
        const progress = line.match(/\[(\d+)\/(\d+)\]\s+Stanica=(.+?),\s+Excel red=/);
        if (progress) {
          const done = Number(progress[1]);
          const total = Number(progress[2]);
          progressText.textContent = `${done} / ${total}`;
          progressBar.style.width = total ? `${Math.min(100, Math.round(done * 100 / total))}%` : "0%";
          currentStation.textContent = progress[3].replace(/^'|'$/g, "");
        }
        const found = line.match(/PRONAĐENO: .* za "(.+)"$/);
        if (found) currentStation.textContent = found[1];
      }
    }

    async function openBrowser(path) {
      document.getElementById("fileModal").classList.add("open");
      await loadFiles(path);
    }

    async function loadFiles(path) {
      const response = await fetch(`/api/files?path=${encodeURIComponent(path || "")}`);
      const payload = await response.json();
      const list = document.getElementById("fileList");
      document.getElementById("dialogPath").textContent = payload.current || "";
      list.textContent = "";
      if (payload.parent) {
        const up = document.createElement("button");
        up.className = "file-item";
        up.textContent = "..";
        up.addEventListener("click", () => loadFiles(payload.parent));
        list.appendChild(up);
      }
      for (const item of payload.items || []) {
        const button = document.createElement("button");
        button.className = "file-item";
        button.textContent = `${item.type === "dir" ? "Folder" : "Excel"}  ${item.name}`;
        button.addEventListener("click", () => {
          if (item.type === "dir") {
            loadFiles(item.path);
          } else {
            document.getElementById("workbook").value = item.path;
            document.getElementById("fileModal").classList.remove("open");
          }
        });
        list.appendChild(button);
      }
    }
    setInterval(poll, 1000);
    poll();
  </script>
</body>
</html>
"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format: str, *args: object) -> None:
        return

    def send_json(self, payload: object, status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            html = (
                HTML.replace("__WORKBOOK__", str(DEFAULT_WORKBOOK))
            )
            body = html.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        if parsed.path == "/api/state":
            cursor = int(parse_qs(parsed.query).get("cursor", ["0"])[0] or 0)
            with STATE_LOCK:
                lines = STATE.lines[cursor:]
                next_cursor = len(STATE.lines)
                payload = {
                    "running": STATE.running,
                    "returnCode": STATE.return_code,
                    "cursor": next_cursor,
                    "lines": lines,
                    "logPath": short_path(STATE.log_path) if STATE.log_path else "",
                    "startedAt": STATE.started_at,
                }
            self.send_json(payload)
            return

        if parsed.path == "/api/files":
            raw_path = parse_qs(parsed.query).get("path", [""])[0]
            current = Path(raw_path).expanduser() if raw_path else ROOT
            if current.is_file():
                current = current.parent
            if not current.exists():
                current = ROOT
            try:
                current = current.resolve()
                entries = []
                for child in sorted(current.iterdir(), key=lambda item: (not item.is_dir(), item.name.lower())):
                    if child.name.startswith("."):
                        continue
                    if child.is_dir():
                        entries.append({"type": "dir", "name": child.name, "path": str(child)})
                    elif child.suffix.lower() in {".xlsx", ".xlsm"}:
                        entries.append({"type": "file", "name": child.name, "path": str(child)})
                payload = {
                    "current": str(current),
                    "parent": str(current.parent) if current.parent != current else "",
                    "items": entries,
                }
            except OSError as exc:
                payload = {"current": str(current), "parent": str(ROOT), "items": [], "error": str(exc)}
            self.send_json(payload)
            return

        self.send_error(404)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/start":
            length = int(self.headers.get("Content-Length", "0") or 0)
            body = self.rfile.read(length).decode("utf-8") if length else "{}"
            config = json.loads(body or "{}")
            errors = validate_config(config)
            if errors:
                self.send_json({"error": "\n".join(errors)}, status=400)
                return

            with STATE_LOCK:
                if STATE.running:
                    self.send_json({"error": "Obračun je već u toku."}, status=409)
                    return
                LOG_DIR.mkdir(exist_ok=True)
                STATE.lines.clear()
                STATE.return_code = None
                STATE.log_path = LOG_DIR / f"ecoking-web-{datetime.now().strftime('%Y%m%d-%H%M%S')}.log"
                STATE.started_at = datetime.now().strftime("%d.%m.%Y. %H:%M:%S")
                STATE.running = True

            cmd, env = build_command(config)
            workbook_path = Path(str(config.get("workbook") or DEFAULT_WORKBOOK))
            open_after = bool(config.get("openAfter", True))
            append_log("Pokretanje obračuna.")
            append_log(f"Log fajl: {short_path(STATE.log_path)}")
            append_log(f"Excel fajl: {short_path(workbook_path)}")
            append_log("Upis: direktno u isti Excel fajl, list za jučerašnji datum.")
            thread = threading.Thread(target=run_process, args=(cmd, env, workbook_path, open_after), daemon=True)
            thread.start()
            display = " ".join(f'"{part}"' if " " in part else part for part in cmd)
            self.send_json({"ok": True, "command": display, "logPath": short_path(STATE.log_path)})
            return

        if parsed.path == "/api/stop":
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
            return

        if parsed.path == "/api/open-workbook":
            length = int(self.headers.get("Content-Length", "0") or 0)
            body = self.rfile.read(length).decode("utf-8") if length else "{}"
            config = json.loads(body or "{}")
            workbook_path = Path(str(config.get("workbook") or DEFAULT_WORKBOOK))
            LOG_DIR.mkdir(exist_ok=True)
            with STATE_LOCK:
                if STATE.log_path is None:
                    STATE.log_path = LOG_DIR / f"ecoking-web-{datetime.now().strftime('%Y%m%d-%H%M%S')}.log"
            opened = open_local_file(workbook_path)
            if opened:
                self.send_json({"ok": True})
            else:
                self.send_json({"error": "Nije moguće automatski otvoriti Excel. Provjeri log za detalje."}, status=500)
            return

        if parsed.path == "/api/browse-workbook":
            length = int(self.headers.get("Content-Length", "0") or 0)
            body = self.rfile.read(length).decode("utf-8") if length else "{}"
            config = json.loads(body or "{}")
            workbook_path = Path(str(config.get("workbook") or DEFAULT_WORKBOOK))
            selected, error = choose_workbook_with_native_dialog(workbook_path)
            if selected:
                self.send_json({"path": str(selected)})
            else:
                self.send_json({"error": error or "Fajl nije izabran."}, status=503)
            return

        self.send_error(404)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the EcoKing local web launcher.")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    server = ThreadingHTTPServer((args.host, args.port), Handler)
    print(f"EcoKing web launcher: http://{args.host}:{args.port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
