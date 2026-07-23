from __future__ import annotations

import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

# 1. Import load_dotenv
from dotenv import load_dotenv

APP_TITLE = "EcoKing Dnevni Obračun"


def app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


ROOT = app_root()

# 2. Explicitly load the .env file next to the executable / script
load_dotenv(dotenv_path=ROOT / ".env")

# Now your app can safely access environment variables
DEFAULT_WORKBOOK = ROOT / "EcoKing - tabela potrošnje - Jul 2026. TEST.xlsx"
DEFAULT_STATIONS = ROOT / "herceg_novi_stations.json"
LOG_DIR = ROOT / "logs"
ANSI_RE = re.compile(r"\x1b\[[0-9;]*m")


def run_scraper_from_frozen_exe() -> int:
    args = [arg for arg in sys.argv[1:] if arg != "--run-scraper"]
    sys.argv = ["ecoking_daily.py", *args]
    from ecoking_daily import main as scraper_main

    return scraper_main()


if "--run-scraper" in sys.argv:
    raise SystemExit(run_scraper_from_frozen_exe())


try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ModuleNotFoundError as exc:
    raise SystemExit(
        "Tkinter nije instaliran u ovom Python okruženju. "
        "Na Windowsu instalirajte standardni Python sa Tcl/Tk podrškom."
    ) from exc


class ToolTip:
    """Creates a tooltip popup when hovering over a Tkinter widget."""

    def __init__(self, widget: tk.Widget, text: str) -> None:
        self.widget = widget
        self.text = text
        self.tip_window: tk.Toplevel | None = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event: tk.Event | None = None) -> None:
        if self.tip_window or not self.text:
            return
        x = self.widget.winfo_rootx() + 15
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw,
            text=self.text,
            justify=tk.LEFT,
            background="#0f172a",
            foreground="#f8fafc",
            relief=tk.SOLID,
            borderwidth=1,
            font=("Segoe UI", 9),
            padx=8,
            pady=5,
        )
        label.pack()

    def hide_tip(self, event: tk.Event | None = None) -> None:
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None


@dataclass
class LauncherState:
    process: subprocess.Popen[str] | None = None
    log_path: Path | None = None
    running: bool = False


def strip_ansi(text: str) -> str:
    return ANSI_RE.sub("", text).rstrip("\r\n")


def short_path(path: str | Path) -> str:
    try:
        return str(Path(path).resolve().relative_to(ROOT))
    except Exception:
        return str(path)


def sanitize_path(raw_path: str | Path) -> Path:
    clean_str = str(raw_path).strip().strip('"').strip("'")
    return Path(clean_str).resolve()


def open_path(path: Path | str) -> tuple[bool, str]:
    target_path = sanitize_path(path)
    if not target_path.exists():
        return False, f"Putanja ne postoji: {target_path}"

    # Windows multi-stage fallback launcher
    if sys.platform.startswith("win"):
        # 1. Standard Windows Shell Launch
        try:
            os.startfile(str(target_path))  # type: ignore[attr-defined]
            return True, f"Otvaram Excel: {target_path.name}"
        except Exception:
            pass

        # 2. Command Prompt Launcher
        try:
            subprocess.Popen(["cmd.exe", "/c", "start", "", str(target_path)], shell=True)
            return True, f"Otvaram Excel (via CMD): {target_path.name}"
        except Exception:
            pass

        # 3. PowerShell Process Starter
        try:
            subprocess.Popen(["powershell.exe", "-Command", f'Start-Process "{target_path}"'])
            return True, f"Otvaram Excel (via PowerShell): {target_path.name}"
        except Exception as exc:
            return False, f"Nije moguće otvoriti fajl: {exc}"

    # macOS / Linux
    candidates = [["open", str(target_path)]] if sys.platform == "darwin" else [
        ["gio", "open", str(target_path)],
        ["xdg-open", str(target_path)],
        ["libreoffice", "--calc", str(target_path)],
        ["localc", str(target_path)],
        ["soffice", "--calc", str(target_path)],
    ]
    errors: list[str] = []
    for command in candidates:
        if not shutil.which(command[0]):
            errors.append(f"{command[0]} nije instaliran")
            continue
        try:
            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            try:
                stdout, stderr = process.communicate(timeout=2)
            except subprocess.TimeoutExpired:
                return True, f"Otvaram preko: {' '.join(command[:2])}"
            if process.returncode == 0:
                return True, f"Otvaram preko: {' '.join(command[:2])}"
            errors.append((stderr or stdout or f"kod {process.returncode}").strip())
        except Exception as exc:
            errors.append(str(exc))
    return False, "Nije moguće automatski otvoriti fajl. " + " | ".join(errors[:3])


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
        (re.compile(r"^Clicked device dropdown using top-left content heuristic$"), "Otvoren je padajući meni za izbor uređaja"),
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
        (re.compile(r"^Updated (.+) for (.+) with (\d+) row\(s\)\.$"), r"Ažuriran je \1 za datum \2 sa \3 redova."),
        (re.compile(r"^Created sheet (.+) in (.+) with (\d+) scraped rows$"), r"Kreiran je list \1 u fajlu \2 sa \3 očitanih redova"),
        (re.compile(r"^========== EXECUTION REPORT ==========$"), "========== IZVJEŠTAJ IZVRŠENJA =========="),
        (re.compile(r"^SUCCESSFUL: (\d+)$"), r"USPJEŠNO: \1"),
        (re.compile(r"^NO DATA / NO ENTRIES: (\d+)$"), r"BEZ PODATAKA / BEZ UNOSA: \1"),
        (re.compile(r"^FAILED: (\d+)$"), r"NEUSPJELO: \1"),
        (re.compile(r"^======================================$"), "======================================"),
        (re.compile(r"^Run failed: (.+)$"), r"Pokretanje nije uspjelo: \1"),
    ]

    station_match = re.match(
        r"^\[(\d+)/(\d+)\] Station key=(.+), Excel row=(.+), Excel LOKACIJA=(.+), VODOMJER=(.+), browser search=(.+)$",
        body,
    )
    if station_match:
        idx, total, key, row, location, meter, search = station_match.groups()
        return f"{prefix}[{idx}/{total}] Stanica={key}, Excel red={row}, Lokacija={location}, Vodomjer={meter}, pretraga={search}"

    for pattern, replacement in replacements:
        if pattern.search(body):
            return prefix + pattern.sub(replacement, body)

    return prefix + body


class EcoKingLauncher(ttk.Frame):
    def __init__(self, master: tk.Tk) -> None:
        super().__init__(master, style="Main.TFrame")
        self.master = master
        self.state = LauncherState()
        self.log_queue: queue.Queue[str | tuple[str, int]] = queue.Queue()

        self.workbook_var = tk.StringVar(value=str(DEFAULT_WORKBOOK))
        self.browser_visible_var = tk.BooleanVar(value=True)
        self.verbose_var = tk.BooleanVar(value=True)
        self.keep_open_var = tk.BooleanVar(value=False)
        self.open_after_var = tk.BooleanVar(value=True)
        self.limit_var = tk.StringVar(value="")
        self.workers_var = tk.IntVar(value=1)
        self.slowmo_var = tk.IntVar(value=0)
        self.chart_wait_var = tk.IntVar(value=5000)
        self.search_wait_var = tk.IntVar(value=2000)

        self.status_var = tk.StringVar(value="Spremno za rad")
        self.command_var = tk.StringVar(value="")

        self._apply_styles()
        self._build_ui()
        self._refresh_command_preview()
        self.after(100, self._drain_log_queue)

    def _apply_styles(self) -> None:
        self.master.title(APP_TITLE)
        self.master.geometry("1150x760")
        self.master.minsize(1000, 680)

        try:
            self.master.state("zoomed")
        except Exception:
            self.master.attributes("-fullscreen", True)

        self.master.configure(bg="#f8fafc")
        self.master.protocol("WM_DELETE_WINDOW", self._on_close)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        BG_MAIN = "#f8fafc"
        CARD_BG = "#ffffff"
        PRIMARY = "#0284c7"
        PRIMARY_HOVER = "#0369a1"
        DANGER = "#e11d48"
        TEXT_MAIN = "#0f172a"
        TEXT_MUTED = "#64748b"

        style.configure(".", background=BG_MAIN, font=("Segoe UI", 10), foreground=TEXT_MAIN)
        style.configure("Main.TFrame", background=BG_MAIN)
        style.configure("Header.TFrame", background="#0f172a")
        style.configure("HeaderTitle.TLabel", background="#0f172a", foreground="#ffffff", font=("Segoe UI", 16, "bold"))
        style.configure("HeaderSub.TLabel", background="#0f172a", foreground="#94a3b8", font=("Segoe UI", 9))
        style.configure("Card.TFrame", background=CARD_BG, relief="flat", borderwidth=1)
        style.configure("CardTitle.TLabel", background=CARD_BG, foreground=TEXT_MAIN, font=("Segoe UI", 11, "bold"))
        style.configure("CardSub.TLabel", background=CARD_BG, foreground=TEXT_MUTED, font=("Segoe UI", 8))
        style.configure("CardLabel.TLabel", background=CARD_BG, foreground=TEXT_MAIN, font=("Segoe UI", 10))
        style.configure("CardCheck.TCheckbutton", background=CARD_BG, font=("Segoe UI", 10))

        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), background=PRIMARY, foreground="white", borderwidth=0, padding=(14, 8))
        style.map("Primary.TButton", background=[("active", PRIMARY_HOVER), ("disabled", "#cbd5e1")])

        style.configure("Danger.TButton", font=("Segoe UI", 9, "bold"), background=DANGER, foreground="white", borderwidth=0, padding=(10, 6))
        style.map("Danger.TButton", background=[("active", "#be123c"), ("disabled", "#f1f5f9")])

        style.configure("Secondary.TButton", font=("Segoe UI", 9), padding=(8, 4))
        style.configure("Status.TLabel", background=BG_MAIN, foreground=PRIMARY, font=("Segoe UI", 10, "bold"))

    def _build_ui(self) -> None:
        self.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(self, style="Header.TFrame", padding=(20, 14))
        header.pack(fill=tk.X)
        ttk.Label(header, text="EcoKing Dnevni Obračun", style="HeaderTitle.TLabel").pack(anchor=tk.W)
        ttk.Label(
            header,
            text="Automatizovano očitavanje vodomjera i generisanje izvještaja za jučerašnji dan.",
            style="HeaderSub.TLabel",
        ).pack(anchor=tk.W, pady=(2, 0))

        content = ttk.Frame(self, padding=16, style="Main.TFrame")
        content.pack(fill=tk.BOTH, expand=True)
        content.columnconfigure(0, weight=0, minsize=440)
        content.columnconfigure(1, weight=1)
        content.rowconfigure(0, weight=1)

        left_panel = ttk.Frame(content, style="Main.TFrame")
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 14))

        right_panel = ttk.Frame(content, style="Main.TFrame")
        right_panel.grid(row=0, column=1, sticky="nsew")
        right_panel.rowconfigure(0, weight=1)
        right_panel.columnconfigure(0, weight=1)

        self._build_files_card(left_panel)
        self._build_options_card(left_panel)
        self._build_actions_card(left_panel)
        self._build_log_panel(right_panel)

    def _card(self, parent: ttk.Frame, title: str) -> ttk.Frame:
        container = ttk.Frame(parent, style="Main.TFrame", padding=(0, 0, 0, 12))
        container.pack(fill=tk.X)
        card = ttk.Frame(container, style="Card.TFrame", padding=14)
        card.pack(fill=tk.X)
        ttk.Label(card, text=title, style="CardTitle.TLabel").pack(anchor=tk.W, pady=(0, 8))
        return card

    def _build_files_card(self, parent: ttk.Frame) -> None:
        card = self._card(parent, "Excel Izvještaj")

        row = ttk.Frame(card, style="Card.TFrame")
        row.pack(fill=tk.X, pady=(0, 4))
        row.columnconfigure(0, weight=1)

        entry = ttk.Entry(row, textvariable=self.workbook_var, font=("Segoe UI", 9))
        entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ToolTip(entry, "Putanja do ciljnog Excel fajla u koji će se upisati novi dnevni podaci.")

        browse_btn = ttk.Button(row, text="Izaberi...", style="Secondary.TButton", command=lambda: self._browse_path(self.workbook_var))
        browse_btn.grid(row=0, column=1)
        ToolTip(browse_btn, "Otvara dijalog za izbor Excel fajla sa računara.")

        sub_label = ttk.Label(
            card,
            text="Podaci se upisuju u list za jučerašnji datum. Postojeći list za taj dan biće zamijenjen.",
            style="CardSub.TLabel",
            wraplength=380,
        )
        sub_label.pack(anchor=tk.W)
        ToolTip(sub_label, "Skripta automatski kreira ili prebrisava radni list sa imenom jučerašnjeg datuma.")

        self.workbook_var.trace_add("write", lambda *_: self._refresh_command_preview())

    def _build_options_card(self, parent: ttk.Frame) -> None:
        card = self._card(parent, "Postavke Pokretanja")

        checks = ttk.Frame(card, style="Card.TFrame")
        checks.pack(fill=tk.X, pady=(0, 10))

        chk1 = ttk.Checkbutton(checks, text="Prikaži browser tokom rada", variable=self.browser_visible_var, style="CardCheck.TCheckbutton", command=self._refresh_command_preview)
        chk1.grid(row=0, column=0, sticky=tk.W, pady=2)
        ToolTip(chk1, "Ako je označeno, otvara se vidljiv Chrome prozor tako da možete pratiti navigaciju uživo.")

        chk2 = ttk.Checkbutton(checks, text="Detaljan prikaz logova", variable=self.verbose_var, style="CardCheck.TCheckbutton", command=self._refresh_command_preview)
        chk2.grid(row=1, column=0, sticky=tk.W, pady=2)
        ToolTip(chk2, "Aktivira 'Verbose' režim rada koji u konzoli prikazuje sve tehničke korake i pretipkavanja.")

        chk3 = ttk.Checkbutton(checks, text="Ostavi browser otvoren po završetku", variable=self.keep_open_var, style="CardCheck.TCheckbutton", command=self._refresh_command_preview)
        chk3.grid(row=2, column=0, sticky=tk.W, pady=2)
        ToolTip(chk3, "Sprečava automatsko zatvaranje browsera nakon izvršenja rada (korisno za provjeru stanja).")

        chk4 = ttk.Checkbutton(checks, text="Automatski otvori Excel po završetku", variable=self.open_after_var, style="CardCheck.TCheckbutton", command=self._refresh_command_preview)
        chk4.grid(row=3, column=0, sticky=tk.W, pady=2)
        ToolTip(chk4, "Kada se očitavanje uspješno završi, automatski otvara ažurirani Excel fajl u sistemu.")

        grid = ttk.Frame(card, style="Card.TFrame")
        grid.pack(fill=tk.X)
        grid.columnconfigure(1, weight=1)
        grid.columnconfigure(3, weight=1)

        self._numeric_row(grid, 0, "Broj radnika", self.workers_var, 1, 8, 1, "Broj paralelnih browser procesa. Više radnika ubrzava rad, ali troši više memorije.")
        self._numeric_row(grid, 1, "Usporenje (ms)", self.slowmo_var, 0, 1000, 50, "Dodatna pauza u milisekundama između klikova/unosa (korisno ako je internet spor).")
        self._numeric_row(grid, 2, "Čekanje grafa (ms)", self.chart_wait_var, 1000, 60000, 1000, "Maksimalno vrijeme čekanja da se učita grafikon i očitaju m3 podaci.")
        self._numeric_row(grid, 3, "Čekanje pretrage (ms)", self.search_wait_var, 500, 15000, 500, "Vrijeme čekanja da padajući meni prikaže rezultate nakon unosa naziva stanice.")

        lbl_limit = ttk.Label(grid, text="Limit stanica", style="CardLabel.TLabel")
        lbl_limit.grid(row=4, column=0, sticky=tk.W, pady=4)
        limit_entry = ttk.Entry(grid, textvariable=self.limit_var, width=10, font=("Segoe UI", 9))
        limit_entry.grid(row=4, column=1, sticky=tk.W, padx=(6, 12), pady=4)

        limit_tip = "Ograničava obradu na prvih N stanica (npr. unesite '3' za brzi test bez čitanja svih 58 stanica)."
        ToolTip(lbl_limit, limit_tip)
        ToolTip(limit_entry, limit_tip)

        for var in [self.limit_var, self.workers_var, self.slowmo_var, self.chart_wait_var, self.search_wait_var]:
            var.trace_add("write", lambda *_: self._refresh_command_preview())

    def _numeric_row(self, parent: ttk.Frame, row: int, label_text: str, variable: tk.IntVar, from_: int, to: int, increment: int, tooltip_text: str) -> None:
        col = 0 if row < 2 else 2
        local_row = row if row < 2 else row - 2

        lbl = ttk.Label(parent, text=label_text, style="CardLabel.TLabel")
        lbl.grid(row=local_row, column=col, sticky=tk.W, pady=4)

        spin = ttk.Spinbox(parent, from_=from_, to=to, increment=increment, textvariable=variable, width=8, font=("Segoe UI", 9))
        spin.grid(row=local_row, column=col + 1, sticky=tk.W, padx=(6, 12), pady=4)

        ToolTip(lbl, tooltip_text)
        ToolTip(spin, tooltip_text)

    def _build_actions_card(self, parent: ttk.Frame) -> None:
        card = self._card(parent, "Upravljanje")

        self.run_button = ttk.Button(card, text="▶  Pokreni Obračun", style="Primary.TButton", command=self._start_run)
        self.run_button.pack(fill=tk.X, pady=(0, 6))
        ToolTip(self.run_button, "Pokreće automatski proces očitavanja vodomjera sa zadatim postavkama.")

        self.stop_button = ttk.Button(card, text="■ Zaustavi", style="Danger.TButton", command=self._stop_run, state=tk.DISABLED)
        self.stop_button.pack(fill=tk.X, pady=(0, 10))
        ToolTip(self.stop_button, "Prekida trenutni proces očitavanja i zatvara aktivne browsere.")

        row = ttk.Frame(card, style="Card.TFrame")
        row.pack(fill=tk.X)

        btn_excel = ttk.Button(row, text="Otvori Excel", style="Secondary.TButton", command=lambda: self._open_path_with_feedback(self.workbook_var.get()))
        btn_excel.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 3))
        ToolTip(btn_excel, "Ručno otvara izabrani Excel fajl u Vašem podrazumijevanom programu.")

        btn_logs = ttk.Button(row, text="Otvori Folder Logova", style="Secondary.TButton", command=lambda: self._open_path_with_feedback(LOG_DIR))
        btn_logs.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(3, 0))
        ToolTip(btn_logs, "Otvara Windows Explorer folder u kojem se čuvaju istorije svih pokretanja.")

        ttk.Label(card, textvariable=self.status_var, style="Status.TLabel", wraplength=380).pack(anchor=tk.W, pady=(10, 0))

    def _build_log_panel(self, parent: ttk.Frame) -> None:
        top = ttk.Frame(parent, style="Main.TFrame")
        top.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(top, text="Konzola & Logovi", font=("Segoe UI", 11, "bold"), background="#f8fafc").pack(side=tk.LEFT)
        btn_clear = ttk.Button(top, text="Očisti Konzolu", style="Secondary.TButton", command=self._clear_log)
        btn_clear.pack(side=tk.RIGHT)
        ToolTip(btn_clear, "Briše sav trenutni tekst prikazan u konzoli ispod.")

        log_frame = ttk.Frame(parent, style="Main.TFrame")
        log_frame.pack(fill=tk.BOTH, expand=True)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_frame,
            wrap=tk.WORD,
            font=("Consolas", 10),
            bg="#0f172a",
            fg="#f8fafc",
            insertbackground="#f8fafc",
            relief=tk.FLAT,
            padx=12,
            pady=12,
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.log_text.tag_configure("error", foreground="#f87171")
        self.log_text.tag_configure("warning", foreground="#fbbf24")
        self.log_text.tag_configure("success", foreground="#4ade80")
        self.log_text.tag_configure("normal", foreground="#e2e8f0")

        cmd_frame = ttk.Frame(parent, style="Main.TFrame", padding=(0, 6, 0, 0))
        cmd_frame.pack(fill=tk.X)
        lbl_cmd = ttk.Label(cmd_frame, text="Komanda:", font=("Segoe UI", 9, "bold"), foreground="#64748b", background="#f8fafc")
        lbl_cmd.pack(anchor=tk.W)

        cmd_preview = ttk.Label(cmd_frame, textvariable=self.command_var, font=("Consolas", 8), foreground="#475569", background="#f8fafc", wraplength=600)
        cmd_preview.pack(anchor=tk.W)
        ToolTip(cmd_preview, "Pregled tačne komande i parametara koji se prosljeđuju pozadinskoj skripti.")

    def _browse_path(self, variable: tk.StringVar) -> None:
        initial = sanitize_path(variable.get()).parent if variable.get() else ROOT
        selected = filedialog.askopenfilename(initialdir=str(initial), filetypes=[("Excel fajlovi", "*.xlsx;*.xls")])
        if selected:
            variable.set(selected)

    def _build_command(self) -> tuple[list[str], dict[str, str]]:
        if getattr(sys, "frozen", False):
            cmd = [sys.executable, "--run-scraper"]
        else:
            cmd = [sys.executable, str(ROOT / "ecoking_daily.py")]

        cmd.extend([
            "--workbook", str(sanitize_path(self.workbook_var.get())),
            "--workers", str(max(1, int(self.workers_var.get()))),
            "--slow-mo-ms", str(max(0, int(self.slowmo_var.get()))),
        ])
        cmd.append("--headed" if self.browser_visible_var.get() else "--headless")
        if self.verbose_var.get():
            cmd.append("--verbose")
        if self.keep_open_var.get():
            cmd.append("--keep-browser-open")
        if self.limit_var.get().strip():
            cmd.extend(["--limit", self.limit_var.get().strip()])

        env = os.environ.copy()
        env["NO_COLOR"] = "1"
        env["LOCATION_MAP_PATH"] = str(DEFAULT_STATIONS)
        env["CHART_WAIT_MS"] = str(max(1000, int(self.chart_wait_var.get())))
        env["SEARCH_RESULTS_WAIT_MS"] = str(max(500, int(self.search_wait_var.get())))
        env["PYTHONIOENCODING"] = "utf-8"
        return cmd, env

    def _refresh_command_preview(self) -> None:
        try:
            cmd, _ = self._build_command()
            display = " ".join(f'"{part}"' if " " in part else part for part in cmd)
            self.command_var.set(display)
        except Exception:
            self.command_var.set("Komanda će biti generisana nakon ispravne konfiguracije.")

    def _validate(self) -> bool:
        wb_path = sanitize_path(self.workbook_var.get())
        required_files = [
            ("Excel fajl", wb_path),
            ("Mapiranje stanica", DEFAULT_STATIONS),
        ]
        if not getattr(sys, "frozen", False):
            required_files.append(("Skripta", ROOT / "ecoking_daily.py"))
        missing = [f"{label}: {path}" for label, path in required_files if not path.exists()]
        if missing:
            messagebox.showerror("Nedostaju fajlovi", "Nije moguće pokrenuti obradu:\n\n" + "\n".join(missing))
            return False

        try:
            if self.limit_var.get().strip():
                limit = int(self.limit_var.get().strip())
                if limit < 1:
                    raise ValueError
            int(self.workers_var.get())
            int(self.slowmo_var.get())
            int(self.chart_wait_var.get())
            int(self.search_wait_var.get())
        except ValueError:
            messagebox.showerror("Neispravne opcije", "Provjerite unesene numeričke parametre (limit, radnici, milisekunde).")
            return False

        return True

    def _start_run(self) -> None:
        if self.state.running or not self._validate():
            return

        LOG_DIR.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        self.state.log_path = LOG_DIR / f"ecoking-{timestamp}.log"
        cmd, env = self._build_command()

        self._clear_log()
        self._append_log("Pokretanje obračuna...\n", "success")
        self._append_log(f"Excel fajl: {short_path(self.workbook_var.get())}\n", "normal")
        self._append_log(f"Log fajl: {short_path(self.state.log_path)}\n\n", "normal")

        self.state.running = True
        self.run_button.configure(state=tk.DISABLED)
        self.stop_button.configure(state=tk.NORMAL)
        self.status_var.set("Obračun je u toku...")

        thread = threading.Thread(target=self._run_process, args=(cmd, env), daemon=True)
        thread.start()

    def _run_process(self, cmd: list[str], env: dict[str, str]) -> None:
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
            self.state.process = process
            assert process.stdout is not None
            for raw_line in process.stdout:
                self.log_queue.put(raw_line)
            return_code = process.wait()
            self.log_queue.put(("__DONE__", return_code))
        except Exception as exc:
            self.log_queue.put(f"GREŠKA: Pokretanje procesa nije uspjelo: {exc}\n")
            self.log_queue.put(("__DONE__", 1))

    def _stop_run(self) -> None:
        process = self.state.process
        if not process or process.poll() is not None:
            return
        if not messagebox.askyesno("Zaustavljanje", "Da li želite prekiniti obračun koji je u toku?"):
            return
        self._append_log("Zaustavljanje procesa...\n", "warning")
        process.terminate()
        try:
            process.wait(timeout=8)
        except subprocess.TimeoutExpired:
            process.kill()

    def _drain_log_queue(self) -> None:
        try:
            while True:
                item = self.log_queue.get_nowait()
                if isinstance(item, tuple) and item[0] == "__DONE__":
                    self._finish_run(item[1])
                    continue
                assert isinstance(item, str)
                translated = translate_log_line(item)
                if not translated:
                    continue
                tag = self._log_tag(translated)
                self._append_log(translated + "\n", tag)
                self._write_log_file(translated + "\n")
        except queue.Empty:
            pass
        self.after(100, self._drain_log_queue)

    def _finish_run(self, return_code: int) -> None:
        self.state.running = False
        self.state.process = None
        self.run_button.configure(state=tk.NORMAL)
        self.stop_button.configure(state=tk.DISABLED)

        wb_path = sanitize_path(self.workbook_var.get())

        if return_code == 0:
            msg = "Obračun je uspješno završen."
            self.status_var.set(msg)
            self._append_log("\n" + msg + "\n", "success")
            self._write_log_file("\n" + msg + "\n")
            if self.open_after_var.get():
                opened, detail = open_path(wb_path)
                tag = "success" if opened else "warning"
                self._append_log(f"{detail}\n", tag)
                self._write_log_file(f"{detail}\n")
        else:
            msg = f"Obračun je završen sa greškom (kod: {return_code})."
            self.status_var.set(msg)
            self._append_log("\n" + msg + "\n", "error")
            self._write_log_file("\n" + msg + "\n")

    def _open_path_with_feedback(self, raw_path: str | Path) -> None:
        opened, detail = open_path(raw_path)
        if opened:
            self._append_log(detail + "\n", "success")
        else:
            messagebox.showwarning("Nije moguće otvoriti", detail)
            self._append_log(detail + "\n", "warning")

    def _append_log(self, text: str, tag: str = "normal") -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, text, tag)
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _write_log_file(self, text: str) -> None:
        if not self.state.log_path:
            return
        with self.state.log_path.open("a", encoding="utf-8") as handle:
            handle.write(text)

    def _log_tag(self, text: str) -> str:
        lowered = text.lower()
        if "greška" in lowered or "neuspjelo" in lowered or "nije uspjelo" in lowered:
            return "error"
        if "upozorenje" in lowered or "nema rezultata" in lowered or "bez podataka" in lowered:
            return "warning"
        if "uspješno" in lowered or "pronađeno" in lowered or "očitano" in lowered:
            return "success"
        return "normal"

    def _clear_log(self) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _on_close(self) -> None:
        if self.state.running:
            if not messagebox.askyesno("Zatvaranje", "Obračun je u toku. Da li želite zaustaviti proces i zatvoriti aplikaciju?"):
                return
            self._stop_run()
        self.master.destroy()


def main() -> None:
    root = tk.Tk()
    EcoKingLauncher(root)
    root.mainloop()


if __name__ == "__main__":
    main()