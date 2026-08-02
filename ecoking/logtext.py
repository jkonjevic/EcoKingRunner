"""Turn the scraper's English log lines into Serbian ijekavica for the UI.

Both launchers show the same text, so the rules live here rather than in each.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Callable

ANSI_RE = re.compile(r"\x1b\[[0-9;]*m")
_HEADER_RE = re.compile(r"^(\d{2}:\d{2}:\d{2})\s+([A-Z]+)\s+(.*)$")
_STATION_RE = re.compile(r"^\[(\d+)/(\d+)\] Station=(.+?), Excel row=(.+?), device=(.+)$")

LEVELS = {
    "DEBUG": "DETALJ",
    "INFO": "INFO",
    "WARNING": "UPOZORENJE",
    "ERROR": "GREŠKA",
    "CRITICAL": "KRITIČNO",
}

#: Log level -> the severity the UI colours and filters by.
SEVERITIES = {
    "DEBUG": "info",
    "INFO": "info",
    "WARNING": "warning",
    "ERROR": "error",
    "CRITICAL": "error",
}

#: Names the scraper gives the page elements it clicks and fills. They appear
#: inside otherwise translated debug lines, so they are translated too.
ELEMENT_LABELS = {
    "email": "e-mail",
    "password": "lozinku",
    "login next/submit": "dugme za nastavak prijave",
    "login submit": "dugme za prijavu",
    "device search": "polje za pretragu uređaja",
    "configured device dropdown": "podešeni padajući meni uređaja",
    "interval dropdown": "padajući meni intervala",
    "calendar OK button": "dugme OK u kalendaru",
    "website date picker": "izbor datuma na sajtu",
}


def _element_label(text: str) -> str:
    """Translate an element name, keeping any trailing note in brackets."""
    stripped = text.strip()
    note = ""
    bracket = stripped.find(" (")
    if bracket > 0 and stripped.endswith(")"):
        stripped, note = stripped[:bracket], " " + stripped[bracket + 1 :]
    return ELEMENT_LABELS.get(stripped, stripped) + note


REPLACEMENTS: list[tuple[re.Pattern[str], str | Callable[[re.Match[str]], str]]] = [
    (re.compile(r"^Clicking (.+?) with selector (.+)$"), lambda m: f"Klik na {_element_label(m.group(1))} preko selektora {m.group(2)}"),
    (re.compile(r"^Filling (.+?) with selector (.+)$"), lambda m: f"Unosim {_element_label(m.group(1))} preko selektora {m.group(2)}"),
    (re.compile(r"^REPORT GENERATED SUCCESSFULLY: (.+) for (.+) with (\d+) mapped rows$"), r"USPJEŠNO KREIRAN IZVJEŠTAJ: \1 za datum \2 sa \3 mapiranih redova"),
    (re.compile(r"^Loaded (\d+) workbook rows from (.+)$"), r"Učitano je \1 redova iz Excel template-a \2"),
    (re.compile(r"^Loaded (\d+) stations from (.+)$"), r"Učitano je \1 stanica iz \2"),
    (re.compile(r"^Built (\d+) station jobs from (\d+) stations$"), r"Pripremljeno je \1 zadataka od \2 stanica"),
    (re.compile(r"^Station (.+) is disabled; skipping\.$"), r"Stanica \1 je isključena i preskače se."),
    (re.compile(r"^Station (.+) has no device name; skipping\.$"), r"Stanica \1 nema naziv uređaja i preskače se."),
    (re.compile(r"^No template row for LOKACIJA=(.+) VODOMJER=(.+); skipping (.+)\.$"), r"Nema reda u template-u za LOKACIJA=\1, VODOMJER=\2; preskačem \3."),
    (re.compile(r"^Launching browser: headless=(\w+) slow_mo_ms=(\d+)$"), r"Pokrećem browser: skriven=\1, usporenje=\2 ms"),
    (re.compile(r"^Running (\d+) jobs across (\d+) browser workers\.$"), r"Pokrećem \1 zadataka kroz \2 paralelna browser procesa."),
    (re.compile(r"^Opening (.+)$"), r"Otvaram stranicu \1"),
    (re.compile(r"^Login flow submitted$"), "Prijava je poslata"),
    (re.compile(r"^Searching device: (.+)$"), r"Tražim uređaj: \1"),
    (re.compile(r"^Clicked device dropdown using precise top-left selector$"), "Otvoren je padajući meni za izbor uređaja"),
    (re.compile(r"^Filling location search with visible dropdown input (.+)$"), r"Unosim vrijednost u polje pretrage (\1)"),
    (re.compile(r"^Retyped location search with keyboard into (.+)$"), r"Ponovo unosim pretragu preko tastature (\1)"),
    (re.compile(r"^No dropdown results for (.+) after DOM fill\. Retrying with keyboard typing\.$"), r"Nema rezultata za \1 nakon prvog unosa. Pokušavam ponovo unosom preko tastature."),
    (re.compile(r"^Dropdown options for (.+): (.+)$"), r"Rezultati u padajućem meniju za \1: \2"),
    (re.compile(r"^Choosing device: (.+)$"), r"Biranje uređaja: \1"),
    (re.compile(r"^Selected (.+) using shorter query (.+)$"), r"Uređaj \1 je izabran kraćom pretragom \2"),
    (re.compile(r"^Search query (.+) failed for (.+)\. Trying a shorter query\.$"), r"Pretraga \1 nije uspjela za \2. Pokušavam kraću pretragu."),
    (re.compile(r"^Device name (.+) matches (\d+) devices: (.+)\. Make the name in the station list more specific\.$"), r"Naziv \1 odgovara za \2 uređaja: \3. Precizirajte naziv u listi stanica."),
    (re.compile(r"^No dropdown result matches (.+)\. Visible: (.+)$"), r"Nijedan rezultat ne odgovara nazivu \1. Vidljivo je: \2"),
    (re.compile(r"^Selecting interval: (.+)$"), r"Biranje intervala: \1"),
    (re.compile(r"^Clicked interval option '(.+)'$"), r"Kliknut je interval '\1'"),
    (re.compile(r"^Clicking interval button with selector (.+)$"), r"Klik na dugme intervala preko selektora \1"),
    (re.compile(r"^Read (.+): daily=(.+) m3, max=(.+) m3, min=(.+) m3$"), r"Očitano za \1: dnevno=\2 m3, maksimum=\3 m3, minimum=\4 m3"),
    (re.compile(r'^FOUND: MIN: (.+) MAX: (.+) DAILY: (.+) BATTERY: (.+) for "(.+)"$'), r'PRONAĐENO: MIN=\1, MAX=\2, DNEVNO=\3, BATERIJA=\4 za "\5"'),
    (re.compile(r'^FOUND: MIN: (.+) MAX: (.+) DAILY: (.+) for "(.+)"$'), r'PRONAĐENO: MIN=\1, MAX=\2, DNEVNO=\3 za "\4"'),
    (re.compile(r"^Failed to scrape station (.+) for Excel row (.+)$"), r"Neuspjelo očitavanje stanice \1 za Excel red \2"),
    (re.compile(r"^Skipping station (.+) \(Excel row (.+)\): (.+)$"), r"Preskačem stanicu \1 (Excel red \2): \3"),
    (re.compile(r"^Saved debug artifacts: (.+)\.png and (.+)\.html$"), r"Sačuvani su debug fajlovi: \1.png i \2.html"),
    (re.compile(r"^Generated report (.+) for (.+) with (\d+) mapped rows\.$"), r"Kreiran je izvještaj \1 za datum \2 sa \3 mapiranih redova."),
    (re.compile(r"^Could not place (\d+) scraped station\(s\): (.+)$"), r"Nije moguće smjestiti \1 očitanih stanica: \2"),
    (re.compile(r"^No station matched a template row\..*$"), "Nijedna stanica ne odgovara redu u template-u. Otvorite listu stanica i ispravite unose."),
    (re.compile(r"^========== EXECUTION REPORT ==========$"), "========== IZVJEŠTAJ IZVRŠENJA =========="),
    (re.compile(r"^SUCCESSFUL: (\d+)$"), r"USPJEŠNO: \1"),
    (re.compile(r"^NO DATA / NO ENTRIES: (\d+)$"), r"BEZ PODATAKA / BEZ UNOSA: \1"),
    (re.compile(r"^FAILED: (\d+)$"), r"NEUSPJELO: \1"),
    # FAIL / NO DATA / TELEMETRY FAIL are rebuilt in translate(), which also
    # translates the reason they carry.
    (re.compile(r"^OK \| row=(.+?) \| (.+)$"), r"OK | red=\1 | \2"),
    (re.compile(r"^Run failed: (.+)$"), r"Pokretanje nije uspjelo: \1"),
    (re.compile(r"^Selected date=(.+?); interval target=(.+?); total target=(.+)$"), r"Izabrani datum: \1; ciljni datum intervala: \2; ciljni datum ukupnog: \3"),
    (re.compile(r"^Read battery level: (.+)$"), r"Očitana baterija: \1"),
    (re.compile(r"^Battery level component was not found for the selected station\.$"), "Nije pronađen prikaz baterije za izabranu stanicu."),
    (re.compile(r"^Clicked device dropdown using top-left content heuristic$"), "Padajući meni za uređaje je otvoren prepoznavanjem sadržaja gore lijevo"),
    (re.compile(r"^Typed location search into focused (.+?) element$"), r"Pretraga je unesena preko tastature u aktivno polje \1"),
    (re.compile(r"^DATE APPLY \| metric=(.+?) \| website date=(.+)$"), r"POSTAVLJANJE DATUMA | mjerenje=\1 | datum na sajtu=\2"),
    (re.compile(r"^30-day last bar \| UI date=(.+?) \| website date=(.+?) \| Y-axis value=(.+?) m3$"), r"Posljednja kolona 30-dnevnog grafa | datum u aplikaciji=\1 | datum na sajtu=\2 | vrijednost=\3 m3"),
    (re.compile(r"^Selected website date (.+)$"), r"Izabran je datum na sajtu \1"),
    (re.compile(r"^Entered website date (.+?) using WEBSITE_DATE_INPUT_SELECTOR$"), r"Datum \1 je unesen preko WEBSITE_DATE_INPUT_SELECTOR"),
    (re.compile(r"^Entered inline website date (.+?) in the top-right date control$"), r"Datum \1 je unesen u polje gore desno"),
    (re.compile(r"^Reached inline website date (.+?) using toolbar navigation$"), r"Do datuma \1 se stiglo strelicama u alatnoj traci"),
    (re.compile(r"^Could not click WEBSITE_DATE_PICKER_SELECTOR; trying the inline top-right date field\.$"), "Klik na WEBSITE_DATE_PICKER_SELECTOR nije uspio; pokušavam preko polja gore desno."),
    (re.compile(r"^Date picker tooltip button was not directly identified; trying the inline top-right date field\.$"), "Dugme kalendara nije prepoznato; pokušavam preko polja gore desno."),
    (re.compile(r"^Login fields were not detected\. Waiting (\d+) seconds for manual login\.$"), r"Polja za prijavu nisu prepoznata. Čekam \1 sekundi da se prijavite ručno."),
    (re.compile(r"^Running (\d+) headed browsers in parallel\..*$"), r"Paralelno se pokreće \1 vidljivih browsera. Za brži i tiši rad koristite skriveni browser."),
    (re.compile(r"^--keep-browser-open is ignored for parallel runs\.$"), "Opcija „ostavi browser otvoren“ se zanemaruje kod paralelnog rada."),
    (re.compile(r"^KEEP_BROWSER_OPEN is enabled\..*$"), "Browser ostaje otvoren. Pritisnite Ctrl+C u terminalu kada završite pregled."),
    (re.compile(r"^Browser worker failed for (\d+) job\(s\)\.$"), r"Browser proces nije uspio za \1 zadataka."),
    (re.compile(r"^Traceback for station (.+?) \(Excel row (.+?)\)$"), r"Detalji greške za stanicu \1 (Excel red \2)"),
    (re.compile(r"^Only the telemetry stage was requested; filling (.+)\.$"), r"Zatražena je samo telemetrija; dopunjavam \1."),
    (re.compile(r"^Desktop folder was not found; keeping report at (.+)\.$"), r"Desktop folder nije pronađen; izvještaj ostaje na \1."),
    (re.compile(r"^Workbook calculation flags are not available in this openpyxl version\.$"), "Ova verzija openpyxl-a ne podržava podešavanje ponovnog računanja formula."),
    # Self-update, shown before the UI is even up.
    (re.compile(r"^Could not check for updates \((.+)\)\. Continuing with the current version\.$"), r"Provjera ažuriranja nije uspjela (\1). Nastavljam sa trenutnom verzijom."),
    (re.compile(r"^Updated application code from (.+)\.$"), r"Aplikacija je ažurirana sa \1."),
    # Second pass: 17h reservoir levels from the telemetry site.
    (re.compile(r"^Telemetry: reading (\d+):00 levels for (.+) across (\d+) location\(s\)$"), r"Telemetrija: očitavam nivoe u \1:00 za datum \2 na \3 lokacija"),
    (re.compile(r"^Telemetry: opening (.+)$"), r"Telemetrija: otvaram stranicu \1"),
    (re.compile(r"^Telemetry: already signed in$"), "Telemetrija: prijava već postoji"),
    (re.compile(r"^Telemetry: signed in as (.+)$"), r"Telemetrija: prijavljen kao \1"),
    (re.compile(r"^Telemetry \[(\d+)/(\d+)\] Location=(.+)$"), r"Telemetrija [\1/\2] Lokacija=\3"),
    (re.compile(r"^Telemetry: LEVEL (.+) m at (.+) for (.+)$"), r"Telemetrija: NIVO \1 m u \2 za \3"),
    (re.compile(r"^Telemetry: wrote (.+) m into (.+) \(rows (.+)\)$"), r"Telemetrija: upisano \1 m za \2 (redovi \3)"),
    (re.compile(r"^Telemetry: could not read (.+?): (.+)$"), r"Telemetrija: neuspjelo očitavanje za \1: \2"),
    (re.compile(r"^Telemetry: (.+) is not in the site's location table; skipping\.$"), r"Telemetrija: \1 ne postoji u tabeli lokacija na sajtu; preskačem."),
    (re.compile(r"^Telemetry: (.+) has no entry in (.+); its level is not written\.$"), r"Telemetrija: \1 nema unos u \2; nivo nije upisan."),
    (re.compile(r"^Telemetry: no report row with LOKACIJA=(.+) for (.+)\.$"), r"Telemetrija: nema reda LOKACIJA=\1 za \2."),
    (re.compile(r"^Telemetry: could not write the levels into (.+?): (.+)$"), r"Telemetrija: nije moguće upisati nivoe u \1: \2"),
    (re.compile(r"^Telemetry stage is disabled .*$"), "Telemetrija je isključena; preskačem."),
    (re.compile(r"^Telemetry stage skipped: (.+)$"), r"Telemetrija je preskočena: \1"),
    (re.compile(r"^Telemetry stage failed: (.+)\. The report is kept without the 17h levels\.$"), r"Telemetrija nije uspjela: \1. Izvještaj ostaje bez nivoa u 17h."),
    (re.compile(r"^TELEMETRY DONE: (\d+) of (\d+) location\(s\) read, (\d+) report row\(s\) filled$"), r"TELEMETRIJA ZAVRŠENA: očitano \1 od \2 lokacija, popunjeno \3 redova"),
]


#: Why a station or a level did not make it into the report. These come from
#: exception text, so unlike the lines above they nest -- "could not select
#: device X, attempts: <this query>: <that reason>" -- and are applied
#: everywhere in the string rather than anchored to its start. The scraper
#: keeps them in English on purpose: it branches on the wording to tell a
#: failure from missing data, so translating happens here, on the way out.
REASON_REPLACEMENTS: list[tuple[re.Pattern[str], str]] = [
    (re.compile(r"Could not select device (.+?)\. Attempts: "), r"Nije moguće izabrati uređaj \1. Pokušaji: "),
    (re.compile(r"No dropdown result matches (.+?)\. Visible: "), r"Nijedan rezultat ne odgovara nazivu \1. Vidljivo je: "),
    (re.compile(r"No dropdown results after searching (.+?)\."), r"Nema rezultata pretrage za \1."),
    (re.compile(r"Device name (.+?) matches (\d+) devices: "), r"Naziv \1 odgovara za \2 uređaja: "),
    (re.compile(r"Make the name in the station list more specific\."), "Precizirajte naziv u listi stanica."),
    (re.compile(r"No numeric 1-day chart values found for (.+?)\."), r"Nema brojčanih vrijednosti na dnevnom grafu za \1."),
    (re.compile(r"No latest 30-day bar found for UI-selected date (.+?) for (.+?)\."), r"Nema posljednje kolone 30-dnevnog grafa za datum \1 za \2."),
    (re.compile(r"Could not choose interval (.+?) using INTERVAL_SELECTOR\."), r"Nije moguće izabrati interval \1 preko INTERVAL_SELECTOR."),
    (re.compile(r"Could not choose interval (.+?)\."), r"Nije moguće izabrati interval \1."),
    (re.compile(r"Could not find interval selector\. Set INTERVAL_SELECTOR\."), "Nije pronađen izbor intervala. Podesite INTERVAL_SELECTOR."),
    (re.compile(r"Could not find a selectable cell for website date (.+?) in the displayed month \((\d+) day cell\(s\) scanned\)\. Set WEBSITE_DATE_DAY_SELECTOR\."), r"Nije pronađen dan \1 u prikazanom mjesecu (pregledano \2 polja). Podesite WEBSITE_DATE_DAY_SELECTOR."),
    (re.compile(r"Could not select website date (.+?)\. Set WEBSITE_DATE_DAY_SELECTOR\."), r"Nije moguće izabrati datum \1 na sajtu. Podesite WEBSITE_DATE_DAY_SELECTOR."),
    (re.compile(r"Could not confirm the date picker selection\. Set WEBSITE_DATE_OK_SELECTOR\."), "Nije moguće potvrditi izbor datuma. Podesite WEBSITE_DATE_OK_SELECTOR."),
    (re.compile(r"Could not navigate to the requested calendar month\. Set (.+?)\."), r"Nije moguće doći do traženog mjeseca u kalendaru. Podesite \1."),
    (re.compile(r"Date picker dialog did not open and no top-right DD/MM/YYYY date field was found\."), "Kalendar se nije otvorio, a polje za datum gore desno nije pronađeno."),
    (re.compile(r"Could not complete login\. Add LOGIN_\* selectors or set WAIT_FOR_LOGIN_SECONDS\."), "Prijava nije uspjela. Dodajte LOGIN_* selektore ili podesite WAIT_FOR_LOGIN_SECONDS."),
    (re.compile(r"Could not find the blue dropdown/search button\. Set SEARCH_TOGGLE_SELECTOR\."), "Nije pronađeno plavo dugme za pretragu. Podesite SEARCH_TOGGLE_SELECTOR."),
    (re.compile(r"Could not find the opened dropdown search input\. Set SEARCH_INPUT_SELECTOR\."), "Nije pronađeno polje za pretragu u otvorenom meniju. Podesite SEARCH_INPUT_SELECTOR."),
    (re.compile(r"Output workbook must be different from (.+?)\."), r"Izlazni fajl mora biti različit od \1."),
    (re.compile(r"Template is missing LOKACIJA and VODOMJER headers\."), "Template nema zaglavlja LOKACIJA i VODOMJER."),
    (re.compile(r"Report is missing LOKACIJA and VODOMJER headers\."), "Izvještaj nema zaglavlja LOKACIJA i VODOMJER."),
    (re.compile(r"\.env must define url, email/gmail, and password\."), "U .env fajlu moraju biti podešeni url, email/gmail i lozinka."),
    (re.compile(r"No station matched a template row\..*"), "Nijedna stanica ne odgovara redu u template-u. Otvorite listu stanica i ispravite unose."),
    (re.compile(r"Selected date must use YYYY-MM-DD or DD/MM/YYYY format\."), "Datum mora biti u formatu YYYY-MM-DD ili DD/MM/YYYY."),
    (re.compile(r"Selected date cannot be in the future\."), "Datum ne može biti u budućnosti."),
    # Playwright's own wording, which surfaces verbatim as a failure reason.
    (re.compile(r"Timeout (\d+)ms exceeded\.?"), r"Isteklo je vrijeme čekanja (\1 ms)."),
    (re.compile(r"\bwaiting for (locator|selector)\b"), r"čekanje na \1"),
    (re.compile(r"Element is not visible"), "Element nije vidljiv"),
    (re.compile(r"Target (page|frame|browser) has been closed"), r"\1 je zatvoren(a)"),
]


def translate_reason(text: str) -> str:
    """Translate a failure reason, including the ones nested inside it."""
    result = str(text or "")
    for pattern, replacement in REASON_REPLACEMENTS:
        result = pattern.sub(replacement, result)
    return result


#: The roll-up lines ``log_run_report`` and the telemetry stage end a run with.
#: They are the only place a failure is named next to its reason, so they are
#: what the "Problemi" panel is built from.
_FAILURE_RES: list[tuple[str, re.Pattern[str]]] = [
    ("station", re.compile(r"^FAIL \| row=(?P<row>.*?) \| (?P<label>.*?) \| (?P<reason>.+)$")),
    ("no-data", re.compile(r"^NO DATA \| row=(?P<row>.*?) \| (?P<label>.*?) \| (?P<reason>.+)$")),
    ("telemetry", re.compile(r"^TELEMETRY FAIL \| (?P<label>.*?) \| (?P<reason>.+)$")),
]


@dataclass(frozen=True)
class Failure:
    """One named thing that did not make it into the report, and why."""

    kind: str  # "station" | "no-data" | "telemetry"
    label: str
    reason: str
    row: str = ""

    @property
    def severity(self) -> str:
        # A station that failed outright leaves an empty row; no data and a
        # missing level are gaps in an otherwise finished report.
        return "error" if self.kind == "station" else "warning"

    def to_json(self) -> dict[str, str]:
        return {
            "kind": self.kind,
            "label": self.label,
            "reason": self.reason,
            "row": self.row,
            "severity": self.severity,
        }


def strip_ansi(text: str) -> str:
    return ANSI_RE.sub("", text).rstrip("\r\n")


def _split_header(line: str) -> tuple[str, str]:
    """``("WARNING", "body")`` for a logged line, ``("", line)`` for raw output."""
    header = _HEADER_RE.match(strip_ansi(line))
    if not header:
        return "", strip_ansi(line)
    _, level, body = header.groups()
    return level, body


def classify(line: str) -> str:
    """The severity a log line carries: ``error``, ``warning`` or ``info``."""
    level, _ = _split_header(line)
    return SEVERITIES.get(level, "info")


def parse_failure(line: str) -> Failure | None:
    """Pull the device/location and reason out of a run report's failure line."""
    _, body = _split_header(line)
    body = body.strip()
    for kind, pattern in _FAILURE_RES:
        found = pattern.match(body)
        if found:
            fields = found.groupdict()
            return Failure(
                kind=kind,
                label=fields["label"].strip(),
                reason=translate_reason(fields["reason"].strip()),
                row=fields.get("row", "").strip(),
            )
    return None


def translate(line: str) -> str:
    """Translate one log line, keeping its ``HH:MM:SS LEVEL`` prefix."""
    clean = strip_ansi(line)
    if not clean:
        return ""

    prefix = ""
    body = clean
    header = _HEADER_RE.match(clean)
    if header:
        time_part, level, body = header.groups()
        prefix = f"{time_part} {LEVELS.get(level, level):<10} "

    station = _STATION_RE.match(body)
    if station:
        index, total, label, row, device = station.groups()
        return f"{prefix}[{index}/{total}] Stanica={label}, Excel red={row}, uređaj={device}"

    # The roll-up lines carry a reason that is itself English prose, so they
    # are rebuilt here instead of going through a plain regex template.
    failure = parse_failure(body)
    if failure:
        heading = {"station": "NEUSPJEH", "no-data": "BEZ PODATAKA", "telemetry": "TELEMETRIJA NEUSPJEH"}
        where = f" | red={failure.row}" if failure.row else ""
        return f"{prefix}{heading[failure.kind]}{where} | {failure.label} | razlog: {failure.reason}"

    for pattern, replacement in REPLACEMENTS:
        if pattern.search(body):
            return prefix + pattern.sub(replacement, body)
    return prefix + body
