"""Turn the scraper's English log lines into Serbian ijekavica for the UI.

Both launchers show the same text, so the rules live here rather than in each.
"""

from __future__ import annotations

import re

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

REPLACEMENTS: list[tuple[re.Pattern[str], str]] = [
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
    (re.compile(r"^FAIL \| row=(.+?) \| (.+?) \| (.+)$"), r"NEUSPJEH | red=\1 | \2 | razlog: \3"),
    (re.compile(r"^NO DATA \| row=(.+?) \| (.+?) \| (.+)$"), r"BEZ PODATAKA | red=\1 | \2 | razlog: \3"),
    (re.compile(r"^OK \| row=(.+?) \| (.+)$"), r"OK | red=\1 | \2"),
    (re.compile(r"^Run failed: (.+)$"), r"Pokretanje nije uspjelo: \1"),
]


def strip_ansi(text: str) -> str:
    return ANSI_RE.sub("", text).rstrip("\r\n")


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

    for pattern, replacement in REPLACEMENTS:
        if pattern.search(body):
            return prefix + pattern.sub(replacement, body)
    return prefix + body
