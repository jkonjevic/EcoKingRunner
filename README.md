# EcoKing Daily Automation

This project logs in to EcoKing, reads station measurements from the diagrams, and writes them into the selected Excel workbook.

The production flow is:

- user selects one Excel workbook
- station mapping is loaded from `herceg_novi_stations.json` in the app folder
- the script creates/replaces yesterday's sheet, for example `21.07.2026.`
- only that sheet is replaced; other workbook sheets are left untouched
- values are written directly back into the selected workbook

Written fields:

- `DNEVNA POTROŠNJA (m3)`
- `DNEVNA POTROŠNJA (l/s)`
- `MAKSIMALNA DNEVNA POTROŠNJA (m3)`
- `MAKSIMALNA DNEVNA POTROŠNJA (l/s)`
- `MINIMALNA DNEVNA POTROŠNJA (m3)`
- `MINIMALNA DNEVNA POTROŠNJA (l/s)`

## Setup

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
python -m playwright install chromium
```

Fill `.env`:

```dotenv
url=https://your-site.example
email=your-email@gmail.com
password=your-password

HEADLESS=false
SLOW_MO_MS=0
CHART_WAIT_MS=5000
WORKERS=1
KEEP_BROWSER_OPEN=false
WORKBOOK_PATH=EcoKing - tabela potrošnje - Jul 2026. TEST.xlsx
LOCATION_MAP_PATH=herceg_novi_stations.json
```

## Run

Desktop launcher:

```bash
python ecoking_launcher.py
```

Linux browser preview:

```bash
python ecoking_web_launcher.py
```

Then open:

```text
http://127.0.0.1:8765
```

Direct scraper run:

```bash
python ecoking_daily.py --workbook "EcoKing - tabela potrošnje - Jul 2026. TEST.xlsx" --headed --verbose
```

Headless run:

```bash
python ecoking_daily.py --workbook "EcoKing - tabela potrošnje - Jul 2026. TEST.xlsx" --headless --verbose
```

Parallel headless run:

```bash
python ecoking_daily.py --workbook "EcoKing - tabela potrošnje - Jul 2026. TEST.xlsx" --headless --workers 4 --verbose
```

## Desktop App

The launcher provides:

- one-button execution
- visible/hidden browser option
- worker count, limit, and wait-time options
- file picker for the Excel workbook
- direct write-back into the selected Excel workbook
- automatic yesterday sheet creation, replacing that sheet if it already exists
- station mapping loaded from `herceg_novi_stations.json` beside the app
- live Serbian ijekavica logs
- timestamped log files under `logs/`
- buttons to open the Excel file and logs folder
- automatic opening of the updated Excel file after a successful run

Build a Windows folder-based EXE on the Windows machine:

```bat
build_windows.bat
```

The output is:

```text
dist\EcoKingRunner\EcoKingRunner.exe
```

Ship the whole `dist\EcoKingRunner` folder to the employee, not only the EXE. Keep `.env`, `herceg_novi_stations.json`, and the workbook beside `EcoKingRunner.exe`.

## Windows 10 Deployment

Build on a Windows machine with the same architecture as the employee machine, usually 64-bit Windows:

1. Install Python 3.10 or 3.11 from python.org and enable `Add python.exe to PATH`.
2. Copy this project folder to the Windows machine.
3. Put the real `.env`, `herceg_novi_stations.json`, and workbook in the project folder.
4. Double-click `build_windows.bat`.
5. Give the employee the whole `dist\EcoKingRunner` folder.
6. The employee runs `EcoKingRunner.exe`.

The EXE folder should contain the app, Playwright runtime, bundled Chromium, `.env`, `herceg_novi_stations.json`, and the workbook.

If Windows Defender blocks the EXE, choose `More info` -> `Run anyway`, or add the folder to the allowed list after internal review.

## Selector Overrides

The script includes generic fallbacks, but exact selectors can be added to `.env` if the EcoKing UI changes:

```dotenv
LOGIN_EMAIL_SELECTOR=input[type="email"]
LOGIN_PASSWORD_SELECTOR=input[type="password"]
LOGIN_SUBMIT_SELECTOR=button[type="submit"]

SEARCH_TOGGLE_SELECTOR=.content-header button.btn-primary
SEARCH_INPUT_SELECTOR=input[type="search"]

INTERVAL_SELECTOR=select
INTERVAL_1_DAY_LABEL=1 Day (15-minutely)
INTERVAL_30_DAYS_LABEL=30 days
```

When chart extraction fails, the script writes a screenshot and HTML snapshot under `debug/` for inspection.
