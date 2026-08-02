# EcoKing Runner

Reads daily water-consumption values from the EcoKing site and fills in the
`ECO KING BLANKO TABLICA.xlsx` report — the work that used to be done by hand.

There are two ways to run it:

| | Web UI | Windows desktop app |
| --- | --- | --- |
| Start | `python -m ecoking.webapp` or the hosted URL | `EcoKingRunner.exe` |
| Report lands in | browser download / Desktop | Desktop |
| Edit the station list | yes, built in | no — use the web UI |

---

## How a run works

1. You pick a date — or several, see below.
2. For every station in `stations.json`, the scraper searches the site for the
   device, opens its diagram, and reads the 30-day total plus the 15-minute
   min/max.
3. `ECO KING BLANKO TABLICA.xlsx` is copied, the sheet is renamed to the date,
   and four cells per station are filled in: daily m³, max m³, min m³, battery.
   The `l/s` formulas and all styling come from the template untouched.
4. A second pass opens the telemetry site, reads each reservoir's level at
   17:00 on the same date, and fills `NIVO REZERVOARA U 17h` in the report just
   written. See below.

---

## Several days at once

In the web UI the date is a month grid, folded away behind **Kalendar** so the
usual one-day run stays a single click. Open it and: **click** picks a day,
**Shift+click** picks a range from the last one you clicked, and a second click
takes a day back out — so three consecutive days and three scattered ones are
the same gesture. A dot under a day means a report for it already exists, so
re-running it overwrites that file. Up to 14 days per run. Whatever is picked
shows as chips under the calendar, and on the closed summary.

Nothing about a day changes because it is part of a batch: it is still one
`ecoking_daily.py --selected-date` process writing one
`EcoKing_Report_<date>.xlsx`. The days simply run **one after another** —
overlapping them would mean several logins and several browsers against the
same site. Pick one day and the screen is exactly what it always was.

Every date is checked before the first one starts, so a typo cannot surface
twenty minutes in. After that:

* **a day that fails does not stop the queue** — that is the point; the other
  days are still worth having;
* **Zaustavi** ends the whole run and marks the days that never started as
  *preskočeno*;
* **Preskoči tekući dan** drops only the day in progress and moves on.

When it is done, *Rezultat po danima* lists each day with its outcome and a
**Ponovi** button — which is just a one-day run of the same thing.

### Seeing what failed

Failures are pulled out of the run's own report lines and grouped by day in the
**Problemi** panel: the device or location, its Excel row, and the reason.
A red ✕ is a station that produced nothing; an amber ⚠ is a gap — `NO DATA`, or
a 17h level that could not be read — in a report that is otherwise finished.

The console keeps every line, and the **Sve / Greške / + upozorenja** filter
plus the **Dan** dropdown narrow it down; *Kopiraj* copies whatever is showing,
so sending on just the errors is one click. Severity comes from the log level
the scraper already prints, so nothing new has to be parsed by hand.

### A note on the language

`ecoking_daily.py` logs in English and `ecoking/logtext.py` translates on the
way to the screen — including the reason inside a `NEUSPJEH` line, which is
exception text and so nests (`translate_reason`). Keep it that way when adding
messages: the scraper *branches on its own English wording* to tell a failure
from missing data, so translating at the source would quietly change which
bucket a station lands in. `tests/test_batch_run.py` checks the coverage.

---

## The 17h reservoir levels

The levels are not in EcoKing — they come from the ViK telemetry system
(`nadzorhnvik`), so they are collected in a separate pass once the report
exists. Two files next to the app drive it:

* **`telemetry_list.py`** — `locations`: which rows of the site's *Spisak
  mjernih mjesta* table to visit, spelled as they read on screen. The trailing
  `Tele` is the antenna icon's alt text; it is ignored during matching, so
  `Rezervoar KulaTele` and `Rezervoar Kula` both work.
* **`telemetry_mapping.json`** — site `Lokacija` → report `LOKACIJA`:

  ```json
  { "Rezervoar Podi Tele": "REZERVOAR PODI" }
  ```

Each location is clicked open, its filter is narrowed to the selected date, and
the `Nivo (m)` cell for `17:00` is read. The value goes into **every** meter row
of the mapped `LOKACIJA` — `ULAZ` and `IZLAZ` alike — because it describes the
reservoir, not one meter.

Pages that report two chambers (Kumbor) have an `M1 Nivo` and an `M2 Nivo`
column; `mjerač 1` / `mjerač 2` in the location name picks the right one. Pages
where two meters share one level (Bajer 2) give both names the same value, so
the mapping only needs one of them.

Credentials live in `.env` as `TELEMETRIJA_URL`, `TELEMETRIJA_USERNAME` and
`TELEMETRIJA_PASSWORD`. A location the site never loaded, a missing 17:00 row,
or a location absent from `telemetry_mapping.json` is logged and skipped — the
consumption report is already written by then and is never lost to it. Set
`TELEMETRY_ENABLED=false` to skip the pass entirely.

### Triggering it

In the web UI (**Obračun**):

* **☑ Telemetrija** — on by default. Leave it checked and *Pokreni obračun*
  runs both passes back to back, into one file.
* **Telemetrija** (button) — runs only this pass, over the reports that already
  exist for the selected dates. Use it when the levels failed but the
  consumption numbers are fine. It refuses, naming the dates, if any of them has
  no report yet, rather than making a half-empty one.

**Napredna podešavanja** is split by pass. Workers, limit, chart wait, search
wait and *Prikaži browser* drive the EcoKing scrape only; the telemetry pass has
its own **Čekanje po lokaciji (ms)** (default 10000) and **Prikaži browser
(telemetrija)**. The desktop app has the same split. From a terminal:

```bash
# both passes (the default)
python ecoking_daily.py --selected-date 2026-07-31

# EcoKing only
python ecoking_daily.py --selected-date 2026-07-31 --skip-telemetry

# levels only, into the report that run already produced
python ecoking_daily.py --selected-date 2026-07-31 --only-telemetry \
  --output "EcoKing_Report_2026-07-31.xlsx"

# watch only the telemetry browser, and give each location 20s
python ecoking_daily.py --selected-date 2026-07-31 --headless \
  --telemetry-headed --telemetry-wait-ms 20000

# levels only, standalone, watching the browser
python -m ecoking.telemetry --workbook "EcoKing_Report_2026-07-31.xlsx" \
  --date 2026-07-31 --headed
```

---

## The station list

`stations.json` is the only place the mapping lives. One entry per report row:

```json
{
  "lokacija": "REZERVOAR BAJER 1",
  "vodomjer": "ULAZ",
  "uredjaj": "Bajer 1 - U"
}
```

* `lokacija` + `vodomjer` must match a `LOKACIJA` / `VODOMJER` pair in the
  template exactly. That pair is unique, so it picks the row with no guessing.
* `uredjaj` is the device name **as it reads on the EcoKing site, without the
  serial number and without `Herceg Novi`**. Serial numbers are not used
  anywhere any more. Pasting a whole dropdown entry
  (`358004092223510 - Herceg Novi - (R-PO) Podi - I (358004092223510)`) works
  too — saving from **Stanice** reduces it to `(R-PO) Podi - I`.
* `"enabled": false` keeps an entry in the file but skips it during a run.

### Adding or changing an entry

Open the web UI → **Stanice**:

* **Red u Excel tabeli** is a dropdown of the real template rows, so it is not
  possible to point an entry at a row that does not exist.
* **Naziv uređaja** is typed by hand — the device name as it reads on the
  EcoKing site, without the serial number and without `Herceg Novi`. Pasting
  the site's full dropdown entry is fine; it is shortened when you save.
* **Sačuvaj** validates before writing. Blocking problems (no such row, two
  entries on one row, empty device name) are refused; the rest show as warnings.

If two devices on the site share a name — `(R-PO) Podi - I` and
`(PS-PO) Podi - I` — a run stops with a message listing both instead of writing
the wrong number. Add the bracketed prefix to whichever entry is ambiguous.

Validate the file from a terminal or in CI:

```bash
python -m ecoking.check
```

### Changing the template

Add or rename a row in `ECO KING BLANKO TABLICA.xlsx`, then reopen **Stanice**.
The row dropdown is read straight from the workbook, so the new row is there
immediately; template rows with no station entry show as a warning.

---

## Local setup

```bash
python -m venv .venv
.venv\Scripts\activate          # Linux/macOS: . .venv/bin/activate
pip install -r requirements.txt
python -m playwright install chromium
cp .env.example .env            # then fill in url / email / password
```

Run the web UI:

```bash
python -m ecoking.webapp        # http://127.0.0.1:8765
```

Run the desktop app:

```bash
python ecoking_launcher.py
```

Run the scraper directly:

```bash
python ecoking_daily.py --selected-date 2026-07-23 --headed --verbose
python ecoking_daily.py --selected-date 2026-07-23 --headless --workers 4
```

Tests:

```bash
python -m unittest discover -s tests -t .
```

---

## Running it locally, with no hosting at all

For a single operator, this is the simplest option of all: the app runs on
their own PC, uses their own hardware, and never touches the internet except
to talk to the EcoKing site and to check for updates. No OOM risk, no hosting
cost, no signup, and fixes arrive with a double-click instead of a
rebuild → zip → send cycle.

There are two ways to get it onto that PC: send them a ready-to-run build (no
Python or Git needed on their end), or have them set it up from source. For
someone who isn't going to run installers or type commands, send the build.

### Option A — send a ready-to-run build (recommended for a non-technical user)

**Building it** (on your machine):

```bat
build_windows_web.bat
```

This produces `dist\EcoKingWebRunner\` — a self-contained folder with Python,
Chromium, and every dependency already installed. It's large (roughly 800 MB
unzipped, ~340 MB as a zip) because Chromium is bundled, so:

- **Zip the whole `dist\EcoKingWebRunner` folder.** Email won't take a 340 MB
  attachment — use a USB drive, or upload the zip to OneDrive/Google Drive and
  send a share link instead.
- Before zipping, make sure the `.env` inside that folder has the real
  `url` / `email` / `password` — `build_windows_web.bat` copies whatever
  `.env` sits in the project root at build time.

**What they do, once:**

1. Copy the zip over (USB, or download it from your share link) and extract
   it anywhere — e.g. `Desktop\EcoKingRunner`.
2. Double-click `EcoKingWebRunner.exe`.
3. Windows SmartScreen will likely say *"Windows protected your PC"* the first
   time, since the exe isn't code-signed — click **More info → Run anyway**.
4. The app opens in their browser automatically, at
   `http://127.0.0.1:8765`. Nothing else on the network can reach it.

**Getting updates:** every launch of `EcoKingWebRunner.exe` quietly pulls the
latest application code from this GitHub repo before starting — so most fixes
just show up the next time they open it. Nothing to resend, nothing to
reinstall. `.env`, `stations.json`, the report template, and any saved
reports are never touched by this — only the app's own code is synced.

This only works because the repo is **public**; the exe downloads a plain zip
of the `main` branch with no login. If it were private, this update mechanism
would need a token and stop working silently (falling back to whatever
version is already installed, which is safe but quiet about it — check the
console window it opens for `Could not check for updates` if something seems
stale).

**When a full rebuild is still needed:** if a change adds a new Python
dependency (rare — logic/UI fixes never do), the auto-update won't pick it up,
since the frozen Python environment itself isn't part of the sync. Run
`build_windows_web.bat` again and resend in that case, same as before.

### Option B — set it up from source (if you're doing it yourself, e.g. over remote access)

**One-time setup**, on the machine that will run it:

1. Install [Python 3.11 or 3.10](https://python.org) — tick *Add python.exe to
   PATH* during install.
2. Install [Git for Windows](https://git-scm.com/download/win) (default
   options are fine).
3. `git clone https://github.com/jkonjevic/EcoKingRunner.git` into a normal
   folder (e.g. `Documents\EcoKingRunner`).
4. Copy a real `.env` into that folder (see `.env.example` for the shape —
   needs `url`, `email`, `password`). This is the one file that has to be
   copied by hand; it's deliberately never committed to git.
5. Double-click **`Update.bat`**. First run installs everything (a few
   minutes, mostly downloading Chromium); every later run just syncs it.

**Daily use:** double-click **`Start.bat`**. It quietly checks for the latest
code, starts the app, and opens it in the browser automatically at
`http://127.0.0.1:8765`.

**Getting updates:** `Start.bat` already pulls the latest code every time it
runs, so most fixes just show up next time it's launched. If a change touches
a dependency (rare) and something breaks after an update, run **`Update.bat`**
once to do a full resync.

---

## Hosting it (Render)

Free, no card required, deploys straight from this GitHub repo, and
redeploys itself on every push to `main` — that's the CI/CD part, handled by
Render's own git integration, not a separate GitHub Action.

> Hugging Face Spaces looks like the obvious free option but isn't: as of
> 2026, Docker (and Gradio) Spaces require an HF **PRO** plan ($9/mo) to
> create at all — only Static Spaces are free, and those can't run this app.
> Render and Koyeb still have genuine free Docker tiers with no card on file.

**1. Create the Render account and connect the repo.**
[dashboard.render.com](https://dashboard.render.com) → sign up (GitHub login
is fastest) → **New** → **Blueprint** → pick this repo. Render reads
[`render.yaml`](render.yaml) and proposes the service automatically — plan
already set to **Free**.

**2. Fill in the secrets it prompts for.** Render only asks for these once,
during that same Blueprint creation screen:

| Name | Value |
| --- | --- |
| `url` | EcoKing site address |
| `email` | login |
| `password` | login password |
| `APP_PASSWORD` | a password you make up — this gates *your app*, not EcoKing |

**3. Click Deploy.** First build takes a few minutes (installing Chromium).
After that, every `git push` to `main` triggers a rebuild automatically.

Render gives you the URL once it's live — something like
`https://ecoking-runner.onrender.com`.

### Things to know about the free tier

* **512 MB RAM, shared CPU.** Enough for one browser at a time — the app
  already defaults to `WORKERS=1` in `render.yaml`. Don't raise it on the free
  plan; Chromium can push past 512 MB with more than one worker and the
  instance gets killed (OOM), not gracefully slowed down.
* **It sleeps after 15 minutes idle.** The first request after that wakes it
  back up, which takes 30-60 seconds. Fine for a once-a-day report; annoying
  if you're actively iterating and keep losing the tab.
* **The disk is not persistent.** Reports and station-list edits live under
  `DATA_DIR` and are wiped on every redeploy. Download the Excel right after a
  run, and mirror any station edits back into `stations.json` in git so they
  survive the next push.
* **If the EcoKing site restricts logins by IP**, Render's address has to be
  allowed, or none of this works — worth confirming with one real run before
  relying on it.

### If Render's RAM is too tight

[Koyeb](https://www.koyeb.com) has a comparable free tier (512 MB / 0.1 vCPU,
no card required, scales to zero after an hour idle) and runs the same
Dockerfile — the app already auto-detects it via Koyeb's `KOYEB_APP_NAME`
environment variable. Push the repo, point Koyeb at the Dockerfile, add the
same four secrets.

### Any other host

The image is a plain Dockerfile — a company VM, Cloud Run, or anything else
that runs containers works too:

```bash
docker build -t ecoking .
docker run -p 8765:8765 --env-file .env -e APP_PASSWORD=... ecoking
```

---

## Building the Windows EXE

On a 64-bit Windows machine with Python 3.10 or 3.11 (`Add python.exe to PATH`
ticked):

```bat
build_windows.bat
```

Ship the whole `dist\EcoKingRunner` folder, not just the EXE. It has to contain
`.env`, `stations.json`, and `ECO KING BLANKO TABLICA.xlsx` beside
`EcoKingRunner.exe`. If Defender blocks it: *More info* → *Run anyway*.

---

## Layout

```
ecoking/
  stations.py    station registry: file format, template rows, validation
  webapp.py      web UI server and JSON API
  logtext.py     English log lines -> Montenegrin, shared by both UIs
  telemetry.py   second pass: 17h reservoir levels
  check.py       python -m ecoking.check
web/             the web UI (index.html, styles.css, app.js)
ecoking_daily.py the scraper and the report writer
ecoking_launcher.py  Windows desktop app
stations.json    station -> template row -> device name
telemetry_list.py       which telemetry locations to visit
telemetry_mapping.json  telemetry Lokacija -> report LOKACIJA
tests/
```

## When something breaks

* **`Naziv ... odgovara za N uređaja`** — the device name is ambiguous. Add the
  bracketed prefix in **Stanice**.
* **`Nijedan rezultat ne odgovara nazivu`** — the name does not exist on the
  site as spelled. Open a station's diagram on the site and copy its exact
  label into **Stanice**; the serial and `Herceg Novi` are stripped on save, so
  copying the whole dropdown entry is fine.
* **Chart extraction fails** — a screenshot and an HTML snapshot are written to
  `debug/`. If the site's markup changed, set the matching `*_SELECTOR` in
  `.env` (see `.env.example`).
* **`Telemetrija: ... nema unos u telemetry_mapping.json`** — the location was
  read fine but has nowhere to go in the report. Add it to
  `telemetry_mapping.json`, or drop it from `telemetry_list.py`.
* **`Nema reda za 17:00 ...`** — the telemetry site has no reading at 17:00 for
  that date and location. Nothing to fix in the app; check the location's table
  on the site.
* **Telemetry pages load slowly** — raise `TELEMETRY_WAIT_MS` in `.env` (default
  10000 ms per location). The run also keeps polling past that wait.
* **One day of a multi-day run failed** — open **Problemi**, which names the
  devices and the reason per day, and press *Ponovi ovaj dan*. The days that did
  work are already saved; nothing has to be re-run for them.
* **A report is missing from Izvještaji** — the list is a live listing of the
  reports folder (**📁 Otvori folder** opens it, and 📁 on a row shows that one
  file in it; locally it is the Desktop), minus any row hidden with ✕. ✕ never
  touches the file, so
  **Vrati uklonjene** brings the rows back, and re-running a hidden date
  un-hides it automatically. The hidden dates live in `hidden_reports.json`
  next to the app.
