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

1. You pick a date.
2. For every station in `stations.json`, the scraper searches the site for the
   device, opens its diagram, and reads the 30-day total plus the 15-minute
   min/max.
3. `ECO KING BLANKO TABLICA.xlsx` is copied, the sheet is renamed to the date,
   and four cells per station are filled in: daily m³, max m³, min m³, battery.
   The `l/s` formulas and all styling come from the template untouched.

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
  anywhere any more.
* `"enabled": false` keeps an entry in the file but skips it during a run.

### Adding or changing an entry

Open the web UI → **Stanice**:

* **Red u Excel tabeli** is a dropdown of the real template rows, so it is not
  possible to point an entry at a row that does not exist.
* **Naziv uređaja** is typed by hand — the device name as it reads on the
  EcoKing site, without the serial number and without `Herceg Novi`.
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
  logtext.py     English log lines -> Serbian, shared by both UIs
  check.py       python -m ecoking.check
web/             the web UI (index.html, styles.css, app.js)
ecoking_daily.py the scraper and the report writer
ecoking_launcher.py  Windows desktop app
stations.json    station -> template row -> device name
tests/
```

## When something breaks

* **`Naziv ... odgovara za N uređaja`** — the device name is ambiguous. Add the
  bracketed prefix in **Stanice**.
* **`Nijedan rezultat ne odgovara nazivu`** — the name does not exist on the
  site as spelled. Open a station's diagram on the site and copy its exact
  label into **Stanice**.
* **Chart extraction fails** — a screenshot and an HTML snapshot are written to
  `debug/`. If the site's markup changed, set the matching `*_SELECTOR` in
  `.env` (see `.env.example`).
