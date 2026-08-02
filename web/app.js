/* EcoKing web UI ---------------------------------------------------------- */
"use strict";

const $ = (id) => document.getElementById(id);

const api = {
  async get(path) {
    const response = await fetch(path, { headers: { Accept: "application/json" } });
    return unwrap(response);
  },
  async post(path, body) {
    const response = await fetch(path, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body || {}),
    });
    return unwrap(response);
  },
};

async function unwrap(response) {
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(payload.error || `Greška ${response.status}`);
    error.status = response.status;
    error.authRequired = Boolean(payload.authRequired);
    throw error;
  }
  return payload;
}

function notify(element, text, tone) {
  element.textContent = text || "";
  element.className = "notice" + (tone ? ` is-${tone}` : "");
  element.hidden = !text;
}

/* Theme -------------------------------------------------------------------- */

function initTheme() {
  const stored = localStorage.getItem("ecoking-theme");
  if (stored) document.documentElement.dataset.theme = stored;
  $("themeToggle").addEventListener("click", () => {
    const dark = document.documentElement.dataset.theme
      ? document.documentElement.dataset.theme === "dark"
      : matchMedia("(prefers-color-scheme: dark)").matches;
    const next = dark ? "light" : "dark";
    document.documentElement.dataset.theme = next;
    localStorage.setItem("ecoking-theme", next);
  });
}

/* Login gate --------------------------------------------------------------- */

async function initSession() {
  const session = await api.get("/api/session");
  if (!session.authRequired || session.authenticated) return true;

  $("gate").hidden = false;
  return new Promise((resolve) => {
    $("loginForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      try {
        await api.post("/api/login", { password: $("password").value });
        $("gate").hidden = true;
        resolve(true);
      } catch (error) {
        $("loginError").textContent = error.message;
      }
    });
  });
}

/* Dates -------------------------------------------------------------------- */

const MONTHS = ["januar", "februar", "mart", "april", "maj", "jun",
  "jul", "avgust", "septembar", "oktobar", "novembar", "decembar"];
const WEEKDAYS = ["po", "ut", "sr", "če", "pe", "su", "ne"];

function iso(date) {
  const month = String(date.getMonth() + 1).padStart(2, "0");
  return `${date.getFullYear()}-${month}-${String(date.getDate()).padStart(2, "0")}`;
}

function fromIso(text) {
  const [year, month, day] = text.split("-").map(Number);
  return new Date(year, month - 1, day);
}

/** Short label for chips and day segments: 2026-07-31 -> 31.07. */
function shortDate(text) {
  const [, month, day] = text.split("-");
  return `${day}.${month}.`;
}

/**
 * The month grid. Click toggles a day, Shift+click extends from the last one,
 * so three consecutive days and three scattered ones are the same gesture.
 */
const picker = {
  selected: new Set(),
  known: new Set(),
  month: null,
  today: "",
  anchor: null,
  onChange: null,

  init(bootstrap, onChange) {
    this.today = bootstrap.today;
    this.onChange = onChange;
    this.selected.add(bootstrap.defaultDate);
    this.anchor = bootstrap.defaultDate;
    this.month = fromIso(bootstrap.defaultDate);
    this.month.setDate(1);

    const header = $("calDow");
    for (const name of WEEKDAYS) {
      const cell = document.createElement("span");
      cell.textContent = name;
      header.appendChild(cell);
    }
    $("calPrev").addEventListener("click", () => this.shiftMonth(-1));
    $("calNext").addEventListener("click", () => this.shiftMonth(1));
    this.render();
  },

  list() {
    return [...this.selected].sort();
  },

  shiftMonth(step) {
    this.month.setMonth(this.month.getMonth() + step);
    this.render();
  },

  /** Dates that already have a report get a dot, so a re-run is deliberate. */
  setKnown(dates) {
    this.known = new Set(dates);
    this.render();
  },

  toggle(date, extend) {
    if (extend && this.anchor) {
      const [from, to] = [this.anchor, date].sort();
      for (const cursor = fromIso(from); iso(cursor) <= to; cursor.setDate(cursor.getDate() + 1)) {
        if (iso(cursor) <= this.today) this.selected.add(iso(cursor));
      }
    } else if (this.selected.has(date)) {
      this.selected.delete(date);
    } else {
      this.selected.add(date);
    }
    this.anchor = date;
    this.render();
  },

  clear() {
    this.selected.clear();
    this.render();
  },

  render() {
    $("calLabel").textContent = `${MONTHS[this.month.getMonth()]} ${this.month.getFullYear()}`;
    const grid = $("calGrid");
    grid.textContent = "";

    const year = this.month.getFullYear();
    const month = this.month.getMonth();
    // Monday-first, which is how the calendar reads here.
    const lead = (new Date(year, month, 1).getDay() + 6) % 7;
    const length = new Date(year, month + 1, 0).getDate();
    for (let index = 0; index < lead; index += 1) grid.appendChild(document.createElement("span"));

    for (let day = 1; day <= length; day += 1) {
      const date = iso(new Date(year, month, day));
      const cell = document.createElement("button");
      cell.type = "button";
      cell.className = "cal-day";
      cell.textContent = String(day);
      cell.title = date;
      if (date > this.today) cell.disabled = true;
      if (date === this.today) cell.classList.add("is-today");
      if (this.known.has(date)) cell.classList.add("has-report");
      if (this.selected.has(date)) {
        cell.classList.add("is-on");
        cell.setAttribute("aria-pressed", "true");
      }
      cell.addEventListener("click", (event) => this.toggle(date, event.shiftKey));
      grid.appendChild(cell);
    }

    this.renderChips();
    if (this.onChange) this.onChange(this.list());
  },

  renderChips() {
    const container = $("dateChips");
    container.textContent = "";
    const dates = this.list();
    for (const date of dates) {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "date-chip";
      chip.title = `Ukloni ${date}`;
      const label = document.createElement("span");
      label.textContent = shortDate(date);
      const cross = document.createElement("b");
      cross.textContent = "✕";
      chip.append(label, cross);
      chip.addEventListener("click", () => this.toggle(date, false));
      container.appendChild(chip);
    }
    if (dates.length > 1) {
      const clear = document.createElement("button");
      clear.type = "button";
      clear.className = "date-chip is-clear";
      clear.textContent = `Očisti (${dates.length})`;
      clear.addEventListener("click", () => this.clear());
      container.appendChild(clear);
    }
  },
};

/* Run view ----------------------------------------------------------------- */

/** Severity is read off the translated prefix the server already writes. */
const LEVEL_RE = /^\d{2}:\d{2}:\d{2}\s+(GREŠKA|KRITIČNO|UPOZORENJE)\s/;
/** Day separator the server writes between the days of a batch. */
const DAY_RE = /DAN\s+\d+\/\d+\s+·\s+(\d{4}-\d{2}-\d{2})/;
/** A long batch can log tens of thousands of lines; only the tail is drawn. */
const LOG_CAP = 4000;

const STATUS_TEXT = {
  pending: "na čekanju",
  running: "u toku",
  ok: "uspješno",
  failed: "greška",
  stopped: "zaustavljeno",
  skipped: "preskočeno",
};
const STATUS_ICON = { pending: "·", running: "◍", ok: "✓", failed: "✕", stopped: "■", skipped: "—" };

function levelOf(text) {
  const found = LEVEL_RE.exec(text);
  if (!found) return "info";
  return found[1] === "UPOZORENJE" ? "warning" : "error";
}

const run = {
  cursor: 0,
  timer: null,
  bootstrap: null,
  entries: [],
  level: "all",
  day: "all",
  currentDate: null,
  daySignature: "",

  async init(bootstrap) {
    this.bootstrap = bootstrap;
    $("modeBadge").textContent = bootstrap.cloud ? "Online" : "Lokalno";
    $("dateHint").textContent =
      `${bootstrap.stationCount} aktivnih stanica · template: ${bootstrap.template}`;
    // The list is a live listing of this folder; the ⓘ next to it explains
    // that, so the path itself can stay short and just be the path.
    $("reportsPath").textContent = bootstrap.reportsLocation;
    $("reportsPath").title = bootstrap.reportsLocation;

    // A hosted container has no screen, no Excel and no file manager.
    $("browserVisibleRow").hidden = bootstrap.cloud;
    $("telemetryVisibleRow").hidden = bootstrap.cloud;
    $("openAfterRow").hidden = bootstrap.cloud;
    $("openFolder").hidden = bootstrap.cloud;

    picker.init(bootstrap, (dates) => this.onDatesChanged(dates));

    $("runBtn").addEventListener("click", () => this.start());
    $("telemetryBtn").addEventListener("click", () => this.start({ onlyTelemetry: true }));
    $("stopBtn").addEventListener("click", () => this.stop("batch"));
    $("skipDayBtn").addEventListener("click", () => this.stop("day"));
    $("copyLog").addEventListener("click", () => navigator.clipboard.writeText(this.visibleText()));
    $("openFolder").addEventListener("click", () => this.openFolder());
    $("problemsCopy").addEventListener("click", () => navigator.clipboard.writeText(this.problemsText()));

    for (const button of document.querySelectorAll(".filter")) {
      button.addEventListener("click", () => {
        this.level = button.dataset.level;
        for (const other of document.querySelectorAll(".filter")) {
          other.classList.toggle("is-active", other === button);
        }
        this.renderLog();
      });
    }
    $("dayFilter").addEventListener("change", () => {
      this.day = $("dayFilter").value;
      this.renderLog();
    });

    await this.refreshReports();
    this.poll();
    this.timer = setInterval(() => this.poll(), 1000);
  },

  /** One day selected must look and behave exactly like it always has. */
  onDatesChanged(dates) {
    const many = dates.length > 1;
    $("runBtn").textContent = many ? `Pokreni obračun (${dates.length} dana)` : "Pokreni obračun";
    $("telemetryBtn").textContent = many ? `Telemetrija (${dates.length} dana)` : "Telemetrija";
    // The calendar is collapsed by default, so its summary carries the answer.
    $("calSummary").textContent = dates.length
      ? many
        ? `${dates.length} dana izabrano`
        : dates[0].split("-").reverse().join(".") + "."
      : "Izaberi datume";
    const idle = !this.running;
    $("runBtn").disabled = !idle || !dates.length;
    $("telemetryBtn").disabled = !idle || !dates.length;
  },

  config() {
    return {
      selectedDates: picker.list(),
      workers: Number($("workers").value || 1),
      limit: $("limit").value,
      chartWait: Number($("chartWait").value || 5000),
      searchWait: Number($("searchWait").value || 2000),
      browserVisible: $("browserVisible").checked,
      verbose: $("verbose").checked,
      withTelemetry: $("withTelemetry").checked,
      telemetryWait: Number($("telemetryWait").value || 10000),
      telemetryVisible: $("telemetryVisible").checked,
    };
  },

  async start(overrides = {}) {
    notify($("runMessage"), "");
    const payload = { ...this.config(), ...overrides };
    if (!payload.selectedDates.length) {
      notify($("runMessage"), "Izaberi bar jedan datum u kalendaru.", "error");
      return;
    }
    try {
      await api.post("/api/run", payload);
      this.resetLog();
      $("summaryCard").hidden = true;
      this.poll();
    } catch (error) {
      notify($("runMessage"), error.message, "error");
    }
  },

  resetLog() {
    this.cursor = 0;
    this.entries = [];
    this.currentDate = null;
    this.daySignature = "";
    this.problemSignature = "";
    this.level = "all";
    this.day = "all";
    $("dayFilter").value = "all";
    $("problems").hidden = true;
    for (const button of document.querySelectorAll(".filter")) {
      button.classList.toggle("is-active", button.dataset.level === "all");
    }
    $("log").textContent = "";
    this.updateCounts();
  },

  async stop(scope) {
    await api.post("/api/stop", { scope }).catch(() => {});
    this.poll();
  },

  async poll() {
    let state;
    try {
      state = await api.get(`/api/state?cursor=${this.cursor}`);
    } catch (error) {
      if (error.authRequired) location.reload();
      return;
    }

    this.cursor = state.cursor;
    if (state.lines.length) this.ingest(state.lines);

    const wasRunning = this.running;
    this.running = state.running;
    const days = state.days || [];
    const batch = state.batch || { total: days.length, index: 0 };

    $("runBtn").disabled = state.running || !picker.list().length;
    $("telemetryBtn").disabled = state.running || !picker.list().length;
    $("stopBtn").disabled = !state.running;
    // Only worth offering while a batch is actually mid-flight.
    $("skipDayBtn").hidden = batch.total < 2 || !state.running;

    const status = $("statStatus");
    if (state.running) {
      status.textContent = "U toku";
      status.className = "is-running";
    } else if (state.returnCode === 0) {
      status.textContent = "Završeno";
      status.className = "is-ok";
    } else if (state.returnCode != null) {
      status.textContent = "Greška";
      status.className = "is-error";
    } else {
      status.textContent = "Spremno";
      status.className = "";
    }

    $("statDayCard").hidden = batch.total < 2;
    $("statDay").textContent = batch.total > 1 ? `${batch.index + 1} / ${batch.total}` : "—";
    $("statProgress").textContent = `${state.done} / ${state.total}`;
    $("statCurrent").textContent = state.current || "—";
    $("statCurrent").title = state.current || "";
    $("statStarted").textContent = state.startedAt || "—";
    $("progressBar").style.width = state.total
      ? `${Math.min(100, Math.round((state.done * 100) / state.total))}%`
      : "0%";

    this.lastDays = days;
    this.renderDayStrip(days);
    this.renderProblems(days);
    if (!state.running) this.renderSummary(days);

    if (wasRunning && !state.running) await this.onFinished(state, days);
  },

  /* -- log ---------------------------------------------------------------- */

  ingest(lines) {
    const fresh = [];
    for (const text of lines) {
      const marker = DAY_RE.exec(text);
      if (marker) this.currentDate = marker[1];
      const entry = { index: this.entries.length, text, level: levelOf(text), date: this.currentDate };
      this.entries.push(entry);
      fresh.push(entry);
    }
    this.updateCounts();
    this.appendLines(fresh.filter((entry) => this.matches(entry)));
  },

  matches(entry) {
    if (this.day !== "all" && entry.date !== this.day) return false;
    if (this.level === "error") return entry.level === "error";
    if (this.level === "warning") return entry.level !== "info";
    return true;
  },

  lineNode(entry) {
    const line = document.createElement("div");
    line.className = entry.level === "info" ? "log-line" : `log-line is-${entry.level}`;
    line.textContent = entry.text;
    return line;
  },

  appendLines(entries) {
    if (!entries.length) return;
    const box = $("log");
    const pinned = $("autoScroll").checked;
    const batch = document.createDocumentFragment();
    for (const entry of entries) batch.appendChild(this.lineNode(entry));
    box.appendChild(batch);
    while (box.childElementCount > LOG_CAP) box.removeChild(box.firstChild);
    if (pinned) box.scrollTop = box.scrollHeight;
  },

  renderLog() {
    const box = $("log");
    box.textContent = "";
    const visible = this.entries.filter((entry) => this.matches(entry));
    const batch = document.createDocumentFragment();
    for (const entry of visible.slice(-LOG_CAP)) batch.appendChild(this.lineNode(entry));
    box.appendChild(batch);
    box.scrollTop = box.scrollHeight;
    if (!visible.length) {
      const empty = document.createElement("div");
      empty.className = "log-line is-empty";
      empty.textContent = "Nema linija koje odgovaraju filteru.";
      box.appendChild(empty);
    }
  },

  visibleText() {
    return this.entries.filter((entry) => this.matches(entry)).map((entry) => entry.text).join("\n");
  },

  updateCounts() {
    let errors = 0;
    let warnings = 0;
    for (const entry of this.entries) {
      if (entry.level === "error") errors += 1;
      else if (entry.level === "warning") warnings += 1;
    }
    document.querySelector('[data-count="error"]').textContent = errors ? `(${errors})` : "";
    document.querySelector('[data-count="warning"]').textContent =
      errors + warnings ? `(${errors + warnings})` : "";
  },

  /* -- batch panels -------------------------------------------------------- */

  renderDayStrip(days) {
    const strip = $("dayStrip");
    strip.hidden = days.length < 2;
    $("dayFilterRow").hidden = days.length < 2;
    const signature = days.map((day) => `${day.date}:${day.status}`).join("|");
    if (days.length < 2 || signature === this.daySignature) return;
    this.daySignature = signature;

    strip.textContent = "";
    for (const day of days) {
      const segment = document.createElement("button");
      segment.type = "button";
      segment.className = `dayseg is-${day.status}`;
      segment.textContent = shortDate(day.date);
      segment.title = `${day.date} — ${STATUS_TEXT[day.status] || day.status}`;
      segment.addEventListener("click", () => {
        $("dayFilter").value = day.date;
        this.day = day.date;
        this.renderLog();
      });
      strip.appendChild(segment);
    }

    const filter = $("dayFilter");
    const previous = filter.value;
    filter.textContent = "";
    const all = document.createElement("option");
    all.value = "all";
    all.textContent = "svi";
    filter.appendChild(all);
    for (const day of days) {
      const option = document.createElement("option");
      option.value = day.date;
      option.textContent = shortDate(day.date);
      filter.appendChild(option);
    }
    filter.value = days.some((day) => day.date === previous) ? previous : "all";
  },

  renderProblems(days) {
    const withProblems = days.filter((day) => day.failures.length);
    const total = withProblems.reduce((sum, day) => sum + day.failures.length, 0);
    $("problems").hidden = !total;
    $("problemsCount").textContent = total ? `(${total})` : "";
    // Redrawing every poll would fight the user's scrolling for no reason.
    const signature = `${total}:${this.running}:${withProblems.map((day) => day.status).join("")}`;
    if (!total || signature === this.problemSignature) return;
    this.problemSignature = signature;

    const container = $("problemsList");
    container.textContent = "";
    for (const day of withProblems) {
      const group = document.createElement("div");
      group.className = "problem-group";

      const head = document.createElement("div");
      head.className = "problem-head";
      const icon = document.createElement("span");
      icon.className = `problem-icon is-${day.status}`;
      icon.textContent = STATUS_ICON[day.status] || "·";
      const label = document.createElement("b");
      label.textContent = day.date;
      const count = document.createElement("small");
      count.textContent = `${day.failures.length} problema`;
      const spacer = document.createElement("span");
      spacer.className = "spacer";
      head.append(icon, label, count, spacer);

      const onlyLevels = day.failures.every((failure) => failure.kind === "telemetry");
      const retry = document.createElement("button");
      retry.className = "btn btn-ghost btn-sm";
      retry.textContent = onlyLevels ? "Ponovi samo nivoe" : "Ponovi ovaj dan";
      retry.disabled = this.running;
      retry.addEventListener("click", () =>
        this.start({ selectedDates: [day.date], onlyTelemetry: onlyLevels })
      );
      head.appendChild(retry);
      group.appendChild(head);

      for (const failure of day.failures) {
        const item = document.createElement("div");
        item.className = `problem is-${failure.severity}`;
        const mark = document.createElement("span");
        mark.className = "problem-mark";
        mark.textContent = failure.severity === "error" ? "✕" : "⚠";
        const name = document.createElement("b");
        name.textContent = failure.kind === "telemetry" ? `TELEMETRIJA · ${failure.label}` : failure.label;
        const where = document.createElement("small");
        where.textContent = failure.row ? `red ${failure.row}` : "";
        const reason = document.createElement("span");
        reason.className = "problem-reason";
        reason.textContent = failure.reason;
        item.append(mark, name, where, reason);
        // Jumping to the day narrows the log to the lines behind this row.
        item.addEventListener("click", () => {
          $("dayFilter").value = days.length > 1 ? day.date : "all";
          this.day = days.length > 1 ? day.date : "all";
          this.renderLog();
        });
        group.appendChild(item);
      }
      container.appendChild(group);
    }
  },

  problemsText() {
    const lines = [];
    for (const day of this.lastDays || []) {
      for (const failure of day.failures) {
        const where = failure.row ? ` (red ${failure.row})` : "";
        lines.push(`${day.date} | ${failure.label}${where} | ${failure.reason}`);
      }
    }
    return lines.join("\n");
  },

  renderSummary(days) {
    const show = days.length > 1 && days.some((day) => day.status !== "pending");
    $("summaryCard").hidden = !show;
    if (!show) return;

    const ok = days.filter((day) => day.status === "ok").length;
    $("summaryHeadline").textContent = `${ok} / ${days.length} uspješno`;
    $("summaryHeadline").className = `summary-headline ${ok === days.length ? "is-ok" : "is-error"}`;

    const container = $("summaryList");
    container.textContent = "";
    for (const day of days) {
      const row = document.createElement("div");
      row.className = `summary-row is-${day.status}`;

      const icon = document.createElement("span");
      icon.className = "summary-icon";
      icon.textContent = STATUS_ICON[day.status] || "·";
      const date = document.createElement("b");
      date.textContent = day.date;
      const note = document.createElement("small");
      note.textContent = day.failures.length
        ? `${STATUS_TEXT[day.status]} · ${day.failures.length} problema`
        : STATUS_TEXT[day.status];
      const spacer = document.createElement("span");
      spacer.className = "spacer";
      row.append(icon, date, note, spacer);

      if (day.reportReady) {
        if (!this.bootstrap.cloud) {
          const open = document.createElement("button");
          open.className = "btn btn-ghost btn-sm";
          open.textContent = "Otvori";
          open.addEventListener("click", () =>
            api.post("/api/open-report", { selectedDate: day.date }).catch(() => {})
          );
          row.appendChild(open);
        }
        const download = document.createElement("a");
        download.className = "btn btn-ghost btn-sm";
        download.href = `/api/report?date=${encodeURIComponent(day.date)}`;
        download.textContent = "Preuzmi";
        row.appendChild(download);
      }
      if (day.status !== "ok") {
        const retry = document.createElement("button");
        retry.className = "btn btn-ghost btn-sm";
        retry.textContent = "Ponovi";
        retry.disabled = this.running;
        retry.addEventListener("click", () => this.start({ selectedDates: [day.date] }));
        row.appendChild(retry);
      }
      container.appendChild(row);
    }
  },

  async onFinished(state, days) {
    await this.refreshReports();
    if (days.length > 1) {
      const ok = days.filter((day) => day.status === "ok").length;
      const failed = days.length - ok;
      notify(
        $("runMessage"),
        failed
          ? `Gotovo: ${ok} od ${days.length} dana. ${failed} nije uspjelo — vidi „Rezultat po danima“.`
          : `Svih ${days.length} izvještaja je spremno.`,
        failed ? "error" : "ok"
      );
      // Opening several workbooks at once is hostile; the summary has buttons.
      return;
    }
    if (state.returnCode !== 0) {
      notify($("runMessage"), "Obračun nije uspio. Provjeri log.", "error");
      return;
    }
    notify($("runMessage"), "Izvještaj je spreman.", "ok");
    if (!this.bootstrap.cloud && $("openAfter").checked) {
      await api.post("/api/open-report", { selectedDate: state.selectedDate }).catch(() => {});
    }
  },

  /* -- reports ------------------------------------------------------------- */

  async openFolder(selectedDate) {
    notify($("runMessage"), "");
    try {
      await api.post("/api/open-folder", selectedDate ? { selectedDate } : {});
    } catch (error) {
      notify($("runMessage"), error.message, "error");
    }
  },

  async refreshReports() {
    const container = $("reportList");
    const { reports, hiddenCount } = await api
      .get("/api/reports")
      .catch(() => ({ reports: [], hiddenCount: 0 }));
    picker.setKnown(reports.map((report) => report.date));
    container.textContent = "";
    if (!reports.length) {
      const empty = document.createElement("p");
      empty.className = "report-empty";
      empty.textContent = hiddenCount
        ? "Svi izvještaji su uklonjeni iz liste."
        : "Još nema generisanih izvještaja.";
      container.appendChild(empty);
    }
    for (const report of reports.slice(0, 8)) {
      const row = document.createElement("div");
      row.className = "report-row";

      const meta = document.createElement("span");
      meta.className = "report-meta";
      const date = document.createElement("b");
      date.textContent = report.date;
      const stamp = document.createElement("small");
      stamp.textContent = report.modified;
      stamp.title = report.name;
      meta.append(date, stamp);

      const actions = document.createElement("span");
      actions.className = "report-actions";
      if (!this.bootstrap.cloud) {
        const reveal = document.createElement("button");
        reveal.className = "btn btn-ghost btn-sm btn-reveal";
        reveal.textContent = "📁";
        reveal.title = `Prikaži ${report.name} u folderu`;
        reveal.setAttribute("aria-label", reveal.title);
        reveal.addEventListener("click", () => this.openFolder(report.date));
        actions.appendChild(reveal);
      }

      const download = document.createElement("a");
      download.className = "btn btn-ghost btn-sm";
      download.href = `/api/report?date=${encodeURIComponent(report.date)}`;
      download.textContent = "Preuzmi";

      const remove = document.createElement("button");
      remove.className = "btn btn-ghost btn-sm btn-remove";
      remove.textContent = "✕";
      remove.title = `Ukloni ${report.name} iz liste (fajl ostaje na disku)`;
      remove.setAttribute("aria-label", remove.title);
      remove.addEventListener("click", () => this.hideReport(report));

      actions.append(download, remove);
      row.append(meta, actions);
      container.appendChild(row);
    }

    if (hiddenCount) {
      const restore = document.createElement("button");
      restore.className = "btn btn-ghost btn-sm report-restore";
      restore.textContent = `Vrati uklonjene (${hiddenCount})`;
      restore.addEventListener("click", () => this.restoreReports());
      container.appendChild(restore);
    }
  },

  // Only the row goes away -- the .xlsx stays where it was saved.
  async hideReport(report) {
    notify($("runMessage"), "");
    try {
      await api.post("/api/hide-report", { selectedDate: report.date });
    } catch (error) {
      notify($("runMessage"), error.message, "error");
    }
    await this.refreshReports();
  },

  async restoreReports() {
    await api.post("/api/unhide-reports").catch(() => {});
    await this.refreshReports();
  },
};

/* Stations view ------------------------------------------------------------ */

const stations = {
  rows: [],
  entries: [],
  dirty: false,

  async init() {
    $("addStation").addEventListener("click", () => this.add());
    $("saveStations").addEventListener("click", () => this.save());
    $("stationSearch").addEventListener("input", () => this.render());
    await this.reload();
  },

  async reload() {
    const payload = await api.get("/api/stations");
    this.rows = payload.rows;
    this.entries = payload.stations.map((station) => ({ ...station, enabled: station.enabled !== false }));
    this.setDirty(false);
    this.renderIssues(payload.issues);
    this.render();
  },

  setDirty(dirty) {
    this.dirty = dirty;
    $("saveStations").disabled = !dirty;
    $("saveStations").textContent = dirty ? "Sačuvaj izmjene" : "Sačuvano";
  },

  renderIssues(issues) {
    const container = $("issues");
    container.textContent = "";
    for (const issue of issues) {
      const item = document.createElement("div");
      item.className = `issue ${issue.severity}`;
      const label = document.createElement("b");
      label.textContent = issue.station;
      const message = document.createElement("span");
      message.textContent = issue.message;
      item.append(label, message);
      container.appendChild(item);
    }
  },

  /** The row key is what ties a station to a template row. */
  keyOf(entry) {
    return `${normalize(entry.lokacija)}|${normalize(entry.vodomjer)}`;
  },

  render() {
    const body = $("stationRows");
    const needle = normalize($("stationSearch").value);
    body.textContent = "";

    const validKeys = new Set(this.rows.map((row) => `${normalize(row.lokacija)}|${normalize(row.vodomjer)}`));
    const counts = new Map();
    for (const entry of this.entries) {
      const key = this.keyOf(entry);
      counts.set(key, (counts.get(key) || 0) + 1);
    }

    let shown = 0;
    this.entries.forEach((entry, index) => {
      const haystack = normalize(`${entry.lokacija} ${entry.vodomjer} ${entry.uredjaj}`);
      if (needle && !haystack.includes(needle)) return;
      shown += 1;

      const key = this.keyOf(entry);
      const invalid = !validKeys.has(key) || counts.get(key) > 1 || !entry.uredjaj;

      const tr = document.createElement("tr");
      tr.className = `${invalid ? "is-invalid " : ""}${entry.enabled ? "" : "is-off"}`.trim();

      const number = document.createElement("td");
      number.className = "col-num";
      number.textContent = String(index + 1);

      const rowCell = document.createElement("td");
      rowCell.appendChild(this.rowSelect(entry, index));

      const deviceCell = document.createElement("td");
      deviceCell.appendChild(this.deviceInput(entry, index));

      const toggleCell = document.createElement("td");
      toggleCell.className = "col-toggle";
      const toggle = document.createElement("input");
      toggle.type = "checkbox";
      toggle.checked = entry.enabled;
      toggle.title = "Uključi ili isključi ovu stanicu iz obračuna";
      toggle.addEventListener("change", () => {
        this.entries[index].enabled = toggle.checked;
        this.setDirty(true);
        this.render();
      });
      toggleCell.appendChild(toggle);

      const actionCell = document.createElement("td");
      actionCell.className = "col-actions";
      const remove = document.createElement("button");
      remove.className = "row-delete";
      remove.textContent = "×";
      remove.title = "Obriši stanicu";
      remove.addEventListener("click", () => {
        this.entries.splice(index, 1);
        this.setDirty(true);
        this.render();
      });
      actionCell.appendChild(remove);

      tr.append(number, rowCell, deviceCell, toggleCell, actionCell);
      body.appendChild(tr);
    });

    $("emptyState").hidden = shown > 0 || this.entries.length === 0;
  },

  rowSelect(entry, index) {
    const select = document.createElement("select");
    const key = this.keyOf(entry);

    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "— izaberi red —";
    select.appendChild(placeholder);

    let matched = false;
    for (const row of this.rows) {
      const option = document.createElement("option");
      option.value = String(row.row);
      option.textContent = `${row.row}. ${row.label}`;
      if (`${normalize(row.lokacija)}|${normalize(row.vodomjer)}` === key) {
        option.selected = true;
        matched = true;
      }
      select.appendChild(option);
    }

    if (!matched && entry.lokacija) {
      // Keep an unknown pairing visible instead of silently reassigning it.
      const orphan = document.createElement("option");
      orphan.value = "";
      orphan.selected = true;
      orphan.textContent = `⚠ ${entry.lokacija} / ${entry.vodomjer} — nema u template-u`;
      select.insertBefore(orphan, select.firstChild.nextSibling);
    }

    select.addEventListener("change", () => {
      const row = this.rows.find((candidate) => String(candidate.row) === select.value);
      this.entries[index].lokacija = row ? row.lokacija : "";
      this.entries[index].vodomjer = row ? row.vodomjer : "";
      this.setDirty(true);
      this.render();
    });
    return select;
  },

  deviceInput(entry, index) {
    const input = document.createElement("input");
    input.type = "text";
    input.value = entry.uredjaj || "";
    input.placeholder = "npr. Bajer 1 - U";
    input.addEventListener("input", () => {
      this.entries[index].uredjaj = input.value;
      this.setDirty(true);
    });
    return input;
  },

  add() {
    const used = new Set(this.entries.map((entry) => this.keyOf(entry)));
    const free = this.rows.find((row) => !used.has(`${normalize(row.lokacija)}|${normalize(row.vodomjer)}`));
    this.entries.push({
      lokacija: free ? free.lokacija : "",
      vodomjer: free ? free.vodomjer : "",
      uredjaj: "",
      enabled: true,
    });
    this.setDirty(true);
    $("stationSearch").value = "";
    this.render();
    const inputs = $("stationRows").querySelectorAll('input[type="text"]');
    if (inputs.length) inputs[inputs.length - 1].focus();
  },

  async save() {
    notify($("stationMessage"), "");
    try {
      const payload = await api.post("/api/stations", { stations: this.entries });
      this.rows = payload.rows;
      this.entries = payload.stations.map((s) => ({ ...s, enabled: s.enabled !== false }));
      this.setDirty(false);
      this.renderIssues(payload.issues);
      this.render();
      notify($("stationMessage"), "Lista stanica je sačuvana.", "ok");
    } catch (error) {
      notify($("stationMessage"), error.message, "error");
    }
  },

};

function normalize(text) {
  return (text || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

/* Boot --------------------------------------------------------------------- */

function initTabs() {
  for (const tab of document.querySelectorAll(".tab")) {
    tab.addEventListener("click", () => {
      for (const other of document.querySelectorAll(".tab")) other.classList.toggle("is-active", other === tab);
      $("view-run").hidden = tab.dataset.view !== "run";
      $("view-stations").hidden = tab.dataset.view !== "stations";
    });
  }
}

(async function boot() {
  initTheme();
  if (!(await initSession())) return;
  $("app").hidden = false;
  initTabs();

  const bootstrap = await api.get("/api/bootstrap");
  await run.init(bootstrap);
  await stations.init();

  window.addEventListener("beforeunload", (event) => {
    if (stations.dirty) event.preventDefault();
  });
})();
