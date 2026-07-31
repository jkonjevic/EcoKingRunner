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

/* Run view ----------------------------------------------------------------- */

const run = {
  cursor: 0,
  timer: null,
  bootstrap: null,

  async init(bootstrap) {
    this.bootstrap = bootstrap;
    $("selectedDate").value = bootstrap.defaultDate;
    $("selectedDate").max = bootstrap.today;
    $("modeBadge").textContent = bootstrap.cloud ? "Online" : "Lokalno";
    $("dateHint").textContent =
      `${bootstrap.stationCount} aktivnih stanica · template: ${bootstrap.template}`;

    // A hosted container has no screen and no Excel to open into.
    $("browserVisibleRow").hidden = bootstrap.cloud;
    $("openAfterRow").hidden = bootstrap.cloud;

    $("runBtn").addEventListener("click", () => this.start());
    $("stopBtn").addEventListener("click", () => this.stop());
    $("copyLog").addEventListener("click", () => navigator.clipboard.writeText($("log").textContent));

    await this.refreshReports();
    this.poll();
    this.timer = setInterval(() => this.poll(), 1000);
  },

  config() {
    return {
      selectedDate: $("selectedDate").value,
      workers: Number($("workers").value || 1),
      limit: $("limit").value,
      chartWait: Number($("chartWait").value || 5000),
      searchWait: Number($("searchWait").value || 2000),
      browserVisible: $("browserVisible").checked,
      verbose: $("verbose").checked,
    };
  },

  async start() {
    notify($("runMessage"), "");
    try {
      await api.post("/api/run", this.config());
      this.cursor = 0;
      $("log").textContent = "";
      this.poll();
    } catch (error) {
      notify($("runMessage"), error.message, "error");
    }
  },

  async stop() {
    await api.post("/api/stop").catch(() => {});
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
    if (state.lines.length) {
      const console = $("log");
      const pinned = $("autoScroll").checked;
      console.textContent += state.lines.join("\n") + "\n";
      if (pinned) console.scrollTop = console.scrollHeight;
    }

    const wasRunning = this.running;
    this.running = state.running;
    $("runBtn").disabled = state.running;
    $("stopBtn").disabled = !state.running;

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

    $("statProgress").textContent = `${state.done} / ${state.total}`;
    $("statCurrent").textContent = state.current || "—";
    $("statCurrent").title = state.current || "";
    $("statStarted").textContent = state.startedAt || "—";
    $("progressBar").style.width = state.total
      ? `${Math.min(100, Math.round((state.done * 100) / state.total))}%`
      : "0%";

    if (wasRunning && !state.running) await this.onFinished(state);
  },

  async onFinished(state) {
    await this.refreshReports();
    if (state.returnCode !== 0) {
      notify($("runMessage"), "Obračun nije uspio. Provjeri log.", "error");
      return;
    }
    notify($("runMessage"), "Izvještaj je spreman.", "ok");
    if (!this.bootstrap.cloud && $("openAfter").checked) {
      await api.post("/api/open-report", { selectedDate: state.selectedDate }).catch(() => {});
    }
  },

  async refreshReports() {
    const container = $("reportList");
    const { reports } = await api.get("/api/reports").catch(() => ({ reports: [] }));
    container.textContent = "";
    if (!reports.length) {
      const empty = document.createElement("p");
      empty.className = "report-empty";
      empty.textContent = "Još nema generisanih izvještaja.";
      container.appendChild(empty);
      return;
    }
    for (const report of reports.slice(0, 8)) {
      const row = document.createElement("div");
      row.className = "report-row";

      const date = document.createElement("b");
      date.textContent = report.date;
      const meta = document.createElement("small");
      meta.textContent = report.modified;
      const spacer = document.createElement("span");
      spacer.className = "spacer";

      const download = document.createElement("a");
      download.className = "btn btn-ghost btn-sm";
      download.href = `/api/report?date=${encodeURIComponent(report.date)}`;
      download.textContent = "Preuzmi";

      row.append(date, meta, spacer, download);
      container.appendChild(row);
    }
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
