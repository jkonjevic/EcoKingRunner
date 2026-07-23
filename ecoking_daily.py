from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Iterable

from dotenv import load_dotenv
from openpyxl import load_workbook
from playwright.sync_api import Browser, Error as PlaywrightError, Page, sync_playwright


DATE_FORMAT = "%d.%m.%Y."
DEFAULT_WORKBOOK = "EcoKing - tabela potrošnje - Jul 2026. TEST.xlsx"
DEFAULT_LOCATION_MAP = "herceg_novi_stations.json"


@dataclass(frozen=True)
class Measurement:
    daily_m3: float | None
    max_daily_m3: float | None
    min_daily_m3: float | None

    @property
    def daily_lps(self) -> float | None:
        return self.daily_m3 * 1000 / 86400 if self.daily_m3 is not None else None

    @property
    def max_daily_lps(self) -> float | None:
        return self.max_daily_m3 * 1000 / 900 if self.max_daily_m3 is not None else None

    @property
    def min_daily_lps(self) -> float | None:
        return self.min_daily_m3 * 1000 / 900 if self.min_daily_m3 is not None else None


@dataclass(frozen=True)
class LocationRow:
    row: int
    location: str
    meter_type: str | None = None
    effective_location: str | None = None


@dataclass(frozen=True)
class StationJob:
    station_key: str
    search_value: str
    excel_row: int | None
    excel_location: str | None
    meter_type: str | None


@dataclass(frozen=True)
class RunIssue:
    station_key: str
    search_value: str
    excel_row: int | None
    excel_location: str | None
    meter_type: str | None
    reason: str


@dataclass(frozen=True)
class RunResult:
    measurements: dict[int, Measurement]
    successes: list[StationJob]
    failures: list[RunIssue]
    no_data: list[RunIssue]


def env(name: str, default: str | None = None) -> str | None:
    return os.getenv(name) or os.getenv(name.lower()) or default


def as_bool(value: str | bool | None, default: bool = False) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


def configure_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    use_color = not as_bool(os.getenv("NO_COLOR"), default=False)

    class ColorFormatter(logging.Formatter):
        COLORS = {
            logging.DEBUG: "\033[36m",
            logging.INFO: "\033[32m",
            logging.WARNING: "\033[33m",
            logging.ERROR: "\033[31m",
            logging.CRITICAL: "\033[1;31m",
        }
        RESET = "\033[0m"

        def format(self, record: logging.LogRecord) -> str:
            if use_color:
                color = self.COLORS.get(record.levelno, "")
                record.levelname = f"{color}{record.levelname}{self.RESET}"
                record.msg = f"{color}{record.msg}{self.RESET}" if record.levelno >= logging.WARNING else record.msg
            return super().format(record)

    handler = logging.StreamHandler()
    handler.setFormatter(ColorFormatter("%(asctime)s %(levelname)-17s %(message)s", datefmt="%H:%M:%S"))
    root = logging.getLogger()
    root.handlers.clear()
    root.setLevel(level)
    root.addHandler(handler)


def parse_date_sheet_name(name: str) -> datetime | None:
    try:
        return datetime.strptime(name.strip(), DATE_FORMAT)
    except ValueError:
        return None


def yesterday_date() -> datetime:
    return datetime.now() - timedelta(days=1)


def newest_date_sheet(workbook: Any) -> Any:
    dated = []
    for ws in workbook.worksheets:
        parsed = parse_date_sheet_name(ws.title)
        if parsed:
            dated.append((parsed, ws))
    if not dated:
        raise RuntimeError("No date-named sheets found in workbook.")
    return max(dated, key=lambda item: item[0])[1]


def newest_date_sheet_before(workbook: Any, target_date: datetime) -> Any:
    dated = []
    for ws in workbook.worksheets:
        parsed = parse_date_sheet_name(ws.title)
        if parsed and parsed < target_date:
            dated.append((parsed, ws))
    if dated:
        return max(dated, key=lambda item: item[0])[1]
    return newest_date_sheet(workbook)


def load_location_rows(workbook_path: Path) -> list[LocationRow]:
    wb = load_workbook(workbook_path, data_only=False)
    ws = newest_date_sheet(wb)
    header_by_col = {}
    for cell in ws[1]:
        if cell.value:
            header_by_col[str(cell.value).strip().upper()] = cell.column
    location_col = header_by_col.get("LOKACIJA")
    meter_col = header_by_col.get("VODOMJER")
    if not location_col:
        raise RuntimeError("Could not find LOKACIJA header in the newest sheet.")

    rows: list[LocationRow] = []
    last_location: str | None = None
    for row in range(2, ws.max_row + 1):
        value = ws.cell(row=row, column=location_col).value
        location = str(value).strip() if value is not None else ""
        if location:
            last_location = location
            meter_value = ws.cell(row=row, column=meter_col).value if meter_col else None
            meter_type = str(meter_value).strip() if meter_value else None
            rows.append(LocationRow(row=row, location=location, meter_type=meter_type, effective_location=location))
        elif last_location:
            meter_value = ws.cell(row=row, column=meter_col).value if meter_col else None
            meter_type = str(meter_value).strip() if meter_value else None
            if meter_type:
                rows.append(LocationRow(row=row, location="", meter_type=meter_type, effective_location=last_location))
    wb.close()
    return rows


def duplicate_daily_sheet(
    workbook_path: Path,
    target_date: datetime,
    measurements: dict[int, Measurement],
    replace_existing_sheet: bool = True,
    clear_existing_values: bool = True,
    clear_new_sheet_values: bool = True,
) -> Path:
    wb = load_workbook(workbook_path)
    title = target_date.strftime(DATE_FORMAT)
    if title in wb.sheetnames and replace_existing_sheet:
        existing = wb[title]
        logging.info("Sheet %s already exists. Replacing only that sheet.", title)
        wb.remove(existing)

    created_sheet = False
    if title in wb.sheetnames:
        ws = wb[title]
        logging.info("Sheet %s already exists. Updating only scraped rows.", title)
    else:
        template = newest_date_sheet_before(wb, target_date)
        ws = wb.copy_worksheet(template)
        ws.title = title
        ws["N1"] = f"DATUM: {title}"
        created_sheet = True

        # Match the source workbook convention: newest sheet first.
        wb._sheets.remove(ws)
        wb._sheets.insert(0, ws)

    if (created_sheet and clear_new_sheet_values) or (not created_sheet and clear_existing_values):
        # Keep identity columns A-E, clear all measurement/note columns, then fill F/H/J m3 values.
        for row in range(2, ws.max_row + 1):
            for column in range(6, ws.max_column + 1):
                ws.cell(row=row, column=column).value = None

    for row, measurement in measurements.items():
        ws.cell(row=row, column=6).value = measurement.daily_m3
        ws.cell(row=row, column=7).value = measurement.daily_lps
        ws.cell(row=row, column=8).value = measurement.max_daily_m3
        ws.cell(row=row, column=9).value = measurement.max_daily_lps
        ws.cell(row=row, column=10).value = measurement.min_daily_m3
        ws.cell(row=row, column=11).value = measurement.min_daily_lps
        logging.info("INGESTED the row into EXCEL, row number %s.", row)

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except AttributeError:
        logging.debug("Workbook calculation flags are not available in this openpyxl version.")
    wb.save(workbook_path)
    wb.close()
    return workbook_path


def load_location_map(path: Path | None) -> dict[str, str]:
    if not path or not path.exists():
        logging.warning("Location map file was not found. Browser searches will use workbook LOKACIJA values.")
        return {}
    with path.open("r", encoding="utf-8") as fh:
        raw = json.load(fh)
    location_map = {str(k).strip(): str(v).strip() for k, v in raw.items() if str(k).strip()}
    logging.info("Loaded %s location mappings from %s", len(location_map), path)
    return location_map


def normalize_lookup(text: str) -> str:
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower()
    text = re.sub(r"\s+-\s+[iu]\s*$", "", text, flags=re.IGNORECASE)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def strip_lookup_direction(text: str) -> str:
    return re.sub(r"\s+[iu]$", "", normalize_lookup(text)).strip()


def preferred_station_suffix(meter_type: str | None) -> str | None:
    normalized = normalize_lookup(meter_type or "")
    if normalized.startswith("ulaz"):
        return "u"
    if normalized.startswith("izlaz"):
        return "i"
    return None


def key_direction_suffix(text: str) -> str | None:
    match = re.search(r"\s+-\s+([iu])\s*$", text, flags=re.IGNORECASE)
    return match.group(1).lower() if match else None


def lookup_tokens(text: str) -> set[str]:
    ignored = {"fs", "ps", "rezervoar", "ulaz", "izlaz"}
    return {
        token
        for token in normalize_lookup(text).split()
        if (len(token) >= 4 or token.isdigit()) and token not in ignored
    }


def score_station_key_for_row(station_key: str, row: LocationRow) -> int:
    row_location = row.effective_location or row.location
    if not row_location:
        return 0

    station_norm = normalize_lookup(station_key)
    station_without_direction = strip_lookup_direction(station_key)
    row_norm = normalize_lookup(row_location)
    row_without_direction = strip_lookup_direction(row_location)
    station_tokens = lookup_tokens(station_key)
    row_tokens = lookup_tokens(f"{row_location} {row.meter_type or ''}")
    meter_tokens = lookup_tokens(row.meter_type or "")
    station_suffix = key_direction_suffix(station_key)
    row_suffix = preferred_station_suffix(row.meter_type)

    score = 0
    if station_norm == row_norm:
        score += 100
    if station_without_direction == row_without_direction:
        score += 90
    elif row_without_direction and row_without_direction in station_without_direction:
        score += 50
    overlap = row_tokens & station_tokens
    if row_tokens and row_tokens.issubset(station_tokens):
        score += 60 + (5 * len(row_tokens))
    elif overlap:
        score += 15 + (10 * len(overlap))
    meter_overlap = meter_tokens & station_tokens
    if meter_overlap:
        score += 30 + (10 * len(meter_overlap))
    station_numbers = {token for token in station_tokens if token.isdigit()}
    row_numbers = {token for token in row_tokens if token.isdigit()}
    if station_numbers and row_numbers and not (station_numbers & row_numbers):
        score -= 50
    if row_suffix and station_suffix == row_suffix:
        score += 20
    if row_suffix and station_suffix and station_suffix != row_suffix:
        score -= 10
    return score


def exact_excel_row_for_station(station_key: str, rows: list[LocationRow]) -> LocationRow | None:
    station_suffix = key_direction_suffix(station_key)
    if station_suffix not in {"u", "i"}:
        return None

    station_location = re.sub(r"\s+-\s+[iu]\s*$", "", station_key, flags=re.IGNORECASE).strip()
    station_location_norm = normalize_lookup(station_location)
    required_meter = "ulaz" if station_suffix == "u" else "izlaz"

    matches = [
        row
        for row in rows
        if normalize_lookup(row.effective_location or row.location) == station_location_norm
        and normalize_lookup(row.meter_type or "") == required_meter
    ]
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        logging.warning(
            "Multiple exact Excel rows match station key %r: %s. Using row %s.",
            station_key,
            [(row.row, row.effective_location or row.location, row.meter_type) for row in matches],
            matches[0].row,
        )
        return matches[0]
    return None


def resolve_excel_row_for_station(station_key: str, rows: list[LocationRow]) -> LocationRow | None:
    exact_row = exact_excel_row_for_station(station_key, rows)
    if exact_row:
        logging.debug(
            "Exact row match for %r -> row %s (%s / %s)",
            station_key,
            exact_row.row,
            exact_row.effective_location or exact_row.location,
            exact_row.meter_type,
        )
        return exact_row

    scored = [(score_station_key_for_row(station_key, row), row) for row in rows]
    scored = [(score, row) for score, row in scored if score > 0]
    if not scored:
        return None
    scored.sort(key=lambda item: item[0], reverse=True)
    best_score = scored[0][0]
    tied = [row for score, row in scored if score == best_score]
    if len(tied) > 1:
        logging.warning(
            "Multiple Excel rows tie for station key %r: %s. Using row %s.",
            station_key,
            [(row.row, row.effective_location or row.location, row.meter_type) for row in tied],
            tied[0].row,
        )
    return tied[0]


def build_station_jobs(location_map: dict[str, str], rows: list[LocationRow]) -> list[StationJob]:
    jobs: list[StationJob] = []
    for station_key, search_value in location_map.items():
        if not valid_site_location(search_value):
            logging.warning("Skipping station %r because its search value is invalid: %r", station_key, search_value)
            continue
        excel_row = resolve_excel_row_for_station(station_key, rows)
        if not excel_row:
            logging.warning("No Excel row matched station key %r. It will not be scraped because there is nowhere to write it.", station_key)
            continue
        jobs.append(
            StationJob(
                station_key=station_key,
                search_value=search_value,
                excel_row=excel_row.row,
                excel_location=excel_row.effective_location or excel_row.location,
                meter_type=excel_row.meter_type,
            )
        )
    return jobs


def valid_site_location(value: str) -> bool:
    normalized = normalize_lookup(value)
    invalid = {"home", "device list", "devices", "logout", "login", "menu"}
    return bool(normalized) and normalized not in invalid


def resolve_site_location(workbook_location: str, location_map: dict[str, str], meter_type: str | None = None) -> str:
    mapped = location_map.get(workbook_location)
    if mapped and valid_site_location(mapped):
        return mapped
    if mapped:
        logging.warning("Ignoring invalid JSON mapping for %r: %r", workbook_location, mapped)

    workbook_norm = normalize_lookup(workbook_location)
    workbook_without_direction = strip_lookup_direction(workbook_location)
    preferred_suffix = preferred_station_suffix(meter_type)
    workbook_tokens = lookup_tokens(workbook_location)
    candidates: list[tuple[int, str, str]] = []
    for key, value in location_map.items():
        if not valid_site_location(value):
            continue
        key_norm = normalize_lookup(key)
        key_without_direction = strip_lookup_direction(key)
        key_suffix_match = re.search(r"\s+-\s+([iu])\s*$", key, flags=re.IGNORECASE)
        key_suffix = key_suffix_match.group(1).lower() if key_suffix_match else None
        key_tokens = lookup_tokens(key)

        score = 0
        if key_norm == workbook_norm:
            score += 100
        if key_without_direction == workbook_without_direction:
            score += 90
        if workbook_tokens and workbook_tokens.issubset(key_tokens):
            score += 60
        elif workbook_tokens and workbook_tokens & key_tokens:
            score += 25
        if preferred_suffix and key_suffix == preferred_suffix:
            score += 15
        if preferred_suffix and key_suffix and key_suffix != preferred_suffix:
            score -= 10
        if score > 0:
            candidates.append((score, key, value))

    candidates.sort(reverse=True)
    if candidates:
        score, key, value = candidates[0]
        tied = [candidate_key for candidate_score, candidate_key, _ in candidates if candidate_score == score]
        if len(tied) > 1:
            logging.warning("Multiple JSON keys tie for %r (%s): %s. Using %r.", workbook_location, meter_type, tied, key)
        else:
            logging.info("Resolved %r (%s) through JSON key %r", workbook_location, meter_type, key)
        return value

    logging.warning("No non-empty JSON mapping for %r. Falling back to workbook value.", workbook_location)
    return workbook_location


def click_first_visible(page: Page, selectors: Iterable[str], label: str, timeout_ms: int = 10_000) -> bool:
    for selector in selectors:
        if not selector:
            continue
        locator = page.locator(selector).first
        try:
            locator.wait_for(state="visible", timeout=timeout_ms)
            logging.debug("Clicking %s with selector %s", label, selector)
            locator.click(timeout=timeout_ms)
            return True
        except PlaywrightError:
            continue
    return False


def fill_first_visible(page: Page, selectors: Iterable[str], value: str, label: str, timeout_ms: int = 10_000) -> bool:
    for selector in selectors:
        if not selector:
            continue
        locator = page.locator(selector).first
        try:
            locator.wait_for(state="visible", timeout=timeout_ms)
            logging.debug("Filling %s with selector %s", label, selector)
            locator.fill(value, timeout=timeout_ms)
            return True
        except PlaywrightError:
            continue
    return False


def fill_open_dropdown_search(page: Page, selectors: Iterable[str], value: str) -> bool:
    selector_list = [selector for selector in selectors if selector]
    filled_selector = page.evaluate(
        """
        ({selectors, value}) => {
          const visible = el => {
            if (!el) return false;
            const style = getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.visibility !== 'hidden'
              && style.display !== 'none'
              && rect.width > 0
              && rect.height > 0;
          };
          const setValue = (el, text) => {
            el.focus();
            const proto = el instanceof HTMLTextAreaElement
              ? window.HTMLTextAreaElement.prototype
              : window.HTMLInputElement.prototype;
            const valueSetter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
            if (valueSetter) {
              valueSetter.call(el, '');
              el.dispatchEvent(new Event('input', {bubbles: true}));
              valueSetter.call(el, text);
            } else {
              el.value = text;
            }
            el.dispatchEvent(new Event('input', {bubbles: true}));
            el.dispatchEvent(new Event('change', {bubbles: true}));
            el.dispatchEvent(new KeyboardEvent('keyup', {bubbles: true, key: text.slice(-1) || ' '}));
          };

          for (const selector of selectors) {
            let elements = [];
            try {
              elements = Array.from(document.querySelectorAll(selector));
            } catch {
              continue;
            }
            const input = elements.find(el =>
              visible(el)
              && (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement)
            );
            if (input) {
              setValue(input, value);
              return selector;
            }
          }
          return null;
        }
        """,
        {"selectors": selector_list, "value": value},
    )
    if filled_selector:
        logging.debug("Filling location search with visible dropdown input %s", filled_selector)
        return True

    try:
        active_tag = page.evaluate("document.activeElement ? document.activeElement.tagName.toLowerCase() : ''")
        if active_tag in {"input", "textarea"}:
            page.keyboard.press("Control+A")
            page.keyboard.type(value, delay=5)
            logging.debug("Typed location search into focused %s element", active_tag)
            return True
    except PlaywrightError:
        pass

    return False


def type_open_dropdown_search(page: Page, selectors: Iterable[str], value: str) -> bool:
    selector_list = [selector for selector in selectors if selector]
    focused_selector = page.evaluate(
        """
        (selectors) => {
          const visible = el => {
            if (!el) return false;
            const style = getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.visibility !== 'hidden'
              && style.display !== 'none'
              && rect.width > 0
              && rect.height > 0;
          };
          for (const selector of selectors) {
            let elements = [];
            try {
              elements = Array.from(document.querySelectorAll(selector));
            } catch {
              continue;
            }
            const input = elements.find(el =>
              visible(el)
              && (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement)
            );
            if (input) {
              input.focus();
              return selector;
            }
          }
          return null;
        }
        """,
        selector_list,
    )
    if not focused_selector:
        return False

    type_delay_ms = int(env("SEARCH_TYPE_DELAY_MS", "10") or "10")
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    page.keyboard.type(value, delay=type_delay_ms)
    logging.debug("Retyped location search with keyboard into %s", focused_selector)
    return True


def search_queries_for_location(location: str) -> list[str]:
    queries: list[str] = []
    for query in [location, *(re.findall(r"\d{8,}", location) or [])]:
        query = query.strip()
        if query and query not in queries:
            queries.append(query)
    return queries


def get_dropdown_search_result_state(page: Page, search_value: str) -> dict[str, Any]:
    state = page.evaluate(
        """
        (searchValue) => {
          const normalize = text => (text || '')
            .normalize('NFD')
            .replace(/[\\u0300-\\u036f]/g, '')
            .toLowerCase()
            .replace(/[^a-z0-9]+/g, ' ')
            .trim();
          const target = normalize(searchValue);
          const visible = el => {
            if (!el) return false;
            const style = getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.visibility !== 'hidden'
              && style.display !== 'none'
              && rect.width > 0
              && rect.height > 0;
          };
          const stationLike = text => {
            const normalizedText = normalize(text);
            return /\\d{8,}/.test(text)
              && /herceg\\s+novi/i.test(text)
              && !/\\b(home|device\\s+list|user\\s+settings|log\\s*out|menu)\\b/i.test(normalizedText);
          };
          const optionSelector = ".dropdown-menu a, .dropdown-menu button, [role=listbox] [role=option], [role=option], option";
          const rawOptions = Array.from(document.querySelectorAll(optionSelector))
            .filter(visible)
            .map(option => {
              const text = (option.innerText || option.textContent || '').replace(/\\s+/g, ' ').trim();
              return {text, normalized: normalize(text)};
            })
            .filter(option =>
              option.text
              && stationLike(option.text)
              && !/no\\s+(results?|matches?)|not\\s+found|nema\\s+rezultata/i.test(option.text)
            );
          const uniqueByText = (items) => {
            const seen = new Set();
            const unique = [];
            for (const item of items) {
              const key = item.normalized;
              if (seen.has(key)) continue;
              seen.add(key);
              unique.push({...item, index: unique.length});
            }
            return unique;
          };
          const options = uniqueByText(rawOptions);
          const matches = options.filter(option =>
            option.normalized === target
            || option.normalized.includes(target)
            || target.includes(option.normalized)
          );
          return {options, matches};
        }
        """,
        search_value,
    )
    return dict(state or {"options": [], "matches": []})


def click_dropdown_option_by_index(page: Page, index: int) -> None:
    page.evaluate(
        """
        (targetIndex) => {
          const visible = el => {
            if (!el) return false;
            const style = getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.visibility !== 'hidden'
              && style.display !== 'none'
              && rect.width > 0
              && rect.height > 0;
          };
          const normalize = text => (text || '')
            .normalize('NFD')
            .replace(/[\\u0300-\\u036f]/g, '')
            .toLowerCase()
            .replace(/[^a-z0-9]+/g, ' ')
            .trim();
          const stationLike = text => {
            const normalizedText = normalize(text);
            return /\\d{8,}/.test(text)
              && /herceg\\s+novi/i.test(text)
              && !/\\b(home|device\\s+list|user\\s+settings|log\\s*out|menu)\\b/i.test(normalizedText);
          };
          const optionSelector = ".dropdown-menu a, .dropdown-menu button, [role=listbox] [role=option], [role=option], option";
          const rawOptions = Array.from(document.querySelectorAll(optionSelector))
            .filter(visible)
            .filter(option => stationLike((option.innerText || option.textContent || '').replace(/\\s+/g, ' ').trim()));
          const seen = new Set();
          const options = [];
          for (const option of rawOptions) {
            const key = normalize(option.innerText || option.textContent || '');
            if (seen.has(key)) continue;
            seen.add(key);
            options.push(option);
          }
          const option = options[targetIndex];
          if (!option) throw new Error(`Dropdown option index ${targetIndex} is no longer visible`);
          option.click();
        }
        """,
        index,
    )


def choose_single_filtered_dropdown_result(page: Page, search_value: str, wait_ms: int | None = None) -> None:
    wait_ms = wait_ms if wait_ms is not None else int(env("SEARCH_RESULTS_WAIT_MS", "2000") or "2000")
    deadline = time.monotonic() + (wait_ms / 1000)
    state: dict[str, Any] = {"options": [], "matches": []}

    while True:
        state = get_dropdown_search_result_state(page, search_value)
        options = state.get("options", [])
        matches = state.get("matches", [])
        if options or matches or time.monotonic() >= deadline:
            break
        page.wait_for_timeout(150)

    options = state.get("options", [])
    matches = state.get("matches", [])
    logging.debug("Filtered dropdown options for %r: %s", search_value, [item.get("text") for item in options])

    if len(matches) == 1:
        match = matches[0]
        logging.info("Choosing dropdown result: %s", match.get("text"))
        click_dropdown_option_by_index(page, int(match["index"]))
        page.wait_for_timeout(250)
        return

    if len(options) == 1:
        option = options[0]
        logging.info("Choosing only visible dropdown result: %s", option.get("text"))
        click_dropdown_option_by_index(page, int(option["index"]))
        page.wait_for_timeout(250)
        return

    if not options:
        raise RuntimeError(f"No dropdown results after searching {search_value!r}.")

    raise RuntimeError(
        f"Expected exactly one dropdown result for {search_value!r}, got {len(options)}: "
        f"{[item.get('text') for item in options[:10]]}"
    )


def click_interval_option(page: Page, label: str) -> bool:
    clicked = page.evaluate(
        """
        (label) => {
          const visible = el => {
            if (!el) return false;
            const style = getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.visibility !== 'hidden'
              && style.display !== 'none'
              && rect.width > 0
              && rect.height > 0;
          };
          const normalize = text => (text || '').replace(/\\s+/g, ' ').trim();
          const candidates = Array.from(document.querySelectorAll("button, a, [role=button], .dropdown-menu a"))
            .filter(visible)
            .filter(el => normalize(el.innerText || el.textContent) === label)
            .map(el => {
              let score = 0;
              const attrs = Array.from(el.attributes || []).map(attr => `${attr.name}=${attr.value}`).join(' ');
              const cls = String(el.className || '');
              if (/changeTimeSlice/i.test(attrs)) score += 100;
              if (el.tagName.toLowerCase() === 'button') score += 20;
              if (/dropdown-toggle/i.test(cls) || /uib-dropdown-toggle/i.test(attrs)) score -= 50;
              if (/btn-primary/i.test(cls)) score += 10;
              return {el, score};
            })
            .sort((a, b) => b.score - a.score);
          if (!candidates.length) return false;
          candidates[0].el.click();
          return true;
        }
        """,
        label,
    )
    if clicked:
        logging.debug("Clicked interval option %r", label)
    return bool(clicked)


def click_device_dropdown(page: Page) -> bool:
    configured = env("SEARCH_TOGGLE_SELECTOR")
    if configured:
        return click_first_visible(page, [configured], "configured device dropdown")

    clicked_precise = page.evaluate(
        """
        (() => {
          const visible = el => {
            if (!el) return false;
            const style = getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.visibility !== 'hidden'
              && style.display !== 'none'
              && rect.width > 0
              && rect.height > 0;
          };
          const candidates = Array.from(document.querySelectorAll("button.btn-primary.dropdown-toggle, .btn-primary.dropdown-toggle"))
            .filter(visible)
            .map(el => ({el, rect: el.getBoundingClientRect()}))
            .filter(item =>
              item.rect.left >= 0
              && item.rect.left < 90
              && item.rect.top > 40
              && item.rect.top < 160
              && item.rect.width >= 20
              && item.rect.width <= 70
              && item.rect.height >= 20
              && item.rect.height <= 55
            )
            .sort((a, b) => (a.rect.left - b.rect.left) || (a.rect.top - b.rect.top));
          if (!candidates.length) return false;
          candidates[0].el.click();
          return true;
        })()
        """
    )
    if clicked_precise:
        logging.debug("Clicked device dropdown using precise top-left selector")
        return True

    clicked = page.evaluate(
        """
        (() => {
          const visible = el => {
            if (!el) return false;
            const style = getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.visibility !== 'hidden'
              && style.display !== 'none'
              && rect.width > 0
              && rect.height > 0;
          };
          const score = el => {
            const rect = el.getBoundingClientRect();
            const text = (el.innerText || el.getAttribute('aria-label') || el.getAttribute('title') || '').trim();
            const html = el.innerHTML || '';
            const cls = String(el.className || '');
            let value = 0;
            if (rect.left < 80 && rect.top > 45 && rect.top < 140) value += 100;
            if (rect.width >= 20 && rect.width <= 55 && rect.height >= 20 && rect.height <= 45) value += 30;
            if (/btn-primary|dropdown-toggle|select2-choice|multiselect/i.test(cls)) value += 20;
            if (/caret|angle-down|chevron-down|fa-caret-down|fa-angle-down|▼|▾|⌄|⌵/i.test(html + ' ' + text)) value += 20;
            if (rect.top < 45 || rect.left > 160) value -= 80;
            if (/navbar|sidebar|hamburger|menu|bars|toggle/i.test(cls + ' ' + html + ' ' + text)) value -= 50;
            return value;
          };
          const candidates = Array.from(document.querySelectorAll("button, a, [role=button]"))
            .filter(visible)
            .map(el => ({el, value: score(el)}))
            .filter(item => item.value > 0)
            .sort((a, b) => b.value - a.value);
          if (!candidates.length) return false;
          candidates[0].el.click();
          return true;
        })()
        """
    )
    if clicked:
        logging.debug("Clicked device dropdown using top-left content heuristic")
        return True

    return click_first_visible(
        page,
        [
            "button.btn-primary.dropdown-toggle",
            ".btn-primary.dropdown-toggle",
            "button:has(.fa-angle-down)",
            "button:has(.fa-caret-down)",
            "button:has-text('▼')",
            "button:has-text('▾')",
        ],
        "device dropdown",
    )


def login(page: Page, base_url: str, email: str, password: str) -> None:
    logging.info("Opening %s", base_url)
    page.goto(base_url, wait_until="domcontentloaded")
    page.wait_for_load_state("domcontentloaded", timeout=30_000)

    email_selectors = [
        env("LOGIN_EMAIL_SELECTOR"),
        "input[type='email']",
        "input[name='email']",
        "input[name='username']",
        "#email",
        "#username",
        "input[type='text']",
    ]
    password_selectors = [
        env("LOGIN_PASSWORD_SELECTOR"),
        "input[type='password']",
        "input[name='password']",
        "#password",
    ]
    submit_selectors = [
        env("LOGIN_SUBMIT_SELECTOR"),
        "button[type='submit']",
        "input[type='submit']",
        "button:has-text('Login')",
        "button:has-text('Log in')",
        "button:has-text('Sign in')",
        "button:has-text('Prijava')",
    ]

    if fill_first_visible(page, email_selectors, email, "email"):
        click_first_visible(page, submit_selectors, "login next/submit", timeout_ms=3_000)
        try:
            page.wait_for_timeout(1_000)
        except PlaywrightError:
            pass

    if fill_first_visible(page, password_selectors, password, "password"):
        if not click_first_visible(page, submit_selectors, "login submit", timeout_ms=5_000):
            page.keyboard.press("Enter")
        page.wait_for_timeout(1_000)
        logging.info("Login flow submitted")
        return

    wait_seconds = int(env("WAIT_FOR_LOGIN_SECONDS", "0") or "0")
    if wait_seconds > 0:
        logging.warning("Login fields were not detected. Waiting %s seconds for manual login.", wait_seconds)
        page.wait_for_timeout(wait_seconds * 1000)
        return

    raise RuntimeError("Could not complete login. Add LOGIN_* selectors or set WAIT_FOR_LOGIN_SECONDS.")


def select_location(page: Page, location: str) -> None:
    logging.info("Searching location: %s", location)
    search_selectors = [
        env("SEARCH_INPUT_SELECTOR"),
        ".dropdown-menu input",
        ".dropdown-menu textarea",
        ".dropdown-menu .form-control",
        ".dropdown.open input",
        ".open > .dropdown-menu input",
        ".uib-dropdown-menu input",
        "[role='listbox'] input",
        ".select2-search__field",
        ".select2-search input",
        "input[type='search']",
    ]
    input_wait_ms = int(env("SEARCH_INPUT_WAIT_MS", "500") or "500")
    errors: list[str] = []

    queries = search_queries_for_location(location)
    for query in queries:
        try:
            try:
                page.keyboard.press("Escape")
                page.wait_for_timeout(100)
            except PlaywrightError:
                pass

            if not click_device_dropdown(page):
                raise RuntimeError("Could not find the blue dropdown/search button. Set SEARCH_TOGGLE_SELECTOR.")

            page.wait_for_timeout(150)
            if not fill_open_dropdown_search(page, search_selectors, query):
                if not fill_first_visible(page, search_selectors, query, "location search", timeout_ms=input_wait_ms):
                    raise RuntimeError("Could not find the opened dropdown search input. Set SEARCH_INPUT_SELECTOR.")

            page.wait_for_timeout(250)
            try:
                choose_single_filtered_dropdown_result(page, query)
            except RuntimeError as exc:
                if "No dropdown results" not in str(exc):
                    raise
                logging.warning("No dropdown results for %r after DOM fill. Retrying with keyboard typing.", query)
                if not type_open_dropdown_search(page, search_selectors, query):
                    raise
                page.wait_for_timeout(500)
                choose_single_filtered_dropdown_result(page, query)
            if query != location:
                logging.info("Selected %r by fallback serial search %r", location, query)
            return
        except Exception as exc:
            errors.append(f"{query!r}: {exc}")
            if query != queries[-1]:
                logging.warning("Search query %r failed for %r. Trying fallback query.", query, location)
                continue
            raise RuntimeError(f"Could not select location {location!r}. Attempts: {' | '.join(errors)}") from exc


def wait_for_chart_data(page: Page, timeout_ms: int | None = None) -> None:
    timeout_ms = timeout_ms or int(env("CHART_WAIT_MS", "5000") or "5000")
    page.wait_for_function(
        """
        () => {
          if (window.Highcharts && Array.isArray(window.Highcharts.charts)) {
            const hasHighchartsData = window.Highcharts.charts
              .filter(Boolean)
              .some(chart => (chart.series || []).some(s => (s.points || []).some(p => p.y !== null && p.y !== undefined)));
            if (hasHighchartsData) return true;
          }
          if (window.Chart) {
            const charts = window.Chart.instances
              ? Object.values(window.Chart.instances)
              : (window.Chart.getChart ? Array.from(document.querySelectorAll('canvas')).map(c => window.Chart.getChart(c)).filter(Boolean) : []);
            const hasChartJsData = charts.some(chart =>
              chart.data && (chart.data.datasets || []).some(dataset => (dataset.data || []).length)
            );
            if (hasChartJsData) return true;
          }
          return Array.from(document.querySelectorAll('table tr')).some(tr => tr.querySelectorAll('td,th').length > 1);
        }
        """,
        timeout=timeout_ms,
    )


def select_interval(page: Page, label: str) -> None:
    logging.info("Selecting interval: %s", label)
    configured = env("INTERVAL_SELECTOR")
    if configured:
        control = page.locator(configured).first
        control.wait_for(state="visible", timeout=10_000)
        tag = control.evaluate("el => el.tagName.toLowerCase()")
        if tag == "select":
            control.select_option(label=label)
        else:
            control.click()
            if not click_interval_option(page, label):
                raise RuntimeError(f"Could not choose interval {label!r} using INTERVAL_SELECTOR.")
    else:
        select = page.locator("select").filter(has_text=re.compile(re.escape(label), re.I)).first
        try:
            select.select_option(label=label, timeout=5_000)
        except PlaywrightError:
            if click_interval_option(page, label):
                wait_for_chart_data(page)
                return
            if click_first_visible(page, [f"button:has-text('{label}')"], "interval button", timeout_ms=5_000):
                wait_for_chart_data(page)
                return
            if not click_first_visible(page, ["[role='combobox']", ".dropdown-toggle"], "interval dropdown", timeout_ms=5_000):
                raise RuntimeError("Could not find interval selector. Set INTERVAL_SELECTOR.")
            if not click_interval_option(page, label):
                raise RuntimeError(f"Could not choose interval {label!r}.")

    wait_for_chart_data(page)


def parse_float(value: Any) -> float | None:
    if isinstance(value, (int, float)):
        return float(value)
    if value is None:
        return None
    text = str(value).strip()
    text = re.sub(r"[^\d,.\-]", "", text)
    if not text:
        return None
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def extract_chart_payload(page: Page) -> dict[str, Any]:
    return page.evaluate(
        """
        () => {
          const payload = {highcharts: [], chartjs: [], tables: []};

          if (window.Highcharts && Array.isArray(window.Highcharts.charts)) {
            for (const chart of window.Highcharts.charts.filter(Boolean)) {
              const series = (chart.series || []).map(s => ({
                name: s.name,
                points: (s.points || []).map(p => ({
                  x: p.x,
                  y: p.y,
                  category: p.category,
                  name: p.name
                }))
              }));
              payload.highcharts.push({series});
            }
          }

          if (window.Chart) {
            const charts = window.Chart.instances
              ? Object.values(window.Chart.instances)
              : (window.Chart.getChart ? Array.from(document.querySelectorAll('canvas')).map(c => window.Chart.getChart(c)).filter(Boolean) : []);
            for (const chart of charts) {
              payload.chartjs.push({
                labels: chart.data && chart.data.labels || [],
                datasets: chart.data && chart.data.datasets || []
              });
            }
          }

          for (const table of document.querySelectorAll('table')) {
            const rows = Array.from(table.querySelectorAll('tr')).map(tr =>
              Array.from(tr.querySelectorAll('th,td')).map(td => td.innerText.trim())
            ).filter(row => row.length);
            if (rows.length) payload.tables.push(rows);
          }
          return payload;
        }
        """
    )


def numbers_from_chart_payload(payload: dict[str, Any]) -> list[float]:
    values: list[float] = []
    for chart in payload.get("highcharts", []):
        for series in chart.get("series", []):
            for point in series.get("points", []):
                value = parse_float(point.get("y"))
                if value is not None:
                    values.append(value)
    for chart in payload.get("chartjs", []):
        for dataset in chart.get("datasets", []):
            for raw in dataset.get("data", []):
                value = parse_float(raw.get("y") if isinstance(raw, dict) else raw)
                if value is not None:
                    values.append(value)
    return values


def latest_value_from_payload(payload: dict[str, Any]) -> float | None:
    highcharts_points: list[tuple[float, float]] = []
    for chart in payload.get("highcharts", []):
        for series in chart.get("series", []):
            for point in series.get("points", []):
                y = parse_float(point.get("y"))
                x = parse_float(point.get("x"))
                if y is not None and x is not None:
                    highcharts_points.append((x, y))
    if highcharts_points:
        return max(highcharts_points, key=lambda item: item[0])[1]

    chartjs_values: list[float] = []
    for chart in payload.get("chartjs", []):
        for dataset in chart.get("datasets", []):
            for raw in dataset.get("data", []):
                value = parse_float(raw.get("y") if isinstance(raw, dict) else raw)
                if value is not None:
                    chartjs_values.append(value)
    if chartjs_values:
        return chartjs_values[-1]

    for table in payload.get("tables", []):
        for row in reversed(table):
            values = [parse_float(cell) for cell in row]
            values = [value for value in values if value is not None]
            if values:
                return values[-1]
    return None


def save_debug_artifacts(page: Page, debug_dir: Path, location: str, suffix: str) -> None:
    debug_dir.mkdir(parents=True, exist_ok=True)
    safe_name = re.sub(r"[^A-Za-z0-9_.-]+", "_", location).strip("_") or "location"
    stem = debug_dir / f"{safe_name}_{suffix}_{int(time.time())}"
    page.screenshot(path=str(stem.with_suffix(".png")), full_page=True)
    stem.with_suffix(".html").write_text(page.content(), encoding="utf-8")
    logging.warning("Saved debug artifacts: %s.png and %s.html", stem, stem)


def scrape_measurement(page: Page, location: str, debug_dir: Path) -> Measurement:
    select_location(page, location)

    select_interval(page, env("INTERVAL_1_DAY_LABEL", "1 Day (15-minutely)") or "1 Day (15-minutely)")
    payload_1_day = extract_chart_payload(page)
    values_1_day = numbers_from_chart_payload(payload_1_day)
    if not values_1_day:
        save_debug_artifacts(page, debug_dir, location, "1day_no_values")
        raise RuntimeError(f"No numeric 1-day chart values found for {location!r}.")
    min_daily = min(values_1_day)
    max_daily = max(values_1_day)

    select_interval(page, env("INTERVAL_30_DAYS_LABEL", "30 days") or "30 days")
    payload_30_days = extract_chart_payload(page)
    daily = latest_value_from_payload(payload_30_days)
    if daily is None:
        save_debug_artifacts(page, debug_dir, location, "30days_no_latest")
        raise RuntimeError(f"No latest 30-day value found for {location!r}.")

    logging.info(
        "Read %s: daily=%s m3, max=%s m3, min=%s m3",
        location,
        daily,
        max_daily,
        min_daily,
    )
    return Measurement(daily_m3=daily, max_daily_m3=max_daily, min_daily_m3=min_daily)


def run_browser(
    jobs: list[StationJob],
    headless: bool,
    slow_mo_ms: int,
    keep_browser_open: bool,
    debug_dir: Path,
    limit: int | None,
) -> RunResult:
    base_url = env("URL") or env("url")
    email = env("EMAIL") or env("GMAIL")
    password = env("PASSWORD")
    if not base_url or not email or not password:
        raise RuntimeError(".env must define url, email/gmail, and password.")

    selected_jobs = jobs[:limit] if limit else jobs
    measurements: dict[int, Measurement] = {}
    successes: list[StationJob] = []
    failures: list[RunIssue] = []
    no_data: list[RunIssue] = []

    with sync_playwright() as p:
        launch_kwargs: dict[str, Any] = {"headless": headless, "slow_mo": slow_mo_ms}
        chrome_path = env("CHROME_PATH")
        if chrome_path and Path(chrome_path).exists():
            launch_kwargs["executable_path"] = chrome_path

        logging.info("Launching browser: headless=%s slow_mo_ms=%s", headless, slow_mo_ms)
        browser: Browser = p.chromium.launch(**launch_kwargs)
        context = browser.new_context(viewport={"width": 1440, "height": 950})
        page = context.new_page()
        try:
            login(page, base_url, email, password)
            for idx, job in enumerate(selected_jobs, start=1):
                logging.info(
                    "[%s/%s] Station key=%r, Excel row=%s, Excel LOKACIJA=%r, VODOMJER=%r, browser search=%r",
                    idx,
                    len(selected_jobs),
                    job.station_key,
                    job.excel_row,
                    job.excel_location,
                    job.meter_type,
                    job.search_value,
                )
                try:
                    if job.excel_row is None:
                        logging.warning("Skipping %r because no Excel row was resolved.", job.station_key)
                        failures.append(
                            RunIssue(job.station_key, job.search_value, job.excel_row, job.excel_location, job.meter_type, "No Excel row was resolved")
                        )
                        continue
                    measurement = scrape_measurement(page, job.search_value, debug_dir)
                    measurements[job.excel_row] = measurement
                    successes.append(job)
                    logging.info(
                        'FOUND: MIN: %s MAX: %s DAILY: %s for "%s"',
                        measurement.min_daily_m3,
                        measurement.max_daily_m3,
                        measurement.daily_m3,
                        job.station_key,
                    )
                except Exception:
                    logging.exception("Failed to scrape station %s for Excel row %s", job.station_key, job.excel_row)
                    reason = "See traceback above"
                    try:
                        reason = str(sys.exc_info()[1])
                    except Exception:
                        pass
                    issue = RunIssue(job.station_key, job.search_value, job.excel_row, job.excel_location, job.meter_type, reason)
                    if "No numeric 1-day chart values" in reason or "No latest 30-day value" in reason:
                        no_data.append(issue)
                    else:
                        failures.append(issue)
                    if as_bool(env("CONTINUE_ON_ERROR"), default=True):
                        continue
                    raise
        finally:
            if keep_browser_open and not headless:
                logging.info("KEEP_BROWSER_OPEN is enabled. Press Ctrl+C in this terminal when done inspecting.")
                try:
                    while True:
                        time.sleep(1)
                except KeyboardInterrupt:
                    pass
            context.close()
            browser.close()

    return RunResult(measurements=measurements, successes=successes, failures=failures, no_data=no_data)


def merge_run_results(results: Iterable[RunResult]) -> RunResult:
    measurements: dict[int, Measurement] = {}
    successes: list[StationJob] = []
    failures: list[RunIssue] = []
    no_data: list[RunIssue] = []
    for result in results:
        measurements.update(result.measurements)
        successes.extend(result.successes)
        failures.extend(result.failures)
        no_data.extend(result.no_data)
    return RunResult(measurements=measurements, successes=successes, failures=failures, no_data=no_data)


def split_jobs(jobs: list[StationJob], workers: int) -> list[list[StationJob]]:
    chunks = [[] for _ in range(workers)]
    for idx, job in enumerate(jobs):
        chunks[idx % workers].append(job)
    return [chunk for chunk in chunks if chunk]


def run_browser_parallel(
    jobs: list[StationJob],
    workers: int,
    headless: bool,
    slow_mo_ms: int,
    keep_browser_open: bool,
    debug_dir: Path,
    limit: int | None,
) -> RunResult:
    selected_jobs = jobs[:limit] if limit else jobs
    if workers <= 1 or len(selected_jobs) <= 1:
        return run_browser(
            jobs=selected_jobs,
            headless=headless,
            slow_mo_ms=slow_mo_ms,
            keep_browser_open=keep_browser_open,
            debug_dir=debug_dir,
            limit=None,
        )

    workers = min(workers, len(selected_jobs))
    if not headless:
        logging.warning("Running %s headed browsers in parallel. Use --headless for faster, less noisy parallel runs.", workers)
    if keep_browser_open:
        logging.warning("--keep-browser-open is ignored for parallel runs.")
        keep_browser_open = False

    chunks = split_jobs(selected_jobs, workers)
    logging.info("Running %s jobs across %s browser workers.", len(selected_jobs), len(chunks))
    results: list[RunResult] = []
    with ThreadPoolExecutor(max_workers=len(chunks)) as executor:
        future_by_chunk = {
            executor.submit(
                run_browser,
                jobs=chunk,
                headless=headless,
                slow_mo_ms=slow_mo_ms,
                keep_browser_open=keep_browser_open,
                debug_dir=debug_dir / f"worker-{idx + 1}",
                limit=None,
            ): chunk
            for idx, chunk in enumerate(chunks)
        }
        for future in as_completed(future_by_chunk):
            chunk = future_by_chunk[future]
            try:
                results.append(future.result())
            except Exception as exc:
                logging.exception("Browser worker failed for %s job(s).", len(chunk))
                failures = [
                    RunIssue(job.station_key, job.search_value, job.excel_row, job.excel_location, job.meter_type, str(exc))
                    for job in chunk
                ]
                results.append(RunResult(measurements={}, successes=[], failures=failures, no_data=[]))

    return merge_run_results(results)


def log_run_report(result: RunResult) -> None:
    logging.info("")
    logging.info("========== EXECUTION REPORT ==========")
    logging.info("SUCCESSFUL: %s", len(result.successes))
    for job in result.successes:
        measurement = result.measurements.get(job.excel_row or -1)
        logging.info(
            "  OK | row=%s | %s | daily=%s max=%s min=%s",
            job.excel_row,
            job.station_key,
            measurement.daily_m3 if measurement else None,
            measurement.max_daily_m3 if measurement else None,
            measurement.min_daily_m3 if measurement else None,
        )

    logging.info("NO DATA / NO ENTRIES: %s", len(result.no_data))
    for issue in result.no_data:
        logging.info("  NO DATA | row=%s | %s | %s", issue.excel_row, issue.station_key, issue.reason)

    logging.info("FAILED: %s", len(result.failures))
    for issue in result.failures:
        logging.info("  FAIL | row=%s | %s | %s", issue.excel_row, issue.station_key, issue.reason)
    logging.info("======================================")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Scrape EcoKing consumption values and update yesterday's sheet in the selected Excel workbook.")
    parser.add_argument("--workbook", default=env("WORKBOOK_PATH", DEFAULT_WORKBOOK), help="Input workbook path.")
    parser.add_argument("--headless", action="store_true", help="Run browser hidden.")
    parser.add_argument("--headed", action="store_true", help="Run browser visible.")
    parser.add_argument("--slow-mo-ms", type=int, default=int(env("SLOW_MO_MS", "0") or "0"), help="Delay browser actions so clicks are visible.")
    parser.add_argument("--workers", type=int, default=int(env("WORKERS", "1") or "1"), help="Number of parallel browser workers. Default: 1.")
    parser.add_argument("--keep-browser-open", action="store_true", help="Keep headed browser open after the run.")
    parser.add_argument("--verbose", action="store_true", default=as_bool(env("VERBOSE"), default=True), help="Enable verbose logs.")
    parser.add_argument("--limit", type=int, default=None, help="Only process the first N non-empty locations.")
    return parser.parse_args()


def main() -> int:
    load_dotenv()
    args = parse_args()
    configure_logging(args.verbose)

    workbook_path = Path(args.workbook)
    if not workbook_path.exists():
        raise FileNotFoundError(workbook_path)
    target_date = yesterday_date()

    headless_default = as_bool(env("HEADLESS"), default=False)
    headless = True if args.headless else False if args.headed else headless_default
    keep_browser_open = args.keep_browser_open or as_bool(env("KEEP_BROWSER_OPEN"), default=False)
    location_map = load_location_map(Path(env("LOCATION_MAP_PATH", DEFAULT_LOCATION_MAP) or DEFAULT_LOCATION_MAP))

    rows = load_location_rows(workbook_path)
    logging.info("Loaded %s workbook rows from %s", len(rows), workbook_path)
    jobs = build_station_jobs(location_map, rows)
    logging.info("Built %s station-driven scrape jobs from %s mappings", len(jobs), len(location_map))
    result = run_browser_parallel(
        jobs=jobs,
        workers=max(1, args.workers),
        headless=headless,
        slow_mo_ms=args.slow_mo_ms,
        keep_browser_open=keep_browser_open,
        debug_dir=Path("debug"),
        limit=args.limit,
    )
    measurements = result.measurements

    duplicate_daily_sheet(workbook_path, target_date, measurements)
    log_run_report(result)
    logging.info("Created sheet %s in %s with %s scraped rows", target_date.strftime(DATE_FORMAT), workbook_path, len(measurements))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        raise SystemExit(130)
    except Exception as exc:
        logging.exception("Run failed: %s", exc)
        raise SystemExit(1)
