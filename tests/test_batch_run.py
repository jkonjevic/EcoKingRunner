import os
import shutil
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

from ecoking import logtext, webapp


class SelectedDatesTests(unittest.TestCase):
    """A batch is validated as a whole, before any of it starts."""

    def test_a_single_date_is_a_batch_of_one(self):
        self.assertEqual(webapp.selected_dates({"selectedDate": "2026-07-31"}), ["2026-07-31"])

    def test_missing_dates_fall_back_to_yesterday(self):
        self.assertEqual(webapp.selected_dates({}), [webapp.yesterday()])

    def test_dates_are_deduplicated_and_sorted(self):
        payload = {"selectedDates": ["2026-07-31", "2026-07-29", "2026-07-31"]}
        self.assertEqual(webapp.selected_dates(payload), ["2026-07-29", "2026-07-31"])

    def test_selected_dates_wins_over_the_single_date_field(self):
        payload = {"selectedDate": "2026-07-01", "selectedDates": ["2026-07-29"]}
        self.assertEqual(webapp.selected_dates(payload), ["2026-07-29"])

    def test_every_bad_date_is_reported_at_once(self):
        payload = {"selectedDates": ["2099-01-01", "not-a-date", "2026-02-30"]}
        with self.assertRaises(ValueError) as caught:
            webapp.selected_dates(payload)
        message = str(caught.exception)
        self.assertIn("2099-01-01", message)
        self.assertIn("not-a-date", message)
        self.assertIn("2026-02-30", message)

    def test_an_empty_selection_is_rejected(self):
        with self.assertRaises(ValueError):
            webapp.selected_dates({"selectedDates": ["   "]})

    def test_too_many_days_are_rejected(self):
        dates = [f"2026-06-{day:02d}" for day in range(1, webapp.MAX_BATCH_DAYS + 2)]
        with self.assertRaises(ValueError):
            webapp.selected_dates({"selectedDates": dates})


class BatchStateTestCase(unittest.TestCase):
    """Runs the queue with a stubbed subprocess, so no browser is launched."""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp, True)
        patch = mock.patch.dict(os.environ, {"DATA_DIR": str(self.tmp), "ECOKING_MODE": "cloud"})
        patch.start()
        self.addCleanup(patch.stop)
        webapp.STATE = webapp.RunState()
        self.addCleanup(setattr, webapp, "STATE", webapp.RunState())

    def wait_for_idle(self, timeout: float = 5.0) -> None:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            with webapp.STATE_LOCK:
                if not webapp.STATE.running:
                    return
            time.sleep(0.01)
        self.fail("the batch did not finish in time")

    def run_batch(self, dates, process):
        with mock.patch.object(webapp, "_pump_process", process):
            webapp.start_batch(dates, {})
            self.wait_for_idle()

    def statuses(self):
        return [day.status for day in webapp.STATE.days]


class BatchQueueTests(BatchStateTestCase):
    def test_a_failing_day_does_not_take_the_rest_of_the_batch_with_it(self):
        codes = iter([1, 0, 0])
        self.run_batch(
            ["2026-07-29", "2026-07-30", "2026-07-31"], lambda cmd, env: next(codes)
        )
        self.assertEqual(self.statuses(), ["failed", "ok", "ok"])
        self.assertEqual(webapp.STATE.return_code, 1)

    def test_a_clean_batch_reports_success(self):
        self.run_batch(["2026-07-30", "2026-07-31"], lambda cmd, env: 0)
        self.assertEqual(self.statuses(), ["ok", "ok"])
        self.assertEqual(webapp.STATE.return_code, 0)

    def test_stopping_the_batch_skips_the_days_that_never_ran(self):
        def cancel(cmd, env):
            with webapp.STATE_LOCK:
                webapp.STATE.cancelled = True
            return 1

        self.run_batch(["2026-07-29", "2026-07-30", "2026-07-31"], cancel)
        self.assertEqual(self.statuses(), ["stopped", "skipped", "skipped"])

    def test_skipping_one_day_leaves_the_queue_running(self):
        seen = []

        def skip_first(cmd, env):
            seen.append(cmd)
            if len(seen) == 1:
                with webapp.STATE_LOCK:
                    webapp.STATE.stop_current = True
                return 1
            return 0

        self.run_batch(["2026-07-29", "2026-07-30", "2026-07-31"], skip_first)
        self.assertEqual(self.statuses(), ["stopped", "ok", "ok"])
        self.assertEqual(len(seen), 3)

    def test_each_day_is_scraped_with_its_own_date(self):
        seen = []
        self.run_batch(
            ["2026-07-30", "2026-07-31"], lambda cmd, env: (seen.append(cmd), 0)[1]
        )
        for date, cmd in zip(["2026-07-30", "2026-07-31"], seen):
            self.assertIn(date, cmd, "the scraper is still told about exactly one day")

    def test_a_batch_of_one_builds_the_command_it_always_did(self):
        seen = []
        self.run_batch(["2026-07-31"], lambda cmd, env: (seen.append(cmd), 0)[1])
        expected, _, _ = webapp.build_run_command({"selectedDate": "2026-07-31"})
        self.assertEqual(seen, [expected])

    def test_a_second_batch_is_refused_while_one_is_running(self):
        with mock.patch.object(webapp, "_pump_process", lambda cmd, env: time.sleep(0.2) or 0):
            webapp.start_batch(["2026-07-31"], {})
            with self.assertRaises(RuntimeError):
                webapp.start_batch(["2026-07-30"], {})
            self.wait_for_idle()

    def test_only_a_multi_day_run_writes_day_separators(self):
        self.run_batch(["2026-07-31"], lambda cmd, env: 0)
        self.assertFalse([line for line in webapp.STATE.lines if "DAN 1/1" in line])

        webapp.STATE = webapp.RunState()
        self.run_batch(["2026-07-30", "2026-07-31"], lambda cmd, env: 0)
        separators = [line for line in webapp.STATE.lines if "DAN " in line]
        self.assertEqual(len(separators), 2)
        self.assertIn("2026-07-30", separators[0])

    def test_a_multi_day_run_ends_with_a_summary_line(self):
        codes = iter([0, 1])
        self.run_batch(["2026-07-30", "2026-07-31"], lambda cmd, env: next(codes))
        self.assertTrue(any(line.startswith("ZBIRNO:") for line in webapp.STATE.lines))


class FailureTrackingTests(BatchStateTestCase):
    """Every failure is attributed to the day whose run produced it."""

    def test_failures_and_counts_land_on_the_right_day(self):
        def emit(cmd, env):
            date = webapp.STATE.active.date
            if date == "2026-07-30":
                webapp.append_log(
                    "13:22:01 ERROR        FAIL | row=12 | REZERVOAR PODI / ULAZ | Timeout 30000ms"
                )
                return 1
            webapp.append_log(
                "13:24:11 WARNING    TELEMETRY FAIL | Rezervoar Kula Tele | not in the location table"
            )
            return 0

        self.run_batch(["2026-07-30", "2026-07-31"], emit)
        first, second = webapp.STATE.days
        self.assertEqual([failure.label for failure in first.failures], ["REZERVOAR PODI / ULAZ"])
        self.assertEqual(first.errors, 1)
        self.assertEqual([failure.label for failure in second.failures], ["Rezervoar Kula Tele"])
        self.assertEqual(second.warnings, 1)
        self.assertEqual(second.errors, 0)

    def test_progress_lines_advance_only_the_running_day(self):
        def emit(cmd, env):
            webapp.append_log("13:22:01 INFO       [7/28] Station=Bajer 1, Excel row=4, device=Bajer 1 - U")
            return 0

        self.run_batch(["2026-07-30", "2026-07-31"], emit)
        for day in webapp.STATE.days:
            self.assertEqual((day.done, day.total), (7, 28))
            self.assertEqual(day.current, "Bajer 1")

    def test_the_payload_keeps_the_single_day_keys_it_always_had(self):
        self.run_batch(["2026-07-30", "2026-07-31"], lambda cmd, env: 0)
        handler = webapp.Handler.__new__(webapp.Handler)
        payload = handler._run_state({})
        for key in ("running", "returnCode", "cursor", "lines", "selectedDate", "done", "total", "current"):
            self.assertIn(key, payload)
        self.assertEqual(payload["selectedDate"], "2026-07-31", "the last day drives the tiles")
        self.assertEqual(payload["batch"]["total"], 2)
        self.assertEqual([day["date"] for day in payload["days"]], ["2026-07-30", "2026-07-31"])

    def test_logging_outside_a_run_never_touches_a_day(self):
        self.run_batch(["2026-07-31"], lambda cmd, env: 0)
        before = len(webapp.STATE.days[0].failures)
        webapp.append_log("13:30:00 ERROR        FAIL | row=1 | LATE | after the run")
        self.assertEqual(len(webapp.STATE.days[0].failures), before)


class LogClassificationTests(unittest.TestCase):
    def test_levels_become_severities(self):
        self.assertEqual(logtext.classify("13:22:01 ERROR      Run failed: boom"), "error")
        self.assertEqual(logtext.classify("13:22:01 CRITICAL   Run failed: boom"), "error")
        self.assertEqual(logtext.classify("13:22:01 WARNING    hmm"), "warning")
        self.assertEqual(logtext.classify("13:22:01 INFO       fine"), "info")
        self.assertEqual(logtext.classify("raw stdout with no header"), "info")

    def test_a_failed_station_is_parsed_into_its_parts(self):
        failure = logtext.parse_failure(
            "13:22:01 ERROR        FAIL | row=12 | REZERVOAR PODI / ULAZ | Timeout 30000ms exceeded"
        )
        self.assertEqual(failure.kind, "station")
        self.assertEqual(failure.row, "12")
        self.assertEqual(failure.label, "REZERVOAR PODI / ULAZ")
        self.assertEqual(failure.severity, "error")
        # The panel shows the reason, so it is translated on the way out.
        self.assertEqual(failure.reason, "Isteklo je vrijeme čekanja (30000 ms).")

    def test_no_data_is_a_warning_not_an_error(self):
        failure = logtext.parse_failure("13:22:01 WARNING      NO DATA | row=7 | KULA / ULAZ | no point")
        self.assertEqual(failure.kind, "no-data")
        self.assertEqual(failure.severity, "warning")

    def test_a_telemetry_failure_has_no_excel_row(self):
        failure = logtext.parse_failure(
            "13:22:01 WARNING    TELEMETRY FAIL | Rezervoar Kula Tele | missing 17:00 row"
        )
        self.assertEqual(failure.kind, "telemetry")
        self.assertEqual(failure.row, "")
        self.assertEqual(failure.label, "Rezervoar Kula Tele")

    def test_ordinary_lines_are_not_failures(self):
        self.assertIsNone(logtext.parse_failure("13:22:01 INFO       Opening https://ecoking"))
        self.assertIsNone(logtext.parse_failure("13:22:01 INFO       OK | row=4 | Bajer 1 | daily=12"))

    def test_a_failure_line_still_translates_for_the_console(self):
        translated = logtext.translate("13:22:01 ERROR        FAIL | row=12 | KULA | Timeout")
        self.assertIn("NEUSPJEH", translated)
        self.assertIn("GREŠKA", translated)


class ReasonTranslationTests(unittest.TestCase):
    """Failure reasons are English in the scraper and Montenegrin in the UI."""

    #: The real line behind "I can't make PODI work".
    PODI = (
        "Could not select device '358004092223510 - Herceg Novi - (R-PO) Podi - I'. "
        "Attempts: '(R-PO) Podi - I': No dropdown result matches '(R-PO) Podi - I'. "
        "Visible: ['358004092223510 - Herceg Novi - (R-PO) Podi - I (358004092223510)']"
    )

    def test_a_nested_reason_is_translated_all_the_way_down(self):
        translated = logtext.translate_reason(self.PODI)
        self.assertIn("Nije moguće izabrati uređaj", translated)
        self.assertIn("Pokušaji:", translated)
        self.assertIn("Nijedan rezultat ne odgovara nazivu", translated)
        self.assertIn("Vidljivo je:", translated)
        for english in ("Could not", "No dropdown", "Attempts", "Visible"):
            self.assertNotIn(english, translated)

    def test_the_device_name_inside_a_reason_is_left_alone(self):
        self.assertIn("(R-PO) Podi - I", logtext.translate_reason(self.PODI))

    def test_playwright_wording_is_translated_too(self):
        self.assertEqual(
            logtext.translate_reason("Timeout 30000ms exceeded."),
            "Isteklo je vrijeme čekanja (30000 ms).",
        )

    def test_an_ambiguous_name_reads_as_an_instruction(self):
        reason = logtext.translate_reason(
            "Device name 'Podi - I' matches 2 devices: ['a', 'b']. "
            "Make the name in the station list more specific."
        )
        self.assertIn("odgovara za 2 uređaja", reason)
        self.assertIn("Precizirajte naziv u listi stanica.", reason)

    def test_telemetry_reasons_are_already_montenegrin_and_survive(self):
        reason = "Nema reda za 17:00 na dan 2026-07-31 u tabeli."
        self.assertEqual(logtext.translate_reason(reason), reason)

    def test_an_unknown_reason_is_passed_through_untouched(self):
        self.assertEqual(logtext.translate_reason("something new"), "something new")


class ElementLabelTests(unittest.TestCase):
    def test_selector_debug_lines_name_the_field_in_montenegrin(self):
        self.assertEqual(
            logtext.translate("13:22:01 DEBUG      Filling email with selector input[type='email']"),
            "13:22:01 DETALJ     Unosim e-mail preko selektora input[type='email']",
        )

    def test_a_bracketed_note_on_a_label_is_kept(self):
        translated = logtext.translate(
            "13:22:01 DEBUG      Clicking website date picker (Choose Date...) with selector div[x]"
        )
        self.assertIn("Klik na izbor datuma na sajtu (Choose Date...)", translated)

    def test_an_unknown_element_name_is_left_as_it_was(self):
        self.assertIn(
            "Klik na neko novo dugme",
            logtext.translate("13:22:01 DEBUG      Clicking neko novo dugme with selector button"),
        )


if __name__ == "__main__":
    unittest.main()
