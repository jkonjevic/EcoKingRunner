import json
import os
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from ecoking import stations as registry
from ecoking import webapp
from ecoking.stations import ExcelRow, Station

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "ECO KING BLANKO TABLICA.xlsx"


class DeviceLabelTests(unittest.TestCase):
    def test_serial_and_city_are_stripped(self) -> None:
        self.assertEqual(
            registry.device_label("358004092234384 - Herceg Novi - (R-BA) Bajer 1 - U"),
            "(R-BA) Bajer 1 - U",
        )

    def test_placeholder_instead_of_a_serial_is_stripped(self) -> None:
        self.assertEqual(
            registry.device_label(
                "NEMA SERIJSKOG BROJA U TABELI - Herceg Novi - Rezervoar Kumbor lijeva komora - U"
            ),
            "Rezervoar Kumbor lijeva komora - U",
        )

    def test_label_without_a_city_prefix_is_left_alone(self) -> None:
        self.assertEqual(
            registry.device_label("358004092229558 - Direktna prema Baošićima DN300"),
            "Direktna prema Baošićima DN300",
        )

    def test_whitespace_is_collapsed(self) -> None:
        self.assertEqual(registry.device_label("  358004092236694 -  Herceg Novi -  Topla  - I "), "Topla - I")

    def test_the_repeated_serial_in_brackets_is_stripped(self) -> None:
        self.assertEqual(
            registry.device_label(
                "358004092223510 - Herceg Novi - (R-PO) Podi - I (358004092223510)"
            ),
            "(R-PO) Podi - I",
        )

    def test_reducing_an_already_short_label_changes_nothing(self) -> None:
        # Idempotence is what lets score_device_match reduce both sides.
        for label in ("(R-PO) Podi - I", "Bajer 1 - U", "Direktna prema Baošićima DN300"):
            self.assertEqual(registry.device_label(label), label)

    def test_a_bracketed_code_is_not_mistaken_for_a_serial(self) -> None:
        self.assertEqual(registry.device_label("Podi - I (rezerva)"), "Podi - I (rezerva)")


class DeviceMatchTests(unittest.TestCase):
    def test_short_label_matches_the_full_site_entry(self) -> None:
        self.assertGreater(registry.score_device_match("Bajer 1 - U", "(R-BA) Bajer 1 - U"), 0)

    def test_direction_suffix_is_not_ignored(self) -> None:
        self.assertEqual(registry.score_device_match("Bajer 1 - U", "(R-BA) Bajer 1 - I"), 0)

    def test_exact_label_outranks_a_suffix_match(self) -> None:
        exact = registry.score_device_match("(R-PO) Podi - I", "(R-PO) Podi - I")
        suffix = registry.score_device_match("Podi - I", "(R-PO) Podi - I")
        self.assertGreater(exact, suffix)

    def test_accents_and_punctuation_do_not_matter(self) -> None:
        self.assertGreater(registry.score_device_match("Spanjola - U", "(RS) Španjola – U"), 0)

    def test_a_bare_name_ties_between_two_prefixed_devices(self) -> None:
        # The tie is what makes the scraper refuse to guess between R-PO and PS-PO.
        left = registry.score_device_match("Podi - I", "(R-PO) Podi - I")
        right = registry.score_device_match("Podi - I", "(PS-PO) Podi - I")
        self.assertEqual(left, right)
        self.assertGreater(left, 0)

    def test_a_dropdown_entry_pasted_whole_matches_the_device_it_came_from(self) -> None:
        # Pasting the site's own entry is the obvious way to disambiguate a
        # name, and it used to score zero because only the option was reduced.
        site_entry = "358004092223510 - Herceg Novi - (R-PO) Podi - I (358004092223510)"
        label = registry.device_label(site_entry)
        self.assertEqual(
            registry.score_device_match("358004092223510 - Herceg Novi - (R-PO) Podi - I", label),
            registry.score_device_match("(R-PO) Podi - I", label),
        )
        self.assertGreater(registry.score_device_match(site_entry, label), 0)

    def test_a_pasted_entry_still_refuses_the_wrong_device(self) -> None:
        pasted = "358004092223510 - Herceg Novi - (R-PO) Podi - I"
        self.assertEqual(registry.score_device_match(pasted, "(PS-PO) Podi - I"), 0)
        self.assertEqual(registry.score_device_match(pasted, "(R-PO) Podi - U"), 0)

    def test_navigation_entries_are_recognised(self) -> None:
        self.assertTrue(registry.is_navigation_option("Device List"))
        self.assertFalse(registry.is_navigation_option("Bajer 1 - U"))


class SearchQueryTests(unittest.TestCase):
    def test_full_label_comes_first_then_shorter_fallbacks(self) -> None:
        self.assertEqual(
            registry.search_queries("(R-PO) Podi - I"),
            ["(R-PO) Podi - I", "Podi - I", "Podi"],
        )

    def test_a_plain_label_produces_no_duplicates(self) -> None:
        self.assertEqual(registry.search_queries("Hotel Park"), ["Hotel Park"])


class ShortestUniqueLabelTests(unittest.TestCase):
    def test_prefix_is_dropped_when_unambiguous(self) -> None:
        labels = ["(R-BA) Bajer 1 - U", "(R-BA) Bajer 1 - I"]
        self.assertEqual(registry.shortest_unique_label(labels[0], labels), "Bajer 1 - U")

    def test_prefix_is_kept_when_it_is_the_only_difference(self) -> None:
        labels = ["(R-PO) Podi - I", "(PS-PO) Podi - I"]
        self.assertEqual(registry.shortest_unique_label(labels[0], labels), "(R-PO) Podi - I")


class TemplateRowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        if not TEMPLATE.exists():
            raise unittest.SkipTest(f"{TEMPLATE.name} is not available.")
        cls.rows = registry.load_excel_rows(TEMPLATE)

    def test_every_row_carries_a_location_and_a_meter(self) -> None:
        self.assertTrue(self.rows)
        for row in self.rows:
            self.assertTrue(row.lokacija, row)
            self.assertTrue(row.vodomjer, row)

    def test_location_is_inherited_across_a_group(self) -> None:
        grouped = [row for row in self.rows if row.row in {24, 25, 26, 27}]
        self.assertEqual({row.lokacija for row in grouped}, {"REZERVOAR BAJER 2"})
        self.assertEqual(
            [row.vodomjer for row in grouped],
            ["ULAZ", "IZLAZ ZA ČELA", "ULAZ REZERVOARA PODI", "IZLAZ ZA ORJENSKI BATALJON"],
        )

    def test_location_and_meter_uniquely_identify_a_row(self) -> None:
        keys = [row.key for row in self.rows]
        self.assertEqual(len(keys), len(set(keys)))


class ValidationTests(unittest.TestCase):
    rows = [
        ExcelRow(row=2, lokacija="REZERVOAR BAJER 1", vodomjer="ULAZ"),
        ExcelRow(row=3, lokacija="REZERVOAR BAJER 1", vodomjer="IZLAZ"),
    ]

    def messages(self, stations, severity="error"):
        return [i.message for i in registry.validate(stations, self.rows) if i.severity == severity]

    def test_a_matching_station_has_no_errors(self) -> None:
        stations = [Station("REZERVOAR BAJER 1", "ULAZ", "Bajer 1 - U")]
        self.assertEqual(self.messages(stations), [])

    def test_a_station_without_a_template_row_is_an_error(self) -> None:
        stations = [Station("NEPOSTOJEĆA", "ULAZ", "X")]
        self.assertIn("Nema reda u template-u sa ovom kombinacijom LOKACIJA + VODOMJER.", self.messages(stations))

    def test_two_stations_writing_to_one_row_is_an_error(self) -> None:
        stations = [
            Station("REZERVOAR BAJER 1", "ULAZ", "Bajer 1 - U"),
            Station("REZERVOAR BAJER 1", "ULAZ", "Nešto drugo"),
        ]
        self.assertIn("Dva unosa pišu u isti red template-a.", self.messages(stations))

    def test_a_missing_device_name_is_an_error(self) -> None:
        stations = [Station("REZERVOAR BAJER 1", "ULAZ", "")]
        self.assertIn("Naziv uređaja na sajtu je obavezan.", self.messages(stations))

    def test_an_uncovered_template_row_is_only_a_warning(self) -> None:
        stations = [Station("REZERVOAR BAJER 1", "ULAZ", "Bajer 1 - U")]
        self.assertEqual(self.messages(stations), [])
        self.assertTrue(any("ostaće prazan" in m for m in self.messages(stations, "warning")))

    def test_ambiguous_device_names_are_reported(self) -> None:
        stations = [
            Station("REZERVOAR BAJER 1", "ULAZ", "Podi - I"),
            Station("REZERVOAR BAJER 1", "IZLAZ", "(R-PO) Podi - I"),
        ]
        self.assertTrue(registry.ambiguous_device_labels(stations))


class FileTests(unittest.TestCase):
    def test_save_then_load_round_trips(self) -> None:
        stations = [
            Station("REZERVOAR BAJER 1", "ULAZ", "Bajer 1 - U"),
            Station("REZERVOAR BAJER 1", "IZLAZ", "Bajer 1 - I", enabled=False),
        ]
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "stations.json"
            registry.save_stations(path, stations)
            self.assertEqual(registry.load_stations(path), stations)
            self.assertEqual(json.loads(path.read_text(encoding="utf-8"))["version"], registry.SCHEMA_VERSION)

    def test_the_legacy_serial_file_is_upgraded_on_read(self) -> None:
        legacy = {
            "REZERVOAR BAJER 1 - U": "358004092234384 - Herceg Novi - (R-BA) Bajer 1 - U",
            "REZERVOAR BAJER 1 - I": "358004092088707 - Herceg Novi - (R-BA) Bajer 1 - I",
        }
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "herceg_novi_stations.json"
            path.write_text(json.dumps(legacy, ensure_ascii=False), encoding="utf-8")
            stations = registry.load_stations(path)

        self.assertEqual(
            stations,
            [
                Station("REZERVOAR BAJER 1", "ULAZ", "(R-BA) Bajer 1 - U"),
                Station("REZERVOAR BAJER 1", "IZLAZ", "(R-BA) Bajer 1 - I"),
            ],
        )

    def test_a_missing_file_reads_as_empty(self) -> None:
        self.assertEqual(registry.load_stations(Path("no-such-file.json")), [])

    def test_explicit_relative_path_is_resolved_against_the_root(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            target = root / "stations.json"
            target.write_text("[]", encoding="utf-8")
            self.assertEqual(registry.resolve_stations_path("stations.json", root), target)

    def test_stations_json_wins_over_the_legacy_file(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / registry.LEGACY_STATIONS_FILE).write_text("{}", encoding="utf-8")
            self.assertEqual(
                registry.resolve_stations_path(None, root).name, registry.LEGACY_STATIONS_FILE
            )
            (root / registry.DEFAULT_STATIONS_FILE).write_text("{}", encoding="utf-8")
            self.assertEqual(
                registry.resolve_stations_path(None, root).name, registry.DEFAULT_STATIONS_FILE
            )


class ReconcileTests(unittest.TestCase):
    rows = [
        ExcelRow(row=24, lokacija="REZERVOAR BAJER 2", vodomjer="ULAZ"),
        ExcelRow(row=26, lokacija="REZERVOAR BAJER 2", vodomjer="ULAZ REZERVOARA PODI"),
    ]

    def test_a_descriptive_legacy_key_snaps_onto_its_row(self) -> None:
        stations = [Station("REZERVOAR BAJER 2 - PUNJENJE REZERVOARA PODI", "IZLAZ", "Bajer 2 - I (punjenje R Podi)")]
        fixed = registry.reconcile_with_template(stations, self.rows)
        self.assertEqual(fixed[0].lokacija, "REZERVOAR BAJER 2")
        self.assertEqual(fixed[0].vodomjer, "ULAZ REZERVOARA PODI")

    def test_an_exact_pairing_is_left_untouched(self) -> None:
        stations = [Station("REZERVOAR BAJER 2", "ULAZ", "Bajer 2 - U")]
        self.assertEqual(registry.reconcile_with_template(stations, self.rows), stations)

    def test_two_stations_never_claim_the_same_row(self) -> None:
        stations = [
            Station("REZERVOAR BAJER 2", "ULAZ", "Bajer 2 - U"),
            Station("REZERVOAR BAJER 2 - PUNJENJE REZERVOARA PODI", "IZLAZ", "Bajer 2 - I (punjenje R Podi)"),
        ]
        fixed = registry.reconcile_with_template(stations, self.rows)
        self.assertEqual(len({station.key for station in fixed}), 2)


class ShippedRegistryTests(unittest.TestCase):
    """The registry that ships with the repo has to stay usable."""

    def test_it_validates_against_the_template(self) -> None:
        if not TEMPLATE.exists():
            self.skipTest(f"{TEMPLATE.name} is not available.")
        path = registry.resolve_stations_path(None, ROOT)
        if not path.exists():
            self.skipTest("No station registry in the repository.")
        stations = registry.load_stations(path)
        rows = registry.load_excel_rows(TEMPLATE)
        errors = [i for i in registry.validate(stations, rows) if i.severity == "error"]
        self.assertEqual(errors, [], f"{len(errors)} blocking issue(s) in {path.name}")

    def test_no_device_name_still_carries_a_serial_number(self) -> None:
        path = registry.resolve_stations_path(None, ROOT)
        if not path.exists():
            self.skipTest("No station registry in the repository.")
        for station in registry.load_stations(path):
            self.assertNotRegex(station.uredjaj, r"\d{8,}", station.label)
            self.assertNotIn("herceg novi", registry.normalize(station.uredjaj), station.label)


class SaveNormalisationTests(unittest.TestCase):
    """Saving from the UI is what keeps the file free of pasted site entries."""

    def setUp(self) -> None:
        if not TEMPLATE.exists():
            self.skipTest(f"{TEMPLATE.name} is not available.")
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp, True)
        patch = mock.patch.dict(
            os.environ,
            {"DATA_DIR": str(self.tmp), "ECOKING_MODE": "cloud", "STATIONS_PATH": str(self.tmp / "stations.json")},
        )
        patch.start()
        self.addCleanup(patch.stop)

    def test_a_pasted_dropdown_entry_is_stored_in_the_short_form(self) -> None:
        row = registry.load_excel_rows(TEMPLATE)[0]
        saved = webapp.save_stations_payload(
            {
                "stations": [
                    {
                        "lokacija": row.lokacija,
                        "vodomjer": row.vodomjer,
                        "uredjaj": "358004092223510 - Herceg Novi - (R-PO) Podi - I (358004092223510)",
                        "enabled": True,
                    }
                ]
            }
        )
        self.assertEqual(saved["stations"][0]["uredjaj"], "(R-PO) Podi - I")
        on_disk = json.loads((self.tmp / "stations.json").read_text(encoding="utf-8"))
        self.assertNotRegex(json.dumps(on_disk), r"\d{8,}")


if __name__ == "__main__":
    unittest.main()
