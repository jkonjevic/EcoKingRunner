from datetime import datetime
import unittest

from ecoking_daily import chart_point_date, latest_value_from_payload, target_dates, value_for_date_from_payload


class ChartSelectionTests(unittest.TestCase):
    def test_30_day_target_date_is_not_offset(self) -> None:
        selected_date = datetime(2026, 7, 23)
        interval_date, total_date = target_dates(selected_date)
        self.assertEqual(interval_date, selected_date)
        self.assertEqual(total_date, selected_date)

    def test_30_day_value_uses_last_bar_after_selecting_ui_date(self) -> None:
        payload = {
            "highcharts": [{"series": [{"points": [
                {"x": datetime(2026, 7, 22).timestamp() * 1000, "y": 12.0},
                {"x": datetime(2026, 7, 23).timestamp() * 1000, "y": 34.5},
            ]}]}],
            "chartjs": [],
            "tables": [],
        }
        self.assertEqual(latest_value_from_payload(payload), 34.5)

    def test_highcharts_point_metadata_selects_ui_date_not_last_point(self) -> None:
        selected_date = datetime(2026, 7, 23)
        payload = {
            "highcharts": [
                {
                    "series": [
                        {
                            "name": "Usage (Cubic Metres)",
                            "type": "column",
                            "points": [
                                {"x": datetime(2026, 7, 22).timestamp() * 1000, "y": 12.0, "category": None, "name": None, "date": "2026-07-22"},
                                {"x": datetime(2026, 7, 23).timestamp() * 1000, "y": 34.5, "category": None, "name": None, "date": "2026-07-23"},
                                {"x": datetime(2026, 7, 24).timestamp() * 1000, "y": 98.0, "category": None, "name": None, "date": "2026-07-24"},
                            ],
                        }
                    ]
                }
            ],
            "chartjs": [],
            "tables": [],
        }

        self.assertEqual(value_for_date_from_payload(payload, selected_date), 34.5)

    def test_highcharts_point_metadata_includes_x_y_category_and_name(self) -> None:
        point = {
            "x": datetime(2026, 7, 23).timestamp() * 1000,
            "y": 34.5,
            "category": "23/07/2026",
            "name": "23/07/2026",
        }
        self.assertEqual(chart_point_date(point["x"]), datetime(2026, 7, 23).date())
        self.assertEqual(point["y"], 34.5)
        self.assertEqual(chart_point_date(point["category"]), datetime(2026, 7, 23).date())
        self.assertEqual(chart_point_date(point["name"]), datetime(2026, 7, 23).date())


if __name__ == "__main__":
    unittest.main()
