"""Offline checks for the promotions calendar dates."""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import price_management as pm  # noqa: E402


class MoscowDateTests(unittest.TestCase):
    def test_utc_evening_is_next_day_in_moscow(self):
        self.assertEqual(pm._moscow_date("2026-10-06T21:00:00Z"), "2026-10-07")

    def test_utc_daytime_keeps_date(self):
        self.assertEqual(pm._moscow_date("2026-10-21T20:59:59Z"), "2026-10-21")

    def test_empty_and_plain_dates(self):
        self.assertIsNone(pm._moscow_date(None))
        self.assertIsNone(pm._moscow_date(""))
        self.assertEqual(pm._moscow_date("2026-10-07"), "2026-10-07")


ACTION = {"id": 1, "name": "Максимальный бустинг", "action_type": "STOCK_DISCOUNT", "fixed_boost": 55.0}
LIMITS = {"A": (1000.0, 1200.0), "B": (1000.0, 1200.0), "C": (1000.0, 1200.0), "D": (1000.0, 1200.0)}
PIDS = {11: "A", 12: "B", 13: "C", 14: "D"}


class UpcomingScopeTests(unittest.TestCase):
    def test_future_start_or_auto_add_is_upcoming(self):
        self.assertIsNotNone(pm.upcoming_action_scope({"date_start": "2026-10-11T21:00:00Z"}, "2026-10-08"))
        scope = pm.upcoming_action_scope({"date_start": "2026-03-19T21:00:00Z",
                                          "auto_add_dates": ["2026-10-01T21:00:00Z", "2026-10-15T21:00:00Z"]}, "2026-10-08")
        self.assertEqual(scope["auto_add"], ["2026-10-16"])

    def test_running_action_without_auto_add_is_skipped(self):
        self.assertIsNone(pm.upcoming_action_scope({"date_start": "2026-09-08T21:00:00Z", "auto_add_dates": []}, "2026-10-08"))


class UpcomingPlanTests(unittest.TestCase):
    def plan(self, participants, candidates):
        scope = {"start": "2026-10-12", "auto_add": ["2026-10-14"]}
        return {i["offer_id"]: i for i in pm.plan_upcoming_action_items(ACTION, scope, participants, candidates, LIMITS, PIDS)}

    def test_participant_below_min_is_raised_when_action_allows(self):
        items = self.plan([{"id": 11, "action_price": 900, "max_action_price": 1100, "add_mode": "AUTO"}], [])
        self.assertEqual((items["A"]["kind"], items["A"]["action_price"]), ("raise", 1100))
        self.assertTrue(items["A"]["added_by_ozon"])

    def test_participant_below_min_is_removed_when_max_too_low(self):
        items = self.plan([{"id": 12, "action_price": 900, "max_action_price": 900, "add_mode": "MANUAL"}], [])
        self.assertEqual(items["B"]["kind"], "remove")

    def test_profitable_participant_is_left_alone(self):
        self.assertEqual(self.plan([{"id": 11, "action_price": 1050, "max_action_price": 1050}], []), {})

    def test_candidates_added_early_or_watched(self):
        items = self.plan([], [{"id": 13, "action_price": 0, "max_action_price": 1500.7},
                               {"id": 14, "action_price": 0, "max_action_price": 950}])
        self.assertEqual((items["C"]["kind"], items["C"]["action_price"]), ("add", 1200))
        self.assertIn("14.10", items["C"]["note"])
        self.assertEqual(items["D"]["kind"], "watch")
        self.assertFalse(items["D"]["selected"])
        self.assertIn("на 50 ₽ ниже минимальной", items["D"]["note"])


if __name__ == "__main__":
    unittest.main()
