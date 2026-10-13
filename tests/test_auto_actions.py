"""Offline checks for auto-add / min-price timer management (no requests reach Ozon)."""
import sys
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from scripts import auto_actions as aa  # noqa: E402

NOW = datetime(2026, 10, 12, 6, 0, tzinfo=timezone.utc)


def price_item(pid, offer, auto_add=True, auto_action=False, min_price="0"):
    return {"product_id": pid, "offer_id": offer, "price": {
        "price": "1000", "old_price": "1500", "min_price": min_price,
        "auto_add_to_ozon_actions_list_enabled": auto_add, "auto_action_enabled": auto_action}}


class BuildRowsTests(unittest.TestCase):
    def test_timer_hours_and_flags(self):
        rows = aa.build_rows(
            [price_item(1, "7005"), price_item(2, "1101", auto_add=False, min_price="900")],
            {1: {"product_id": 1, "min_price_for_auto_actions_enabled": True, "expired_at": "2026-10-14T03:43:54.006026Z"},
             2: {"product_id": 2, "min_price_for_auto_actions_enabled": False}},
            [], now=NOW)
        by = {r["offer_id"]: r for r in rows}
        self.assertEqual(by["7005"]["hours_left"], 45.7)
        self.assertTrue(by["7005"]["min_price_on"])
        self.assertIsNone(by["7005"]["min_price"])
        self.assertIsNone(by["1101"]["hours_left"])
        self.assertFalse(by["1101"]["auto_add"])
        self.assertEqual(by["1101"]["min_price"], 900.0)

    def test_detects_auto_add_switched_back_by_ozon(self):
        log = [{"at": "2026-10-01T10:00:00", "action": "auto_add_off", "product_id": 1, "confirmed": True}]
        rows = aa.build_rows([price_item(1, "1101", auto_add=True)], {}, log, now=NOW)
        self.assertTrue(rows[0]["reverted_by_ozon"])
        rows = aa.build_rows([price_item(1, "1101", auto_add=False)], {}, log, now=NOW)
        self.assertFalse(rows[0]["reverted_by_ozon"])
        self.assertEqual(rows[0]["auto_add_changed_at"], "2026-10-01T10:00:00")


class ApplyTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.addCleanup(self.temp.cleanup)
        patcher = patch.object(aa.time, "sleep", lambda _s: None)
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_disable_auto_add_retries_with_current_prices_and_verifies(self):
        sent = []

        def fake_post(path, payload, attempts=3):
            sent.append((path, payload))
            if len(sent) == 1:
                return {"result": [{"product_id": 1, "updated": False, "errors": [{"message": "price is required"}]}]}
            return {"result": [{"product_id": 1, "updated": True, "errors": []}]}

        row = {"product_id": 1, "offer_id": "1101", "price": 1000.0, "old_price": 1500.0, "min_price": None, "expires_at": None}
        fresh = {"rows": [{**row, "auto_add": False}]}
        with patch.object(aa, "_post", side_effect=fake_post), patch.object(aa, "load_state", return_value=fresh):
            result = aa.apply_change(self.root, "auto_add_off", [row])
        self.assertEqual(result["done"], 1)
        self.assertEqual(sent[0][1]["prices"][0], {"offer_id": "1101", "product_id": 1, "auto_add_to_ozon_actions_list_enabled": "DISABLED"})
        self.assertEqual(sent[1][1]["prices"][0]["price"], "1000")          # повтор с текущими ценами — без изменений
        self.assertEqual(sent[1][1]["prices"][0]["old_price"], "1500")
        log = aa.read_log(self.root)
        self.assertTrue(log[0]["confirmed"])
        self.assertEqual(log[0]["action"], "auto_add_off")

    def test_timer_update_is_confirmed_by_new_expiry(self):
        row = {"product_id": 7, "offer_id": "7005", "expires_at": "2026-10-14T03:43:54+00:00"}
        fresh = {"rows": [{**row, "expires_at": "2026-11-13T06:00:00+00:00"}]}
        with patch.object(aa, "_post", return_value={}) as post, patch.object(aa, "load_state", return_value=fresh):
            result = aa.apply_change(self.root, "timer_update", [row])
        post.assert_called_once_with("/v1/product/action/timer/update", {"product_ids": ["7"]})
        self.assertEqual((result["done"], result["ok"]), (1, True))

    def test_min_price_needs_min_price(self):
        row = {"product_id": 3, "offer_id": "7001", "min_price": None, "expires_at": None}
        with patch.object(aa, "_post") as post, patch.object(aa, "load_state", return_value={"rows": [{**row, "min_price_on": False}]}):
            result = aa.apply_change(self.root, "min_price_on", [row])
        post.assert_not_called()
        self.assertIn("не задана минимальная цена", result["failed"][0]["reason"])


if __name__ == "__main__":
    unittest.main()
