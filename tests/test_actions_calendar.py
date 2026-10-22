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


def build(builder, products, actions=None, today="2026-10-22"):
    """Собирает план по подставным данным Ozon: products — {(action_id, endpoint): [...]}."""
    from pathlib import Path
    from unittest.mock import patch
    context = {"limits": LIMITS, "pid_to_offer": PIDS, "actions": actions or [ACTION]}
    with patch.object(pm, "_load_plan_context", return_value=(context, None)), \
            patch.object(pm, "fetch_action_products", side_effect=lambda action_id, endpoint: products.get((action_id, endpoint), [])), \
            patch.object(pm, "_attach_visibility", return_value=False), \
            patch.object(pm, "_moscow_today", return_value=today):
        plan = builder(Path("."))
    return {item["offer_id"]: item for item in plan["items"]}


class RemovalPlanTests(unittest.TestCase):
    def test_participant_below_min_is_raised_when_action_allows(self):
        items = build(pm.build_action_removal_plan, {(1, "products"): [
            {"id": 11, "price": 900, "action_price": 900, "max_action_price": 1100, "add_mode": "AUTO"}]})
        self.assertEqual((items["A"]["kind"], items["A"]["action_price"], items["A"]["selected"]), ("raise", 1100, True))
        self.assertIn("поднимаем цену вместо вывода", items["A"]["note"])
        self.assertIn("Ozon добавил товар сам", items["A"]["note"])

    def test_raise_respects_limits_of_other_actions_the_product_stays_in(self):
        other = {"id": 2, "name": "Эластичный", "action_type": "STOCK_DISCOUNT", "fixed_boost": 30.0}
        items = build(pm.build_action_removal_plan, {
            (1, "products"): [{"id": 11, "price": 900, "action_price": 900, "max_action_price": 1150}],
            (2, "products"): [{"id": 11, "price": 900, "action_price": 1020, "max_action_price": 1050}],
        }, actions=[ACTION, other])
        self.assertEqual((items["A"]["kind"], items["A"]["action_price"]), ("raise", 1050))

    def test_participant_below_min_is_removed_when_max_too_low(self):
        items = build(pm.build_action_removal_plan, {(1, "products"): [
            {"id": 12, "price": 900, "action_price": 900, "max_action_price": 900, "add_mode": "MANUAL"}]})
        self.assertEqual((items["B"]["kind"], items["B"]["exit_price"]), ("remove", 1000))

    def test_removal_from_not_started_action_keeps_product_price_and_shows_start(self):
        action = {**ACTION, "date_start": "2026-10-27T21:00:00Z"}
        items = build(pm.build_action_removal_plan, {(1, "products"): [
            {"id": 12, "price": 1050, "action_price": 900, "max_action_price": 900, "add_mode": "AUTO"}]}, actions=[action])
        self.assertEqual((items["B"]["kind"], items["B"]["exit_price"], items["B"]["action_start"]), ("remove", 1050, "2026-10-28"))
        self.assertIn("Акция начнётся 28.10", items["B"]["note"])

    def test_profitable_participant_is_left_alone(self):
        from pathlib import Path
        from unittest.mock import patch
        context = {"limits": LIMITS, "pid_to_offer": PIDS, "actions": [ACTION]}
        with patch.object(pm, "_load_plan_context", return_value=(context, None)), \
                patch.object(pm, "fetch_action_products", return_value=[{"id": 11, "action_price": 1050, "max_action_price": 1050}]):
            plan = pm.build_action_removal_plan(Path("."))
        self.assertEqual(plan["items"], [])


class AddPlanWatchTests(unittest.TestCase):
    def test_candidate_ozon_will_add_below_min_is_a_warning_only(self):
        action = {**ACTION, "auto_add_dates": ["2026-10-01T21:00:00Z", "2026-10-25T21:00:00Z"]}
        items = build(pm.build_action_add_plan, {(1, "candidates"): [
            {"id": 13, "action_price": 0, "max_action_price": 1500.7},
            {"id": 14, "action_price": 0, "max_action_price": 950, "min_seller_price": 850},
            # Минимальная в Ozon выше лимита — Ozon сам не добавит, предупреждать не о чем.
            {"id": 12, "action_price": 0, "max_action_price": 950, "min_seller_price": 980},
        ]}, actions=[action])
        self.assertEqual((items["C"]["kind"], items["C"]["action_price"]), ("add", 1200))
        self.assertEqual(set(items), {"C", "D"})
        self.assertEqual((items["D"]["kind"], items["D"]["selected"], items["D"]["ozon_min_price"]), ("watch", False, 850))
        self.assertIn("26.10", items["D"]["note"])
        self.assertIn("на 50 ₽ ниже минимальной", items["D"]["note"])

    def test_no_warning_without_upcoming_auto_add_dates(self):
        items = build(pm.build_action_add_plan, {(1, "candidates"): [
            {"id": 14, "action_price": 0, "max_action_price": 950, "min_seller_price": 850}]})
        self.assertEqual(items, {})

    def test_watch_items_are_never_sent(self):
        from unittest.mock import patch
        plan = {"items": [{"kind": "watch", "action_id": 1, "action_name": "A", "product_id": 14, "offer_id": "D", "selected": True}]}
        with patch.object(pm, "activate_products_in_action") as activate:
            result = pm.execute_action_plan(plan, "add")
        activate.assert_not_called()
        self.assertEqual(result["expected"], 0)


class EntryPriceTests(unittest.TestCase):
    def test_price_is_capped_by_actions_the_product_stays_in(self):
        # Товар C участвует в другой акции (лимит 1080) и остаётся там: цена одна, 1080, и в этой акции.
        entry = pm._plan_entry_prices([(13, 1, 1000.0, 1200.0, 1500.7, False), (13, 2, 1000.0, 1200.0, 1080, True),
                                       (14, 1, 1000.0, 1200.0, 950, False)])
        self.assertEqual(entry, {"price": {13: 1080}, "costly": {}})

    def test_candidate_that_lowers_the_price_noticeably_is_left_unchecked(self):
        # Ради акции с лимитом 1050 товару пришлось бы снизить цену с 1200 во всех акциях — решает пользователь.
        entry = pm._plan_entry_prices([(13, 2, 1000.0, 1200.0, 1300, False), (13, 1, 1000.0, 1200.0, 1050, False)])
        self.assertEqual(entry, {"price": {13: 1200}, "costly": {(13, 1): 1050}})
        self.assertEqual(pm._costly_note(1200, 1050), "Не отмечено: участие снизит цену товара во всех акциях с 1 200 до 1 050 ₽ "
                                                      "(−150 ₽). Отметьте, если бустинг того стоит.")

    def test_small_price_drop_keeps_both_actions(self):
        # 1200 → 1195: меньше max(10 ₽, 1%) — обе акции в общей цене.
        entry = pm._plan_entry_prices([(13, 2, 1000.0, 1200.0, 1300, False), (13, 1, 1000.0, 1200.0, 1195, False)])
        self.assertEqual(entry, {"price": {13: 1195}, "costly": {}})



if __name__ == "__main__":
    unittest.main()
