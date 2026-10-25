"""Offline checks for the promotion methods that replace the blocked /v1/actions/* calls."""
import contextlib
import io
import sys
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import recommended_prices as rp  # noqa: E402


def response(status=200, body=None):
    result = Mock(status_code=status)
    result.json.return_value = body if body is not None else {}
    result.raise_for_status.side_effect = None if status < 400 else rp.requests.HTTPError(f"HTTP {status}", response=result)
    return result


class PromosV2Tests(unittest.TestCase):
    def setUp(self):
        keys = patch.multiple(rp, OZON_CLIENT_ID="id", OZON_API_KEY="key")
        keys.start()
        self.addCleanup(keys.stop)
        output = contextlib.redirect_stdout(io.StringIO())
        output.__enter__()
        self.addCleanup(output.__exit__, None, None, None)

    def test_products_are_read_from_v2_with_numeric_prices(self):
        pages = [
            response(body={"products": [{"id": 1, "action_price": {"amount": "960", "currency": "RUB"},
                                         "max_action_price": {"amount": "1000.5", "currency": "RUB"}, "stock": 0}],
                           "last_id": "next"}),
            response(body={"products": [{"id": 2, "action_price": {"amount": "0", "currency": "RUB"}}], "last_id": ""}),
        ]
        with patch.object(rp.requests, "post", side_effect=pages) as post:
            products = rp.fetch_action_products(7, "candidates")
        self.assertEqual([p["id"] for p in products], [1, 2])
        self.assertEqual(products[0]["action_price"], 960.0)
        self.assertEqual(products[0]["max_action_price"], 1000.5)
        self.assertEqual(products[0]["currency"], "RUB")
        self.assertEqual(post.call_args_list[0].args[0], "https://api-seller.ozon.ru/v2/actions/candidates")
        self.assertEqual(post.call_args_list[1].kwargs["json"], {"action_id": 7, "limit": 100, "last_id": "next"})

    def test_boost_in_hundredths_of_percent_is_converted(self):
        # С 23.10.2026 Ozon присылает бустинг «Эластичного» в сотых долях процента.
        product = rp.normalize_action_product({"current_boost": 1400, "min_boost": 1400, "max_boost": 4900})
        self.assertEqual((product["current_boost"], product["min_boost"], product["max_boost"]), (14, 14, 49))
        product = rp.normalize_action_product({"current_boost": 55, "min_boost": 15, "max_boost": 55})
        self.assertEqual((product["current_boost"], product["max_boost"]), (55, 55))

    def test_activation_uses_update_with_money_objects_and_no_stock_for_boosting(self):
        body = {"active_product_ids": ["11"], "rejected": [{"product_id": 12, "reason": "price too high"}], "warnings": []}
        with patch.object(rp.requests, "post", return_value=response(body=body)) as post:
            result = rp.activate_products_in_action(7, [
                {"product_id": 11, "action_price": 1459, "stock": 0},
                {"product_id": 12, "action_price": 999.5, "stock": 3},
            ])
        self.assertEqual(post.call_args.args[0], "https://api-seller.ozon.ru/v1/actions/products/update")
        self.assertEqual(post.call_args.kwargs["json"], {"action_id": 7, "products": [
            {"product_id": 11, "action_price": {"amount": "1459.00", "currency": "RUB"}},
            {"product_id": 12, "action_price": {"amount": "999.50", "currency": "RUB"}, "stock": 3},
        ]})
        self.assertEqual(result["product_ids"], [11])
        self.assertEqual(result["rejected"], [{"product_id": 12, "reason": "price too high"}])
        self.assertEqual(result["errors"], [])

    def test_blocked_deactivation_is_reported_in_plain_words(self):
        blocked = response(400, {"code": 9, "message": "the method is blocked due to unified price logic"})
        with patch.object(rp.requests, "post", return_value=blocked) as post:
            result = rp.deactivate_products_in_action(7, [11, 12])
        self.assertEqual(post.call_args.args[0], "https://api-seller.ozon.ru/v2/actions/products/deactivate")
        self.assertEqual(post.call_args.kwargs["json"], {"action_id": 7, "product_ids": ["11", "12"]})
        self.assertEqual(result["product_ids"], [])
        self.assertEqual(len(result["errors"]), 1)
        self.assertIn("единую цену", result["errors"][0])


class ActionExitTests(unittest.TestCase):
    def setUp(self):
        PromosV2Tests.setUp(self)

    def test_exit_price_is_minimal_but_above_the_promo_limit(self):
        # С ценой, равной лимиту, товар в акции остаётся — нужен хотя бы лимит + 1 ₽.
        self.assertEqual(rp.action_exit_price(2040, 2293), 2293)
        self.assertEqual(rp.action_exit_price(2040, 1800), 2041)
        self.assertEqual(rp.action_exit_price(1630.4, 1700.2), 1701)
        self.assertEqual(rp.action_exit_price(2280, None), 2281)
        self.assertIsNone(rp.action_exit_price(None, 1800))
        self.assertIsNone(rp.action_exit_price(0, 1800))

    def test_exit_price_does_not_lower_the_current_product_price(self):
        # Артикул 1302: Ozon сам добавил товар в будущую акцию по 936 (лимит 936), минимальная 945, товар продаётся по 978.
        self.assertEqual(rp.action_exit_price(936, 945, 978), 978)
        # Акция идёт: цена товара равна цене в акции (ниже лимита) — выходим как раньше.
        self.assertEqual(rp.action_exit_price(936, 945, 936), 945)
        self.assertEqual(rp.action_exit_price(936, 900, None), 937)

    def test_removal_plan_shows_start_date_ozon_auto_add_and_keeps_price(self):
        import price_management as pm
        actions = [{"id": 4410985, "name": "Максимальный бустинг: усиление 11.11", "action_type": "STOCK_DISCOUNT",
                    "fixed_boost": None, "date_start": "2026-10-27T21:00:00Z"}]
        product = {"id": 1376477022, "price": 978.0, "action_price": 936.0, "max_action_price": 936.0, "add_mode": "AUTO"}
        context = {"limits": {"1302": (945.0, 1603.0)}, "pid_to_offer": {1376477022: "1302"}, "actions": actions}
        with patch.object(pm, "_load_plan_context", return_value=(context, None)), \
                patch.object(pm, "fetch_action_products", return_value=[product]), \
                patch.object(pm, "_attach_visibility", return_value=False), \
                patch.object(pm, "_moscow_today", return_value="2026-10-22"):
            plan = pm.build_action_removal_plan(Path("."))
        item = plan["items"][0]
        self.assertEqual((item["exit_price"], item["product_price"]), (978, 978.0))
        self.assertEqual(item["action_start"], "2026-10-28")
        self.assertTrue(item["added_by_ozon"])
        self.assertIn("Акция начнётся 28.10", item["note"])
        self.assertIn("Ozon добавил товар сам", item["note"])

    def test_plan_shows_the_unified_exit_price(self):
        import price_management as pm
        items = [{"kind": "remove", "product_id": 11, "exit_price": 2293}, {"kind": "remove", "product_id": 11, "exit_price": 2417},
                 {"kind": "remove", "product_id": 12, "exit_price": 1003}, {"kind": "add", "product_id": 11, "action_price": 1500}]
        pm._unify_plan_prices(items)
        self.assertEqual([item.get("exit_price") for item in items], [2417, 2417, 1003, None])
        # Добавление по 1500 вернуло бы товар 11 в акцию с лимитом ниже 2417 — снимаем его с отметки.
        self.assertFalse(items[3]["selected"])
        self.assertTrue(items[3]["conflict"])

    def test_plan_gives_one_entry_price_per_product(self):
        import price_management as pm
        items = [{"kind": "add", "product_id": 11, "action_price": 1200, "selected": True},
                 {"kind": "raise", "product_id": 11, "action_price": 1100, "selected": True},
                 {"kind": "remove", "product_id": 11, "exit_price": 1001},
                 {"kind": "add", "product_id": 12, "action_price": 1500, "selected": True}]
        pm._unify_plan_prices(items)
        self.assertEqual([items[0]["action_price"], items[1]["action_price"], items[3]["action_price"]], [1100, 1100, 1500])
        # Вывод идёт той же ценой, что и добавление: она выше лимита, цена товара меняется один раз.
        self.assertEqual(items[2]["exit_price"], 1100)
        self.assertTrue(all(item.get("selected", True) for item in items))

    def test_add_plan_keeps_one_price_within_every_action_limit(self):
        import price_management as pm
        actions = [{"id": 1, "name": "A", "action_type": "STOCK_DISCOUNT", "fixed_boost": 30.0},
                   {"id": 2, "name": "B", "action_type": "STOCK_DISCOUNT", "fixed_boost": 50.0},
                   {"id": 3, "name": "C", "action_type": "STOCK_DISCOUNT", "fixed_boost": 20.0}]
        data = {
            (1, "candidates"): [{"id": 11, "action_price": 0, "max_action_price": 1180}],
            (2, "candidates"): [{"id": 11, "action_price": 0, "max_action_price": 1090}],
            # Товар уже в акции C по цене не ниже минимальной: её лимит тоже ограничивает общую цену.
            (3, "products"): [{"id": 11, "action_price": 1050, "max_action_price": 1060}],
        }
        context = {"limits": {"X": (1000.0, 1200.0)}, "pid_to_offer": {11: "X"}, "actions": actions}
        with patch.object(pm, "_load_plan_context", return_value=(context, None)), \
                patch.object(pm, "fetch_action_products", side_effect=lambda action_id, endpoint: data.get((action_id, endpoint), [])), \
                patch.object(pm, "_attach_visibility", return_value=False):
            plan = pm.build_action_add_plan(Path("."))
        self.assertEqual({(item["action_name"], item["kind"], item["action_price"]) for item in plan["items"]},
                         {("A", "add", 1060), ("B", "add", 1060)})

    def test_add_plan_leaves_costly_action_unchecked_and_keeps_others_price(self):
        import price_management as pm
        actions = [{"id": 1, "name": "A", "action_type": "STOCK_DISCOUNT", "fixed_boost": 30.0},
                   {"id": 2, "name": "B", "action_type": "STOCK_DISCOUNT", "fixed_boost": 50.0}]
        data = {(1, "candidates"): [{"id": 11, "action_price": 0, "max_action_price": 1180}],
                (2, "candidates"): [{"id": 11, "action_price": 0, "max_action_price": 1050},
                                    {"id": 12, "action_price": 0, "max_action_price": 900}]}
        context = {"limits": {"X": (1000.0, 1200.0), "Y": (1000.0, 1200.0)}, "pid_to_offer": {11: "X", 12: "Y"}, "actions": actions}
        with patch.object(pm, "_load_plan_context", return_value=(context, None)), \
                patch.object(pm, "fetch_action_products", side_effect=lambda action_id, endpoint: data.get((action_id, endpoint), [])), \
                patch.object(pm, "_attach_visibility", return_value=False):
            plan = pm.build_action_add_plan(Path("."))
        items = {item["action_name"]: item for item in plan["items"]}
        # Y не помещается ни в одну акцию; X идёт в A по 1180, а B (1050) — только если отметить.
        self.assertEqual(set(items), {"A", "B"})
        self.assertEqual((items["A"]["action_price"], items["A"]["selected"]), (1180, True))
        self.assertEqual((items["B"]["action_price"], items["B"]["selected"]), (1050, False))
        self.assertIn("с 1 180 до 1 050 ₽", items["B"]["note"])

    def test_add_plan_execution_sends_the_same_price_to_every_action(self):
        import price_management as pm
        plan = {"items": [
            {"kind": "add", "action_id": 1, "action_name": "A", "product_id": 11, "offer_id": "X", "action_price": 1180, "stock": 0},
            {"kind": "add", "action_id": 2, "action_name": "B", "product_id": 11, "offer_id": "X", "action_price": 1090, "stock": 0},
        ]}
        sent = {"product_ids": [11], "rejected": [], "errors": []}
        with patch.object(pm, "activate_products_in_action", return_value=sent) as activate, \
                patch.object(pm, "fetch_action_products", return_value=[{"id": 11, "action_price": 1090}]):
            result = pm.execute_action_plan(plan, "add")
        self.assertEqual([call.args[1][0]["action_price"] for call in activate.call_args_list], [1090, 1090])
        self.assertTrue(result["ok"])

    def test_execution_skips_entry_that_would_undo_a_removal(self):
        import price_management as pm
        plan = {"items": [
            {"kind": "remove", "action_id": 1, "action_name": "A", "product_id": 11, "offer_id": "X", "exit_price": 1301, "selected": True},
            {"kind": "add", "action_id": 2, "action_name": "B", "product_id": 11, "offer_id": "X", "action_price": 1200, "stock": 0, "selected": True},
        ]}
        sent = {"product_ids": [], "rejected": [], "errors": []}
        with patch.object(pm, "activate_products_in_action", return_value=sent) as activate, \
                patch.object(pm, "fetch_action_products", return_value=[]):
            result = pm.execute_action_plan(plan, "remove")
        activate.assert_called_once_with(1, [{"product_id": 11, "action_price": 1301}])
        self.assertFalse(result["ok"])
        self.assertIn("вернёт товар в акцию", result["errors"][0])

    def test_removal_plan_execution_raises_price_and_checks_the_product_left(self):
        import price_management as pm
        plan = {"items": [{"kind": "remove", "action_id": 7, "action_name": "Максимальный бустинг", "product_id": 11,
                           "offer_id": "A", "exit_price": 2041, "selected": True}]}
        sent = {"product_ids": [], "deactivated_product_ids": [11], "rejected": [], "errors": []}
        with patch.object(pm, "activate_products_in_action", return_value=sent) as activate, \
                patch.object(pm, "fetch_action_products", return_value=[]):
            result = pm.execute_action_plan(plan, "remove")
        activate.assert_called_once_with(7, [{"product_id": 11, "action_price": 2041}])
        self.assertTrue(result["ok"])
        self.assertEqual(result["done"], 1)


if __name__ == "__main__":
    unittest.main()
