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

    def test_plan_shows_the_unified_exit_price(self):
        import price_management as pm
        items = [{"kind": "remove", "product_id": 11, "exit_price": 2293}, {"kind": "remove", "product_id": 11, "exit_price": 2417},
                 {"kind": "remove", "product_id": 12, "exit_price": 1003}, {"kind": "add", "product_id": 11, "action_price": 1500}]
        pm._unify_exit_prices(items)
        self.assertEqual([item.get("exit_price") for item in items], [2417, 2417, 1003, None])

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
