"""Offline checks for the limit price slider: thresholds, participation at a price, auto price, what is sent."""
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import limit_price as lp  # noqa: E402

ELASTIC = {"id": 1, "name": "Эластичный бустинг", "action_type": "ELASTIC_BOOSTING", "fixed_boost": None, "date_start": "2025-03-19T21:00:00Z"}
STOCK = {"id": 2, "name": "Максимальный бустинг: усиление 11.11", "action_type": "STOCK_DISCOUNT", "fixed_boost": 55.0,
         "date_start": "2026-10-27T21:00:00Z"}
SUPER = {"id": 3, "name": "Супербустинг", "action_type": "STOCK_DISCOUNT", "fixed_boost": 65.0, "date_start": "2026-10-29T21:00:00Z"}


def board(participants=None, candidates=None):
    """Артикул 1302 из магазина 22.10.2026: в эластичном (порог 1250) и в усилении 11.11 (936), кандидат в супербустинг (937)."""
    context = {"limits": {"1302": (945.0, 1603.0)}, "pid_to_offer": {7: "1302"}, "actions": [ELASTIC, STOCK, SUPER]}
    base = {"id": 7, "price": 978.0, "min_seller_price": 850.0}
    loaded = {
        1: ([{**base, "action_price": 978.0, "max_action_price": 1250.0, "price_min_elastic": 1250.0, "price_max_elastic": 1044.0,
              "min_boost": 15, "max_boost": 55, "add_mode": "SELLER"}], []),
        2: ([{**base, "action_price": 936.0, "max_action_price": 936.0, "add_mode": "AUTO"}], []),
        3: ([], [{**base, "action_price": 0, "max_action_price": 937.0}]),
    }
    return lp.build_board(context, loaded, "2026-10-22")


class ThresholdTests(unittest.TestCase):
    def test_elastic_uses_minimum_boost_price(self):
        self.assertEqual(lp.threshold(ELASTIC, {"price_min_elastic": 1200, "max_action_price": 1250}), 1200)
        self.assertEqual(lp.threshold(ELASTIC, {"max_action_price": 1250}), 1250)
        self.assertEqual(lp.threshold(STOCK, {"price_min_elastic": 900, "max_action_price": 936}), 936)


class BoardTests(unittest.TestCase):
    def test_article_with_its_actions(self):
        article = board()[0]
        self.assertEqual((article["offer_id"], article["price"], article["ozon_min"]), ("1302", 978.0, 850.0))
        self.assertEqual([a["threshold"] for a in article["actions"]], [1250.0, 937.0, 936.0])
        by_id = {a["action_id"]: a for a in article["actions"]}
        self.assertTrue(by_id[2]["added_by_ozon"])
        self.assertEqual(by_id[2]["start"], "2026-10-28")  # ещё не началась
        self.assertIsNone(by_id[1]["start"])

    def test_one_price_changes_every_running_action_at_once(self):
        article = board()[0]
        changes = {a["action_id"]: (a["change"], a["in"], a["boost"]) for a in lp.evaluate(article, 937)}
        # 937: остаётся в эластичном (порог 1250, буст максимальный ниже 1044).
        self.assertEqual(changes[1], ("stay", True, 55))
        # Акции, которые ещё не начались, — отдельное состояние со своей ценой: ползунок участие в них не меняет.
        self.assertEqual(changes[2], ("future", True, 55.0))
        # Кандидат будущей акции: при 937 вошёл бы (порог 937), но это только подсказка — не отправляется.
        self.assertEqual(changes[3], ("future", True, 65.0))
        self.assertEqual(lp.evaluate(article, 1000)[1]["in"], False)

    def test_auto_price_stays_between_min_and_desired(self):
        article = board()[0]
        price = lp.auto_price(article)
        # Эластичный (1250) — участие по нормальной цене, держим его; 936/937 ниже минимальной 945 — недостижимы.
        self.assertEqual(price, 1250)
        self.assertTrue(945 <= price <= 1603)

    def test_auto_price_without_actions_in_range_is_desired(self):
        article = {**board()[0], "actions": [a for a in board()[0]["actions"] if a["action_id"] != 1]}
        self.assertEqual(lp.auto_price(article), 1603)


def simple_article(**overrides):
    """Одна идущая акция с порогом 1000 ₽ и бустингом 50%; цена 1000, мин. 900, жел. 1500; прибыль = 0,8·цена − 400."""
    article = {
        "offer_id": "A", "product_id": 1, "price": 1000.0, "ozon_min": None, "min_price": 900.0, "desired_price": 1500.0,
        "economics": {"cost": 300.0, "commission": 0.2, "logistics": 100.0},
        "actions": [{"action_id": 9, "name": "Скидка на сток", "action_type": "STOCK_DISCOUNT", "fixed_boost": 50.0, "start": None,
                     "threshold": 1000.0, "participating": True, "action_price": 1000.0, "added_by_ozon": False, "alert": False,
                     "stock": 0, "min_boost": None, "max_boost": None, "price_min_elastic": None, "price_max_elastic": None}],
    }
    article.update(overrides)
    return article


class BalanceTests(unittest.TestCase):
    def test_profit_and_total_boost(self):
        article = simple_article()
        self.assertEqual(lp.unit_profit(article, 1000), 400)
        self.assertEqual((lp.total_boost(article, 1000), lp.total_boost(article, 1001)), (50.0, 0.0))

    def test_balance_keeps_the_boost_when_it_pays(self):
        article = simple_article()
        # Ниже 1000: (0,8P − 400) × 1,5 × (1000/P)² — максимум ровно на пороге; выше порога бустинг пропадает.
        self.assertEqual(lp.balance_price(article, elasticity=2.0), 1000)
        self.assertGreater(lp.balance_score(article, 1000, 2.0), lp.balance_score(article, 1500, 2.0))

    def test_balance_raises_price_when_buyers_barely_react_and_boost_is_small(self):
        article = simple_article()
        article["actions"][0]["fixed_boost"] = 5.0
        self.assertEqual(lp.balance_price(article, elasticity=1.0), 1500)

    def test_sale_adds_thirty_percent_on_top_of_the_boost(self):
        self.assertTrue(lp.is_sale("Максимальный бустинг 11.11"))
        self.assertTrue(lp.is_sale("Чёрная пятница"))
        self.assertFalse(lp.is_sale("Эластичный бустинг. Без ограничения срока действия"))
        article = simple_article()
        article["actions"][0].update(name="Чёрная пятница", sale=True)
        self.assertEqual((lp.sale_bonus(article, 1000), lp.sale_bonus(article, 1001)), (30.0, 0.0))
        # (1 + (50 + 30)/100) при текущей цене — в 1,8 раза больше продаж, чем без акций.
        self.assertAlmostEqual(lp.sales_factor(article, 1000, 2.0), 1.8)
        # Не начавшаяся распродажа ещё не действует.
        article["actions"][0]["start"] = "2026-11-27"
        self.assertEqual(lp.sale_bonus(article, 1000), 0.0)

    def test_ozon_points_for_discount_add_to_profit(self):
        import recommended_prices as rp
        description = ("Новая специальная акция на ограниченный ассортимент товаров <b>с дополнительными баллами за скидки 3%</b>. "
                       "После продажи начислим вам эту сумму баллами, которыми оплачивается тариф за продажу")
        self.assertEqual(rp.parse_points_percent(description), 3.0)
        self.assertIsNone(rp.parse_points_percent("Специальная акция с максимальным бустингом в поиске"))
        article = simple_article()
        article["actions"][0]["points"] = 3.0
        # 0,8 × 1000 − 400 = 400 ₽ и ещё 3% от 1000 ₽ баллами — 430 ₽; выше порога акции баллов нет.
        self.assertAlmostEqual(lp.unit_profit(article, 1000), 430.0)
        self.assertAlmostEqual(lp.unit_profit(article, 1001), 400.8)
        article["actions"][0]["start"] = "2026-10-30"
        self.assertAlmostEqual(lp.unit_profit(article, 1000), 400.0)

    def test_without_economics_falls_back_to_keeping_participation(self):
        self.assertEqual(lp.balance_price(simple_article(economics=None)), lp.auto_price(simple_article(economics=None)))


class OzonTransitionTests(unittest.TestCase):
    def test_participant_above_the_limit_stays_at_its_price(self):
        # 1102 23.10.2026: в «Эластичном» по 1 421 ₽ при пороге 1 278 ₽ и схлопнутой шкале, Ozon называет бустинг 14%.
        context = {"limits": {"1102": (1277.0, 2165.0)}, "pid_to_offer": {21: "1102"}, "actions": [ELASTIC]}
        product = {"id": 21, "price": 1421.0, "action_price": 1421.0, "max_action_price": 1278.0, "price_min_elastic": 1278.0,
                   "price_max_elastic": 1278.0, "current_boost": 14.0, "min_boost": 14.0, "max_boost": 49.0, "add_mode": "AUTO"}
        article = lp.build_board(context, {1: ([product], [])}, "2026-10-23")[0]
        entry = article["actions"][0]
        self.assertEqual((entry["threshold"], entry["ozon_threshold"], entry["above_limit"]), (1421.0, 1278.0, True))
        state = lp.evaluate(article, 1421)[0]
        self.assertEqual((state["change"], state["boost"]), ("stay", 14.0))
        self.assertEqual(lp.evaluate(article, 1500)[0]["change"], "exit")


class ApplyTests(unittest.TestCase):
    def test_sends_price_to_every_action_it_stays_enters_or_leaves(self):
        article = board()[0]
        items = {(i["action_id"], i["change"]) for i in lp.plan_changes([article], {"1302": 1000})}
        # 1000: эластичный — остаётся с новой ценой; будущие акции (усиление с 28.10, супербустинг с 30.10) не трогаем.
        self.assertEqual(items, {(1, "stay")})

    def test_running_action_is_left_when_price_goes_above_its_threshold(self):
        article = board()[0]
        running = {**article, "actions": [{**a, "start": None} for a in article["actions"]]}
        items = {(i["action_id"], i["change"]) for i in lp.plan_changes([running], {"1302": 1000})}
        self.assertEqual(items, {(1, "stay"), (2, "exit")})

    def test_unchanged_participation_is_not_sent(self):
        article = board()[0]
        running = {**article, "actions": [{**a, "start": None} for a in article["actions"]]}
        self.assertEqual([i["action_id"] for i in lp.plan_changes([running], {"1302": 936}) if i["action_id"] == 2], [])

    def test_apply_checks_participation_after_sending(self):
        article = board()[0]
        article = {**article, "actions": [{**a, "start": None} for a in article["actions"]]}
        actual = {1: [{"id": 7, "action_price": 1000.0}], 2: [{"id": 7, "action_price": 936.0}]}
        with patch.object(lp, "activate_products_in_action", return_value={"rejected": [], "errors": []}) as activate, \
                patch.object(lp, "fetch_action_products", side_effect=lambda action_id, endpoint: actual.get(action_id, [])):
            result = lp.apply_prices([article], {"1302": 1000}, sleep=lambda seconds: None)
        self.assertEqual(activate.call_count, 2)
        # Из усиления товар не вышел — это видно только по перечитанному участию.
        self.assertEqual((result["done"], [f["action_id"] for f in result["failed"]]), (1, [2]))
        self.assertFalse(result["ok"])


class DirectPriceTests(unittest.TestCase):
    def test_article_outside_actions_gets_the_product_price(self):
        # 7002 23.10.2026: все пороги ниже минимальной, идущих акций нет — цену передаём как цену товара.
        article = simple_article(offer_id="7002", product_id=70, price=850.0)
        article["actions"][0].update(participating=False, action_price=None, threshold=800.0)
        self.assertEqual(lp.plan_changes([article], {"7002": 1074}), [])
        self.assertEqual([a["offer_id"] for a in lp.direct_price_articles([article], {"7002": 1074}, [])], ["7002"])
        with patch.object(lp, "import_product_prices", return_value={}) as send, \
                patch.object(lp, "get_product_prices_from_ozon", return_value=({"7002": 1074.0}, {})):
            result = lp.apply_prices([article], {"7002": 1074}, sleep=lambda seconds: None)
        send.assert_called_once_with([{"offer_id": "7002", "product_id": 70, "price": 1074}])
        self.assertEqual((result["done"], result["expected"], result["ok"]), (1, 1, True))

    def test_unchanged_or_action_routed_prices_are_not_sent_directly(self):
        article = simple_article()
        self.assertEqual(lp.direct_price_articles([article], {"A": 1000}, []), [])
        self.assertEqual(lp.direct_price_articles([article], {"A": 990}, [{"offer_id": "A"}]), [])


class OzonResponseTests(unittest.TestCase):
    def test_no_changes_is_success_and_ozon_warnings_become_notices(self):
        # 5004 23.10.2026: цена 1180 ушла в «усиление.», Ozon опустил ограничение; в «11.11» та же цена — «нет изменений».
        actions = [{**STOCK, "id": 4375943, "name": "Максимальный бустинг: усиление.", "date_start": "2026-10-11T21:00:00Z"},
                   {**STOCK, "id": 4400847, "name": "Максимальный бустинг 11.11", "date_start": "2026-10-21T21:00:00Z"}]
        context = {"limits": {"5004": (1219.0, 2068.0)}, "pid_to_offer": {2591798900: "5004"}, "actions": actions}
        base = {"id": 2591798900, "price": 1200.0, "min_seller_price": 1250.0, "action_price": 1200.0, "add_mode": "SELLER"}
        loaded = {4375943: ([{**base, "max_action_price": 1300.0}], []), 4400847: ([{**base, "max_action_price": 1250.0}], [])}
        article = lp.build_board(context, loaded, "2026-10-23")[0]
        responses = {
            4375943: {"rejected": [], "errors": [], "warnings": [
                {"product_id": "2591798900", "reason": "Min price for product has been changed: new value = 1180.00"}]},
            4400847: {"rejected": [{"product_id": 2591798900, "reason": "Product update has no changes"}], "errors": [], "warnings": []},
        }
        actual = {4375943: [{"id": 2591798900, "action_price": 1180.0}], 4400847: [{"id": 2591798900, "action_price": 1180.0}]}
        with patch.object(lp, "activate_products_in_action", side_effect=lambda action_id, products: responses[action_id]), \
                patch.object(lp, "fetch_action_products", side_effect=lambda action_id, endpoint: actual[action_id]):
            result = lp.apply_prices([article], {"5004": 1180}, sleep=lambda seconds: None)
        self.assertEqual((result["done"], result["failed"]), (2, []))
        self.assertTrue(result["ok"])
        self.assertEqual([(n["kind"], n.get("value")) for n in result["notices"]], [("limit", 1180.0)])
        self.assertIn("ограничение для акций до 1 180 ₽", result["notices"][0]["text"])

    def test_exit_that_lands_a_few_seconds_later_is_success(self):
        # 5004 23.10.2026: выход по 1587; в «11.11» Ozon ответил «нет изменений», и сразу после запроса товар ещё там.
        actions = [{**STOCK, "id": 1, "name": "Максимальный бустинг: усиление.", "date_start": "2026-10-11T21:00:00Z"},
                   {**STOCK, "id": 2, "name": "Максимальный бустинг 11.11", "date_start": "2026-10-21T21:00:00Z"}]
        context = {"limits": {"5004": (1219.0, 2068.0)}, "pid_to_offer": {9: "5004"}, "actions": actions}
        base = {"id": 9, "price": 1180.0, "action_price": 1180.0, "add_mode": "SELLER"}
        article = lp.build_board(context, {1: ([{**base, "max_action_price": 1300.0}], []), 2: ([{**base, "max_action_price": 1250.0}], [])},
                                 "2026-10-23")[0]
        responses = {1: {"rejected": [], "errors": [], "warnings": []},
                     2: {"rejected": [{"product_id": 9, "reason": "Product update has no changes"}], "errors": [], "warnings": []}}
        calls = {"n": 0}

        def fetch(action_id, endpoint):
            calls["n"] += 1
            return [{"id": 9, "action_price": 1587.0}] if action_id == 2 and calls["n"] <= 2 else []

        pauses = []
        with patch.object(lp, "activate_products_in_action", side_effect=lambda action_id, products: responses[action_id]), \
                patch.object(lp, "fetch_action_products", side_effect=fetch):
            result = lp.apply_prices([article], {"5004": 1587}, sleep=pauses.append)
        self.assertEqual((result["done"], result["failed"]), (2, []))
        self.assertEqual(pauses, [3])

    def test_leaving_an_action_is_reported_as_expected(self):
        item = {"offer_id": "5004", "action_name": "Максимальный бустинг: усиление.", "expect_in": False}
        notice = lp.ozon_notice("The product is deactivated from the action. It no longer meets the conditions", item)
        self.assertEqual((notice["kind"], notice["expected"]), ("left", True))
        self.assertEqual(notice["text"], "5004: вышел из «Максимальный бустинг: усиление.»")


if __name__ == "__main__":
    unittest.main()
