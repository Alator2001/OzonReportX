"""Offline checks for the step-by-step CPC analysis."""
import sys
import unittest
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import ads_cpc  # noqa: E402

SETTINGS = {**ads_cpc.DEFAULT_SETTINGS}
BENCH = {"ctr": 0.03, "cart_rate": 0.10, "cart_to_order": 0.35}


def item(**overrides):
    base = {
        "strategy": "TARGET_BIDS", "current_bid": 4.0, "clicks": 200.0, "orders_total": 8.0, "views": 6000.0,
        "to_cart": 20.0, "spend": 800.0, "revenue": 7000.0, "product_gmv": 9000.0, "price": 900.0,
        "profit": 160.0, "cost": 150.0, "stock": 10, "min_bid": 0.23, "days_in_campaign": 21,
    }
    base.update(overrides)
    return base


def analyze(**overrides):
    return ads_cpc.analyze_item(item(**overrides), 0.04, BENCH, SETTINGS)


class StepTests(unittest.TestCase):
    def test_mature_target_is_share_of_breakeven_with_step_limit(self):
        result = analyze()
        self.assertEqual(result["stage"], "mature")
        self.assertAlmostEqual(result["breakeven_bid"], 6.4)       # 4% × 160 ₽
        self.assertAlmostEqual(result["target_bid"], 2.56)         # × 40%
        self.assertEqual(result["recommended_bid"], 3.0)           # шаг −25%
        self.assertEqual(result["action"], "bid_down")
        self.assertEqual([s["name"] for s in result["steps"]],
                         ["Экономика", "Наличие", "Стадия", "Воронка", "Окупаемость", "Ставка"])

    def test_economics_step_explains_allowed_drr(self):
        text = analyze()["steps"][0]["text"]
        self.assertIn("маржа 17,8%", text)          # 160 / 900
        self.assertIn("до 64 ₽ на заказ", text)     # 160 × 40%
        self.assertIn("ДРР до 7,1%", text)          # 17,8% × 40%

    def test_test_stage_keeps_profit_limit_but_cuts_gently(self):
        result = analyze(clicks=100.0, orders_total=4.0, views=3000.0, to_cart=10.0, current_bid=5.0, days_in_campaign=10)
        self.assertEqual(result["stage"], "test")
        self.assertAlmostEqual(result["target_bid"], 2.56)         # лимит 40% действует и в тесте
        self.assertEqual(result["recommended_bid"], 4.5)           # но шаг вниз только −10%
        self.assertIn("товар на тесте", result["steps"][-1]["text"])

    def test_test_stage_does_not_cut_for_missing_orders(self):
        result = analyze(clicks=120.0, orders_total=0.0, views=3000.0, to_cart=6.0, days_in_campaign=10)
        self.assertNotIn("без заказов", result["steps"][-1]["text"])

    def test_test_ends_after_28_days_even_with_few_clicks(self):
        result = analyze(clicks=100.0, orders_total=4.0, views=3000.0, to_cart=10.0, days_in_campaign=30)
        self.assertEqual(result["stage"], "mature")
        self.assertIn("Мало трафика", result["steps"][2]["text"])

    def test_variant_purchases_do_not_blame_the_card(self):
        result = analyze(orders=0.0, model_orders=4.0, orders_total=4.0, to_cart=4.0)   # корзина 2% против 10%
        funnel = result["steps"][3]
        self.assertEqual(funnel["status"], "ok")
        self.assertIn("соседний вариант", funnel["text"])
        self.assertFalse(any("карточке" in a for a in result["advice"]))

    def test_cart_to_order_uses_direct_orders_only(self):
        result = analyze(orders_total=8.0, orders=2.0, to_cart=4.0)
        self.assertIn("из корзины покупают 50%", result["steps"][3]["text"])

    def test_low_profit_relative_to_cost_is_still_advertised(self):
        result = analyze(profit=130.0)                 # прибыль меньше себестоимости 150 ₽ — больше не повод убирать
        self.assertNotEqual(result["action"], "remove")
        self.assertEqual(result["steps"][0]["status"], "ok")

    def test_loss_making_product_is_removed(self):
        result = analyze(profit=-20.0)
        self.assertEqual(result["action"], "remove")
        self.assertEqual(result["steps"][0]["status"], "bad")

    def test_out_of_stock_is_info(self):
        self.assertEqual(analyze(stock=0)["action"], "info")

    def test_other_strategy_is_info_only(self):
        result = analyze(strategy="MAX_CLICKS")
        self.assertEqual(result["action"], "info")
        self.assertEqual(result["recommended_bid"], 4.0)

    def test_many_clicks_without_orders_cut_bid(self):
        result = analyze(orders_total=0.0, to_cart=5.0)
        self.assertIn("без заказов", result["steps"][-1]["text"])
        self.assertEqual(result["recommended_bid"], 3.0)

    def test_organic_products_are_cut_gently(self):
        result = analyze(product_gmv=20000.0, revenue=7000.0)
        self.assertEqual(result["recommended_bid"], 3.6)  # −10% вместо −25%

    def test_weak_funnel_steps_are_explained(self):
        result = analyze(views=20000.0, to_cart=4.0)       # CTR 1% и корзина 2% против 3% и 10%
        funnel = result["steps"][3]
        self.assertEqual(funnel["status"], "warn")
        self.assertTrue(any("главное фото" in a for a in result["advice"]))
        self.assertTrue(any("карточке" in a for a in result["advice"]))

    def test_overpriced_competitive_bid_is_flagged(self):
        result = analyze(competitive_bid=15.0)
        self.assertTrue(any("не разогнать" in a for a in result["advice"]))

    def test_raises_bid_within_limit_when_profitable(self):
        result = analyze(current_bid=1.0, orders_total=20.0)
        self.assertEqual(result["action"], "bid_up")
        self.assertEqual(result["recommended_bid"], 1.25)

    def test_never_below_min_bid(self):
        result = analyze(current_bid=0.6, orders_total=0.0, min_bid=0.5)
        self.assertEqual(result["recommended_bid"], 0.5)

    def test_micro_conversion(self):
        self.assertEqual(ads_cpc.to_micro(3.07), 3_070_000)


class CampaignAdviceTests(unittest.TestCase):
    def campaign_items(self, spend=500.0):
        return [{**item(spend=spend, orders_total=10.0), "stage": "mature"}, {**item(spend=0.0, orders_total=0.0), "stage": "test"}]

    def test_budget_runs_out_and_campaign_pays_off(self):
        notes = ads_cpc.campaign_advice({"weeklyBudget": "1000000000"}, self.campaign_items(), 600.0, date(2026, 10, 6),
                                        date(2026, 10, 8), None, SETTINGS)
        budget = next(n for n in notes if n["name"] == "Бюджет")
        self.assertIn("поднять бюджет на 20%", budget["text"])   # 600 ₽ за 3 дня → 1400 ₽ за неделю > 1000 ₽

    def test_season_rise_and_structure(self):
        notes = ads_cpc.campaign_advice({"weeklyBudget": "0"}, self.campaign_items(), None, None, date(2026, 10, 8),
                                        {"index": 2.5, "before": 40, "after": 100}, SETTINGS)
        names = {n["name"]: n for n in notes}
        self.assertIn("за 1–2 недели до пика", names["Сезон"]["text"])
        self.assertIn("отдельной кампании", names["Структура"]["text"])

    def test_week_start(self):
        self.assertEqual(ads_cpc._week_start(date(2026, 10, 8), "THURSDAY"), date(2026, 10, 8))
        self.assertEqual(ads_cpc._week_start(date(2026, 10, 7), "THURSDAY"), date(2026, 10, 1))


if __name__ == "__main__":
    unittest.main()
