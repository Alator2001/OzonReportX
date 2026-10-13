"""Offline checks for the organic sales analysis."""
import sys
import tempfile
import unittest
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import organic_analysis as oa  # noqa: E402

TODAY = date(2026, 10, 8)
STORE = {"ctr": 0.03, "cart_rate": 0.10, "cart_to_order": 0.35, "season_change": 0.0}


def product(**overrides):
    base = {
        "orders": 10.0, "orders_prev": 10.0, "organic": 8.0, "organic_prev": 8.0, "ad_orders": 2.0,
        "stock": 20, "days_since_sale": 1, "locality": {"local": 8, "total": 10, "recommended_supply": 0},
        "avg_price": 1000.0, "avg_price_prev": 1000.0, "price_index": {"color": "SUPER", "min_price": 1100.0},
        "actions": ["Акция"], "content_rating": 95.0, "content_missing": [],
        "ads": {"views": 3000.0, "clicks": 90.0, "to_cart": 9.0, "orders": 3.0}, "search": None,
    }
    base.update(overrides)
    return base


def steps_by_name(result):
    return {s["name"]: s for s in result["steps"]}


class PeriodTests(unittest.TestCase):
    def test_periods_are_equal_and_year_shift_keeps_weekdays(self):
        spans = oa.analysis_periods(TODAY, 28)
        self.assertEqual(spans["current"], (date(2026, 9, 10), TODAY))
        self.assertEqual(spans["previous"], (date(2026, 8, 13), date(2026, 9, 10)))
        self.assertEqual(spans["current_ly"][0].weekday(), spans["current"][0].weekday())

    def test_season_expectation_from_last_year(self):
        spans = oa.analysis_periods(TODAY, 28)
        rows = [(spans["previous_ly"][0] + timedelta(days=i), "A", 1) for i in range(28)]
        rows += [(spans["current_ly"][0] + timedelta(days=i), "A", 2) for i in range(28)]
        orders = pd.DataFrame(rows, columns=["date", "offer_id", "qty"])
        orders["date"] = pd.to_datetime(orders["date"])
        self.assertAlmostEqual(oa.season_expectation(orders, spans), 1.0)   # год назад продажи выросли вдвое


class ProductTests(unittest.TestCase):
    def test_healthy_product_has_no_main_problem(self):
        result = oa.analyze_product(product(), STORE)
        self.assertIsNone(result["main_problem"])
        self.assertEqual(len(result["steps"]), 7)

    def test_no_stock_is_the_top_problem(self):
        result = oa.analyze_product(product(stock=0, organic=2.0), STORE)
        self.assertEqual(steps_by_name(result)["Наличие и доставка"]["status"], "bad")
        self.assertIn("Пополните остаток", result["advice"][0])

    def test_drop_within_season_is_not_blamed_on_product(self):
        result = oa.analyze_product(product(organic=5.0), {**STORE, "season_change": -0.4})
        self.assertEqual(steps_by_name(result)["Продажи"]["status"], "ok")
        self.assertIn("в пределах сезона", steps_by_name(result)["Продажи"]["text"])

    def test_drop_against_rising_season_is_flagged(self):
        result = oa.analyze_product(product(organic=4.0), {**STORE, "season_change": 0.5})
        sales = steps_by_name(result)["Продажи"]
        self.assertEqual(sales["status"], "bad")
        self.assertIn("По сезону ждали +50%, а органика -50%", sales["text"])

    def test_bad_price_index_from_other_seller_and_price_rise(self):
        index = {"color": "RED", "value": 1.2, "driver": "ozon", "driver_price": 800.0}
        result = oa.analyze_product(product(price_index=index, price_now=1200.0, min_price=700.0, avg_price=1200.0, organic=5.0), STORE)
        price = steps_by_name(result)["Цена"]
        self.assertEqual(price["status"], "bad")
        self.assertIn("другой продавец на Ozon за 800 ₽", price["text"])
        self.assertIn("примерно до 1 000 ₽", price["text"])                  # 1200 / 1,2
        self.assertTrue(any("снизьте цену примерно до 1 000 ₽" in a for a in result["advice"]))
        self.assertTrue(any("выросла на 20%" in a for a in result["advice"]))

    def test_cheapest_is_own_listing_elsewhere(self):
        index = {"color": "RED", "value": 1.32, "driver": "self", "driver_price": 692.0}
        result = oa.analyze_product(product(price_index=index, price_now=1545.0, min_price=1343.0), STORE)
        price = steps_by_name(result)["Цена"]
        self.assertIn("ваш же товар на другой площадке за 692 ₽", price["text"])
        self.assertIn("ниже вашей минимальной цены 1 343 ₽", price["text"])  # 1545 / 1,32 ≈ 1170
        top = result["advice"][0]
        self.assertIn("конкурируете сами с собой", top)
        self.assertIn("поднимите цену на другой площадке выше 692 ₽", top)

    def test_target_below_floor_for_real_competitor(self):
        index = {"color": "RED", "value": 1.5, "driver": "external", "driver_price": 500.0}
        result = oa.analyze_product(product(price_index=index, price_now=1500.0, min_price=1200.0), STORE)
        self.assertTrue(any("конкурировать ценой не выйдет" in a for a in result["advice"]))

    def test_parse_price_index_detects_own_listing(self):
        parsed = oa.parse_price_index({
            "color_index": "RED",
            "ozon_index_data": {"min_price": 0, "price_index_value": 0},
            "external_index_data": {"min_price": 692, "price_index_value": 1.32},
            "self_marketplaces_index_data": {"min_price": 692, "price_index_value": 1.32},
        })
        self.assertEqual((parsed["driver"], parsed["driver_price"], parsed["value"]), ("self", 692.0, 1.32))
        competitor = oa.parse_price_index({
            "color_index": "YELLOW",
            "ozon_index_data": {"min_price": 950, "price_index_value": 1.05},
            "external_index_data": {"min_price": 0, "price_index_value": 0},
            "self_marketplaces_index_data": {"min_price": 0, "price_index_value": 0},
        })
        self.assertEqual(competitor["driver"], "ozon")

    def test_price_for_good_index(self):
        self.assertEqual(oa.price_for_good_index(1134, 1.07), 1060)
        self.assertIsNone(oa.price_for_good_index(1000, 0.9))
        self.assertIsNone(oa.price_for_good_index(None, 1.2))

    def test_low_content_rating_lists_tasks(self):
        result = oa.analyze_product(product(content_rating=60.0, content_missing=["Медиа — добавить от 8 изображений"]), STORE)
        self.assertEqual(steps_by_name(result)["Карточка"]["status"], "bad")
        self.assertTrue(any("от 8 изображений" in a for a in result["advice"]))

    def test_slow_delivery_recommends_supply(self):
        result = oa.analyze_product(product(locality={"local": 0, "total": 6, "recommended_supply": 4}), STORE)
        self.assertEqual(steps_by_name(result)["Наличие и доставка"]["status"], "warn")
        self.assertTrue(any("Ozon советует 4 шт." in a for a in result["advice"]))

    def test_weak_funnel_from_ad_traffic(self):
        result = oa.analyze_product(product(ads={"views": 9000.0, "clicks": 90.0, "to_cart": 2.0, "orders": 1.0}), STORE)
        funnel = steps_by_name(result)["Воронка"]
        self.assertEqual(funnel["status"], "warn")
        self.assertIn("редко кликают", funnel["text"])
        self.assertIn("не кладут в корзину", funnel["text"])

    def test_ad_dependency(self):
        result = oa.analyze_product(product(orders=10.0, ad_orders=9.0, organic=1.0), STORE)
        self.assertEqual(steps_by_name(result)["Реклама"]["status"], "warn")
        self.assertTrue(any("ключевые слова" in a for a in result["advice"]))

    def test_search_step_with_and_without_premium(self):
        free = oa.analyze_product(product(search={"searched": 500, "viewed": None, "position": None, "queries": []}), STORE)
        self.assertIn("только с подпиской Premium", steps_by_name(free)["Поиск"]["text"])
        premium = oa.analyze_product(product(search={"searched": 500, "viewed": 120, "position": 80.0,
                                                     "queries": [{"query": "доска разделочная", "position": 75.0}]}), STORE)
        search = steps_by_name(premium)["Поиск"]
        self.assertEqual(search["status"], "warn")
        self.assertIn("«доска разделочная» (поз. 75)", search["text"])

    def test_store_benchmarks(self):
        bench = oa.store_funnel_benchmarks({"1": {"views": 1000, "clicks": 30, "to_cart": 3, "orders": 1},
                                            "2": {"views": 1000, "clicks": 10, "to_cart": 1, "orders": 1}})
        self.assertAlmostEqual(bench["ctr"], 0.02)
        self.assertAlmostEqual(bench["cart_rate"], 0.1)


class SnapshotTests(unittest.TestCase):
    def test_events_between_snapshots(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            oa.save_snapshot(root, date(2026, 10, 1), {"A": {"price": 1000, "stock": 5, "index": "SUPER", "actions": 1, "rating": 70}})
            oa.save_snapshot(root, date(2026, 10, 3), {"A": {"price": 1200, "stock": 0, "index": "BAD", "actions": 0, "rating": 80}})
            oa.save_snapshot(root, date(2026, 10, 5), {"A": {"price": 1200, "stock": 7, "index": "BAD", "actions": 0, "rating": 80}})
            events = oa.snapshot_events(oa.load_snapshots(root, date(2026, 9, 1)), "A")
        self.assertEqual(events, [
            "03.10: цена 1 000 ₽ → 1 200 ₽",
            "03.10: закончился остаток",
            "03.10: индекс цен «очень выгодная» → «невыгодная»",
            "03.10: акций 1 → 0",
            "03.10: контент-рейтинг 70 → 80",
            "05.10: товар снова в наличии (7 шт.)",
        ])


if __name__ == "__main__":
    unittest.main()
