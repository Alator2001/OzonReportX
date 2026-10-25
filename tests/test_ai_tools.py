"""Offline checks for the AI assistant tools over the warehouse: metrics, overview, product profile, SQL guard, live Ozon cache."""
import sys
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import MagicMock, patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "tests"))
import ai_tools  # noqa: E402
from test_ai_warehouse import WarehouseTestCase, ozon_report  # noqa: E402


def rows(result, name=""):
    return result["tables"][name]["rows"]


class ToolsTestCase(WarehouseTestCase):
    def setUp(self):
        super().setUp()
        ozon_report(self.root / "reports" / "Август 2026.xlsx", [
            ["delivered", "B-1", "Браслет", 7012, 1, 900, -400, -100, 400, -150, 250, "2026-08-10", "FBS"],
            ["delivered", "B-2", "Доска", 1101, 1, 1300, -500, -100, 700, -250, 450, "2026-08-11", "FBO"],
            ["delivered", "B-3", "Набор", "1101, 1105", 2, 2078, -900, -200, 978, -500, 478, "2026-08-12", "FBS"],
        ], [("Общая выручка", 4278), ("Чистая прибыль", 1000), ("Рентабельность по чистой прибыли (Net Profit Margin) %", 8.0)])
        self.clock = [0.0]
        self.tools = ai_tools.AnalyticsTools(self.root, today=date(2026, 10, 24), clock=lambda: self.clock[0])
        self.addCleanup(self.tools.close)

    def call(self, name, **args):
        return self.tools.call(name, args)


class HelperTests(unittest.TestCase):
    def test_periods(self):
        for raw in ("2026-09", "2026-9", "Сентябрь 2026", "сентябрь 2026"):
            self.assertEqual(ai_tools.normalize_period(raw), "2026-09", raw)
        self.assertIsNone(ai_tools.normalize_period("2026-13"))
        self.assertEqual(ai_tools.shift_period("2026-01", -1), "2025-12")
        self.assertEqual(ai_tools.shift_period("2026-09", -12), "2025-09")

    def test_change_and_render(self):
        self.assertEqual(ai_tools._change(70.0, 100.0), "-30.0%")
        self.assertEqual(ai_tools._change(8.95, 8.48, percent_metric=True), "+0.47 п.п.")
        self.assertIsNone(ai_tools._change(5, 0))
        text = ai_tools.render({"title": "T", "tables": {"": ai_tools.table(["a", "b"], [[1.0, "x|y"]])},
                                "facts": {"f": 2.345}, "notes": ["n"]})
        self.assertEqual(text, "## T\n| a | b |\n| 1 | x/y |\n- f: 2.35\nПримечание: n")
        self.assertEqual(ai_tools.render({"error": "e", "hint": "h"}), "ОШИБКА: e\nПодсказка: h")

    def test_schemas_match_methods(self):
        for name in ai_tools.TOOL_NAMES:
            self.assertTrue(callable(getattr(ai_tools.AnalyticsTools, name, None)), name)
        self.assertEqual(ai_tools.LIVE_TOOLS, {"live_prices", "live_stocks", "live_promotions", "live_discount_requests"})


class MetricsTests(ToolsTestCase):
    def test_sales_metrics_total_matches_report_revenue(self):
        result = self.call("sales_metrics", period_from="Сентябрь 2026", period_to="2026-09", metrics=["orders", "revenue", "profit", "cancelled"])
        self.assertEqual(rows(result), [[2, 960, 244.79, 1]])

    def test_group_by_period_is_chronological(self):
        result = self.call("sales_metrics", period_from="2026-08", period_to="2026-09", group_by=["period", "scheme"], metrics=["revenue"])
        self.assertEqual(rows(result), [["2026-08", "FBO", 1300], ["2026-08", "FBS", 2978], ["2026-09", "FBO", 0], ["2026-09", "FBS", 960]])

    def test_group_by_artikul_ranks_by_profit_with_limit_note(self):
        result = self.call("sales_metrics", period_from="2026-08", period_to="2026-08", group_by=["artikul"], metrics=["profit"], limit=2)
        self.assertEqual([r[0] for r in rows(result)], ["1101, 1105", "1101"])
        self.assertTrue(any("Показаны 2 из 3" in n for n in result["notes"]))

    def test_default_period_is_latest_and_incomplete_month_is_flagged(self):
        result = self.call("sales_metrics", marketplace="wb", metrics=["revenue", "profit"])
        self.assertEqual(rows(result), [[1015, 804.25]])
        self.assertIn("Сентябрь 2026", result["title"])

    def test_validation_errors(self):
        self.assertIn("Нельзя группировать", self.call("sales_metrics", group_by=["scheme"], marketplace="wb")["error"])
        self.assertIn("Неизвестные показатели", self.call("sales_metrics", metrics=["gmv"])["error"])
        self.assertIn("Не понял период", self.call("sales_metrics", period_from="вчера")["error"])
        self.assertIn("не найден", self.call("nope")["error"])

    def test_month_metrics_pivot(self):
        result = self.call("month_metrics", period_from="2026-08", metrics=["Общая выручка", "Чистая прибыль", "Нет такого"])
        self.assertEqual(result["tables"][""]["columns"], ["показатель", "2026-08", "2026-09"])
        self.assertEqual(rows(result), [["Общая выручка", 4278, 103989], ["Чистая прибыль", 1000, 9305.14]])
        self.assertTrue(any("Нет показателей: Нет такого" in n for n in result["notes"]))

    def test_store_overview(self):
        result = self.call("store_overview", period="2026-09")
        summary = {r[0]: r for r in rows(result, "Итоги месяца (из отчётов)")}
        self.assertEqual(summary["Общая выручка"][1:], [4278, 103989, "+2330.8%"])
        self.assertEqual(summary["Рентабельность по чистой прибыли (Net Profit Margin) %"][-1], None)
        self.assertEqual(rows(result, "Лидеры по прибыли")[0][0], "7012")
        dropped = {r[0]: r[-1] for r in rows(result, "Сильнее всего упала прибыль")}
        self.assertEqual(dropped["1101, 1105"], -478)
        self.assertEqual(result["facts"]["Выручка Wildberries за месяц"], 22262)


class ProductTests(ToolsTestCase):
    def test_profile_collects_sources(self):
        result = self.call("product_profile", artikul="1101.0")
        self.assertEqual(result["title"], "Товар 1101")
        self.assertEqual(result["facts"]["Название"], "Доска")
        self.assertEqual(result["facts"]["Себестоимость, ₽"], 250)
        months = rows(result, "Продажи Ozon по месяцам (12 мес.)")
        self.assertEqual([m[0] for m in months], ["2026-08", "2026-09"])
        self.assertEqual(rows(result, "Акции"), [["Бустинг 11.11", 1375, 44.6]])
        self.assertNotIn("Продажи Wildberries по месяцам", result["tables"])

    def test_profile_unknown_artikul(self):
        self.assertIn("не найден", self.call("product_profile", artikul="9999")["error"])

    def test_find_products_splits_multi_item_rows(self):
        found = {r[0]: r for r in rows(self.call("find_products", query="11"))}
        self.assertIn("1105", found)
        self.assertEqual(found["1105"][1], "")
        self.assertEqual(found["1101"][1], "Доска")
        self.assertEqual(found["1101"][2], "Ozon, costs.xlsx")
        self.assertEqual([r[0] for r in rows(self.call("find_products", query="икона"))], ["5004"])


class ProductTableTests(ToolsTestCase):
    def setUp(self):
        super().setUp()
        snapshots = self.root / "analytics history" / "snapshots"
        snapshots.mkdir(parents=True)
        (snapshots / "2026-10-20.json").write_text('{"1101": {"price": 1375, "stock": 4}}', encoding="utf-8")
        (snapshots / "2026-10-24.json").write_text('{"1101": {"price": 1200, "stock": 0, "index": "RED"}, "7012": {"stock": 3}}', encoding="utf-8")

    def test_one_row_per_artikul_with_latest_snapshot(self):
        result = self.call("product_table", months=2)
        table = {r[0]: r for r in rows(result)}
        self.assertEqual(set(table), {"1101", "7012"})  # отправление «1101, 1105» отдельной строкой не идёт
        self.assertEqual(table["1101"][3], 0)
        self.assertEqual(table["1101"][7], "RED")
        self.assertEqual(table["1101"][4:7], [1375, 1253, 250])
        self.assertEqual(table["7012"][9:12], [2, 494.79, 0])
        self.assertIn("Август 2026 … Сентябрь 2026 (прошлые: Июнь 2026 … Июль 2026)", result["notes"][0])

    def test_filters(self):
        self.assertEqual([r[0] for r in rows(self.call("product_table", out_of_stock=True))], ["1101"])
        self.assertEqual(rows(self.call("product_table", abc=["A"])), [])
        self.assertEqual([r[0] for r in rows(self.call("product_table", artikuls=["7012"]))], ["7012"])


class CalculateTests(ToolsTestCase):
    def test_arithmetic_and_rejections(self):
        result = rows(self.call("calculate", expressions=["103989 - 22262", "1 234,5 * 2", "round(2/3, 2)", "max(1, 5)",
                                                          "__import__('os')", "1/0", "2**1000"]))
        self.assertEqual([r[1] for r in result[:4]], [81727, 2469, 0.67, 5])
        self.assertTrue(all(str(r[1]).startswith("ошибка") for r in result[4:]))
        self.assertIn("Нужно выражение", self.call("calculate")["error"])


class OverviewPeriodTests(ToolsTestCase):
    def test_current_month_is_not_the_default_overview(self):
        ozon_report(self.root / "reports" / "Октябрь 2026.xlsx", [
            ["delivered", "C-1", "Доска", 1101, 1, 1000, -400, -100, 500, -250, 250, "2026-10-05", "FBS"],
        ], [("Общая выручка", 1000), ("Чистая прибыль", 200), ("Общее количество заказов", 1)])
        result = self.call("store_overview")
        self.assertIn("Сентябрь 2026", result["title"])
        self.assertIn("выручка 1000, чистая прибыль 200, заказов 1", result["facts"]["Октябрь 2026 (идёт, по 2026-10-24)"])
        self.assertIn("Октябрь 2026", self.call("store_overview", period="2026-10")["title"])


class SupplyTests(ToolsTestCase):
    def test_supply_needs_does_not_sum_artikul_totals_per_zone(self):
        plan = self.root / "stocks reports" / "Поставка FBO 01.10.2026.json"
        plan.parent.mkdir(parents=True)
        plan.write_text(ai_tools.warehouse.json.dumps({
            "arrival": "2026-10-15",
            "supply": [{"article": "1308", "by_zone": {"Москва": 32, "СПб": 9}, "total": 41, "per_day": 1.05, "cost": 7175}],
            "fbs": {"rows": [{"article": "1308", "need": 24, "stock": 2, "shortage": 22, "cost": 3850},
                             {"article": "1101", "need": 3, "stock": 5, "shortage": 0, "cost": 0}]},
        }), encoding="utf-8")
        result = self.call("supply_needs")
        self.assertEqual(rows(result), [["1308", 41, "Москва: 32; СПб: 9", 7175, 24, 2, 22, 3850, 11025, 1.05]])
        self.assertEqual(result["facts"]["Всего денег на пополнение (по себестоимости), ₽"], 11025)
        self.assertTrue(any("23 дн. назад" in n for n in result["notes"]))

    def test_supply_needs_without_plan(self):
        self.assertIn("Плана поставки нет", self.call("supply_needs")["error"])


class SqlTests(ToolsTestCase):
    def test_select_and_limit(self):
        result = self.call("run_sql", sql="SELECT order_number FROM ozon_orders ORDER BY order_number;", limit=2)
        self.assertEqual(rows(result), [["A-1"], ["A-2"]])
        self.assertTrue(any("первые 2" in n for n in result["notes"]))

    def test_writes_are_rejected(self):
        self.assertIn("только запросы SELECT", self.call("run_sql", sql="DELETE FROM ozon_orders")["error"])
        self.assertIn("SQL:", self.call("run_sql", sql="WITH x AS (SELECT 1) DELETE FROM ozon_orders")["error"])
        self.assertIn("SQL:", self.call("run_sql", sql="SELECT 1; DELETE FROM ozon_orders")["error"])
        self.assertEqual(rows(self.call("run_sql", sql="SELECT COUNT(*) FROM ozon_orders")), [[5]])

    def test_errors_carry_hint(self):
        result = self.call("run_sql", sql="SELECT * FROM nope")
        self.assertIn("no such table", result["error"])
        self.assertIn("describe_data", result["hint"])

    def test_describe_data(self):
        catalog = {r[0]: r for r in rows(self.call("describe_data"))}
        self.assertEqual(catalog["ozon_orders"][1:3], [5, "2026-08 … 2026-09"])
        columns = rows(self.call("describe_data", table="products"), "Колонки")
        self.assertIn("min_price", [c[0] for c in columns])
        self.assertIn("нет", self.call("describe_data", table="nope")["error"])

    def test_refresh_is_throttled(self):
        self.call("describe_data")
        with patch.object(ai_tools.warehouse, "refresh") as refresh:
            self.call("describe_data")
            refresh.assert_not_called()
            self.clock[0] += ai_tools.REFRESH_INTERVAL_SECONDS
            self.call("describe_data")
            refresh.assert_called_once()


class LiveTests(ToolsTestCase):
    def test_live_prices_cached_and_formatted(self):
        organic = MagicMock()
        organic.load_prices.return_value = {"1101": {"price": 1375.0, "index": {"color": "RED", "value": 1.09, "min_price": 845.0}, "actions": ["A"]}}
        organic.price_for_good_index.return_value = 1262
        with patch.object(ai_tools, "_require_ozon_keys"), patch.object(ai_tools, "_import", return_value=organic):
            first = self.call("live_prices", artikuls=["1101", "1308"])
            second = self.call("live_prices", artikuls=["1101", "1308"])
            self.clock[0] += ai_tools.LIVE_CACHE_SECONDS
            self.call("live_prices", artikuls=["1101", "1308"])
        self.assertIs(first, second)
        self.assertEqual(organic.load_prices.call_count, 2)
        self.assertEqual(rows(first), [["1101", 1375, 1253, "RED", 1.09, 845, 1262, 1, "можно снизить до этой цены"],
                                       ["1308", None, None, None, None, None, None, None, "нет в Ozon"]])
        organic.price_for_good_index.return_value = 1100
        with patch.object(ai_tools, "_require_ozon_keys"), patch.object(ai_tools, "_import", return_value=organic):
            verdict = rows(self.call("live_prices", artikuls=["1101"]))[0][-1]
        self.assertEqual(verdict, "снижать нельзя: нужная цена ниже вашей минимальной")

    def test_live_stocks_default_to_costs_artikuls(self):
        supply = MagicMock()
        supply.ozon_post.return_value = {"items": [{"offer_id": "1101", "stocks": [
            {"type": "fbo", "present": 5, "reserved": 2}, {"type": "fbs", "present": 1, "reserved": 3}]}]}
        with patch.object(ai_tools, "_require_ozon_keys"), patch.object(ai_tools, "_import", return_value=supply):
            result = self.call("live_stocks")
        self.assertEqual(supply.ozon_post.call_args.args[1]["filter"]["offer_id"], ["1101"])
        self.assertEqual(rows(result), [["1101", 3, 0, 3]])

    def test_discount_requests_map_sku_and_drop_personal_fields(self):
        prices = MagicMock()
        prices.get_discount_requests.return_value = [{"created_at": "2026-10-23T17:53:03Z", "status": "NEW", "sku": 18, "name": "Доска",
                                                      "original_price": 856, "requested_price": 814, "email": "x@y", "last_name": "Иванов"}]
        with patch.object(ai_tools, "_require_ozon_keys"), patch.object(ai_tools, "_import", return_value=prices):
            text = ai_tools.render(self.call("live_discount_requests"))
        prices.get_discount_requests.assert_called_once_with(status="NEW")
        self.assertIn("2026-10-23 17:53 | NEW |", text)
        self.assertNotIn("Иванов", text)
        self.assertNotIn("x@y", text)

    def test_missing_keys(self):
        prices = MagicMock(OZON_CLIENT_ID="", OZON_API_KEY="")
        with patch.object(ai_tools, "_import", return_value=prices):
            result = self.call("live_promotions")
        self.assertIn("Ключи Ozon не настроены", result["error"])


if __name__ == "__main__":
    unittest.main()
