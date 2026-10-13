"""Offline checks for the seasonal FBO supply planner."""
import sys
import unittest
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import fbo_supply_plan as plan  # noqa: E402

TODAY = date(2026, 10, 7)
MSK, SPB = plan.MOSCOW, plan.SPB


def orders_frame(rows):
    frame = pd.DataFrame(rows, columns=["date", "offer_id", "qty", "cluster_to"])
    frame["date"] = pd.to_datetime(frame["date"])
    return frame


def daily_orders(start, days, per_day, offer_id="A", cluster=MSK):
    return [(start + timedelta(days=i), offer_id, per_day, cluster) for i in range(days)]


def settings(**overrides):
    result = {**plan.DEFAULT_SETTINGS, "zones": {MSK: [MSK, "Ростов"], SPB: [SPB]}}
    result.update(overrides)
    return result


EMPTY_STOCK = pd.DataFrame(columns=["offer_id", "cluster", "available", "transit", "requested"])


class PoissonTests(unittest.TestCase):
    def test_quantile(self):
        self.assertEqual(plan.poisson_quantile(0, 0.9), 0)
        self.assertEqual(plan.poisson_quantile(1.0, 0.9), 2)   # P(X<=2)=0.92
        self.assertEqual(plan.poisson_quantile(10.0, 0.9), 14)  # P(X<=14)=0.917


class ForecastTests(unittest.TestCase):
    def test_season_window_times_growth(self):
        start = TODAY - timedelta(days=500)
        rows = daily_orders(start, 500, 1)
        # Пик год назад в окне поставки: +4 шт. в день.
        peak_start = plan.year_ago(TODAY + timedelta(days=14))
        rows += daily_orders(peak_start, 30, 4)
        daily = orders_frame(rows).groupby("date")["qty"].sum()
        forecast = plan.store_forecast(daily, start, TODAY, TODAY + timedelta(days=14), TODAY + timedelta(days=44))
        self.assertEqual(forecast["method"], "season")
        self.assertAlmostEqual(forecast["growth"], 1.0)
        self.assertAlmostEqual(forecast["units"], 150.0)

    def test_without_last_year_uses_recent_rate(self):
        start = TODAY - timedelta(days=60)
        daily = orders_frame(daily_orders(start, 60, 2)).groupby("date")["qty"].sum()
        forecast = plan.store_forecast(daily, start, TODAY, TODAY + timedelta(days=14), TODAY + timedelta(days=44))
        self.assertEqual(forecast["method"], "recent")
        self.assertAlmostEqual(forecast["units"], 60.0)

    def test_growth_is_clipped(self):
        start = TODAY - timedelta(days=500)
        rows = daily_orders(start, 410, 1) + daily_orders(TODAY - timedelta(days=90), 90, 20)
        daily = orders_frame(rows).groupby("date")["qty"].sum()
        forecast = plan.store_forecast(daily, start, TODAY, TODAY + timedelta(days=14), TODAY + timedelta(days=44))
        self.assertEqual(forecast["growth"], 2.0)


class PlanTests(unittest.TestCase):
    def history(self):
        start = TODAY - timedelta(days=500)
        rows = daily_orders(start, 500, 1, "A", MSK)          # A: Москва
        rows += daily_orders(start, 500, 1, "A", "Ростов")    # A: Ростов -> зона Москвы
        rows += daily_orders(start, 500, 1, "A", "Новосибирск")  # A: вне зон -> FBS
        rows += [(TODAY - timedelta(days=5), "B", 1, SPB)]     # B: почти нет продаж
        return orders_frame(rows)

    def test_zone_demand_includes_assigned_clusters_and_stock_is_subtracted(self):
        stock = pd.DataFrame([
            {"offer_id": "A", "cluster": MSK, "available": 10, "transit": 5, "requested": 0},
            {"offer_id": "A", "cluster": "Ростов", "available": 3, "transit": 0, "requested": 0},
        ])
        result = plan.build_plan(self.history(), stock, settings(), TODAY)
        row = result["details"].set_index(["Артикул", "Склад FBO"]).loc[("A", MSK)]
        # Зона Москвы = Москва + Ростов ≈ 2/3 спроса A (3 шт./день), сглаживание почти не влияет.
        self.assertAlmostEqual(row["Доля зоны, %"], 66.7, delta=0.5)
        self.assertAlmostEqual(row["Спрос за цикл, шт."], 60.0, delta=1.5)
        self.assertEqual(row["Свободно"], 13)
        self.assertAlmostEqual(row["Запас на приходе"], 0.0)  # 18 шт. закончатся за 14 дней (~28 шт.)
        self.assertEqual(row["Цель"], plan.poisson_quantile(row["Спрос за цикл, шт."], 0.9))
        self.assertEqual(row["К поставке"], row["Цель"])

    def test_small_demand_is_left_to_fbs(self):
        result = plan.build_plan(self.history(), EMPTY_STOCK, settings(), TODAY, articles=["A", "B", "C"])
        details = result["details"].set_index(["Артикул", "Склад FBO"])
        self.assertEqual(details.loc[("B", SPB), "К поставке"], 0)
        self.assertIn("покрывает FBS", details.loc[("B", SPB), "Комментарий"])
        self.assertEqual(details.loc[("C", MSK), "Комментарий"], "Нет продаж — не возим")

    def test_enough_stock_means_nothing_to_ship(self):
        stock = pd.DataFrame([{"offer_id": "A", "cluster": MSK, "available": 500, "transit": 0, "requested": 0}])
        result = plan.build_plan(self.history(), stock, settings(), TODAY)
        row = result["details"].set_index(["Артикул", "Склад FBO"]).loc[("A", MSK)]
        self.assertEqual(row["К поставке"], 0)
        self.assertEqual(row["Комментарий"], "Запаса хватит")

    def test_fbs_need_covers_own_regions_and_fbo_gap_before_arrival(self):
        result = plan.build_plan(self.history(), EMPTY_STOCK, settings(), TODAY, own_stock={"A": 5})
        row = result["fbs"].set_index("Артикул").loc["A"]
        # Новосибирск (1 шт./день) обслуживает FBS: 14 дней до прихода + 30 дней окна.
        self.assertAlmostEqual(row["Регионы FBS до прихода, шт."], 14.0, delta=1.0)
        self.assertAlmostEqual(row["Регионы FBS в окне, шт."], 30.0, delta=1.5)
        # На FBO пусто: спрос зоны Москвы за 14 дней (~28 шт.) до прихода поставки уйдёт со своего склада.
        self.assertAlmostEqual(row["FBO не хватит до прихода, шт."], 28.0, delta=1.5)
        self.assertEqual(row["Нужно на FBS"], plan.poisson_quantile(row["Спрос на FBS, шт."], 0.9))
        self.assertEqual(row["Остаток FBS"], 5)
        self.assertEqual(row["Не хватает на FBS"], row["Нужно на FBS"] - 5)

    def test_fbs_takes_small_fbo_zone_demand_and_no_gap_when_fbo_is_stocked(self):
        stock = pd.DataFrame([{"offer_id": "A", "cluster": MSK, "available": 500, "transit": 0, "requested": 0}])
        result = plan.build_plan(self.history(), stock, settings(), TODAY, own_stock={})
        fbs = result["fbs"].set_index("Артикул")
        self.assertAlmostEqual(fbs.loc["A", "FBO не хватит до прихода, шт."], 0.0, delta=0.5)
        # B в СПб не возим (спрос мал) — его спрос остаётся на своём складе.
        self.assertGreater(fbs.loc["B", "Мелкий спрос зон FBO, шт."], 0)
        self.assertEqual(fbs.loc["A", "Не хватает на FBS"], fbs.loc["A", "Нужно на FBS"])

    def test_fbs_shortage_cost(self):
        fbs = pd.DataFrame([{"Артикул": "A", "Не хватает на FBS": 3}, {"Артикул": "B", "Не хватает на FBS": 2},
                            {"Артикул": "C", "Не хватает на FBS": 0}])
        self.assertEqual(plan.fbs_costs(fbs, {"A": 100.0}), {"total": 300.0, "missing_articles": ["B"]})


class OzonRetryTests(unittest.TestCase):
    def response(self, status, body=None):
        from unittest.mock import Mock
        return Mock(status_code=status, json=Mock(return_value=body or {}), text="error")

    def test_retries_server_errors_then_succeeds(self):
        from unittest.mock import patch
        replies = [self.response(500), self.response(502), self.response(200, {"items": [1]})]
        waits = []
        with patch("requests.post", side_effect=replies) as post:
            body = plan.ozon_post("https://api-seller.ozon.ru/v1/analytics/stocks", {}, sleep=waits.append)
        self.assertEqual(body, {"items": [1]})
        self.assertEqual(post.call_count, 3)
        self.assertEqual(waits, list(plan.RETRY_DELAYS[:2]))

    def test_gives_clear_error_after_last_attempt(self):
        from unittest.mock import patch
        with patch("requests.post", return_value=self.response(500)):
            with self.assertRaises(plan.OzonUnavailable) as error:
                plan.ozon_post("https://api-seller.ozon.ru/v1/analytics/stocks", {}, sleep=lambda _s: None)
        self.assertIn("повторите расчёт", str(error.exception))

    def test_client_errors_are_not_retried(self):
        from unittest.mock import patch
        with patch("requests.post", return_value=self.response(403)) as post:
            with self.assertRaises(plan.OzonUnavailable):
                plan.ozon_post("https://api-seller.ozon.ru/v1/analytics/stocks", {}, sleep=lambda _s: None)
        self.assertEqual(post.call_count, 1)


class SupplyCostTests(unittest.TestCase):
    def test_cost_is_quantity_times_unit_cost_and_missing_costs_are_listed(self):
        details = pd.DataFrame([
            {"Артикул": "A", "Склад FBO": MSK, "К поставке": 10},
            {"Артикул": "A", "Склад FBO": SPB, "К поставке": 4},
            {"Артикул": "B", "Склад FBO": MSK, "К поставке": 3},
            {"Артикул": "C", "Склад FBO": MSK, "К поставке": 0},
        ])
        costs = plan.supply_costs(details, {"A": 250.0, "C": 999.0}, [MSK, SPB])
        self.assertEqual(costs["by_zone"], {MSK: 2500.0, SPB: 1000.0})
        self.assertEqual(costs["total"], 3500.0)
        self.assertEqual(costs["missing_articles"], ["B"])

    def test_unit_cost_parsing(self):
        self.assertEqual(plan.parse_unit_cost("250,5"), 250.5)
        self.assertEqual(plan.parse_unit_cost(300), 300.0)
        for empty in (None, float("nan"), "", "abc", 0, -5):
            self.assertIsNone(plan.parse_unit_cost(empty))


if __name__ == "__main__":
    unittest.main()
