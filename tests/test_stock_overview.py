"""Offline checks for the «Остатки» section: same planners as «Поставки», 30 days, one FBS for both stores."""
import sys
import unittest
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import fbo_supply_plan as ozon_plan  # noqa: E402
import stock_overview as so  # noqa: E402

TODAY = date(2026, 10, 22)
MSK = ozon_plan.MOSCOW


def source(needs, own=None, stock=None, transit=None):
    return {"needs": needs, "own": own or {}, "stock": stock or {}, "transit": transit or {}}


class HorizonTests(unittest.TestCase):
    def test_supply_planner_runs_for_the_next_30_days_without_lead_time(self):
        orders = pd.DataFrame([(TODAY - timedelta(days=i), "A", 2, MSK) for i in range(1, 61)],
                              columns=["date", "offer_id", "qty", "cluster_to"])
        orders["date"] = pd.to_datetime(orders["date"])
        stock = pd.DataFrame(columns=["offer_id", "cluster", "available", "transit", "requested"])
        settings = so.horizon_settings({**ozon_plan.DEFAULT_SETTINGS, "zones": {MSK: [MSK]}})
        plan = ozon_plan.build_plan(orders, stock, settings, TODAY, own_stock={})
        self.assertEqual(plan["window"], (TODAY, TODAY + timedelta(days=30)))
        # Истории за год нет — темп 30 дней: 2 шт. в день × 30 = 60 шт., все в Москве → цель FBO = квантиль Пуассона(60).
        self.assertEqual(so.ozon_needs(plan)["A"]["need_fbo"], ozon_plan.poisson_quantile(60.0, 0.9))


class NeedsTests(unittest.TestCase):
    def test_ozon_need_is_the_sum_of_zone_targets(self):
        plan = {
            "details": pd.DataFrame({"Артикул": ["A", "A", "B"], "Цель": [5, 3, 0]}),
            "fbs": pd.DataFrame({"Артикул": ["A", "B"], "Спрос на FBS, шт.": [1.5, 4.0]}),
        }
        self.assertEqual(so.ozon_needs(plan), {"A": {"need_fbo": 8, "fbs_demand": 1.5}, "B": {"need_fbo": 0, "fbs_demand": 4.0}})

    def test_wb_need(self):
        plan = {"details": pd.DataFrame({"Артикул": ["W"], "Цель FBW": [6], "Спрос на своём складе, шт.": [2.0]})}
        self.assertEqual(so.wb_needs(plan), {"W": {"need_fbw": 6, "fbs_demand": 2.0}})


class RowsTests(unittest.TestCase):
    def test_one_fbs_for_both_stores(self):
        ozon = source({"A": {"need_fbo": 14, "fbs_demand": 6.0}}, own={"A": 3}, stock={"A": 5}, transit={"A": 2})
        wb = source({"A": {"need_fbw": 4, "fbs_demand": 4.0}}, own={"A": 1}, stock={"A": 4})
        row = so.build_rows(ozon, wb, 0.9)[0]
        # FBS — один склад: берём наибольший остаток, спрос Ozon и WB складываем → квантиль Пуассона(10) = 14.
        self.assertEqual((row["fbs"], row["fbs_by_store"]), (3, {"Ozon": 3, "WB": 1}))
        self.assertEqual((row["need_fbs"], row["short_fbs"]), (14, 11))
        # FBO: нужно 14, есть 5 + 2 в пути → не хватает 7. FBW: нужно 4, есть 4.
        self.assertEqual((row["need_fbo"], row["short_fbo"]), (14, 7))
        self.assertEqual((row["need_fbw"], row["short_fbw"]), (4, 0))

    def test_store_that_did_not_load_leaves_empty_columns(self):
        rows = so.build_rows(source({"A": {"need_fbo": 2, "fbs_demand": 0.0}}, own={"A": 1}), None, 0.9)
        self.assertEqual((rows[0]["fbw"], rows[0]["need_fbw"], rows[0]["need_fbo"]), (None, None, 2))

    def test_articles_without_stock_or_need_are_hidden(self):
        rows = so.build_rows(source({"A": {"need_fbo": 0, "fbs_demand": 0.0}}, own={"A": 0}), None, 0.9)
        self.assertEqual(rows, [])


class ForecastNoteTests(unittest.TestCase):
    def test_notes_explain_the_gap_with_last_month(self):
        self.assertEqual(so.forecast_note("Ozon", 304.87, 91.0, "season", 380.0, 0.8),
                         "Ozon: прогноз на 30 дней — 305 шт. (за последние 30 дней 91), год назад в эти дни 380 шт. × темп 0,80")
        self.assertEqual(so.forecast_note("WB", 40.2, 43.0, "season", factor=2.16),
                         "WB: прогноз на 30 дней — 40 шт. (за последние 30 дней 43), сезонность ×2,16")


if __name__ == "__main__":
    unittest.main()
