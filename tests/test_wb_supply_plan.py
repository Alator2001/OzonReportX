"""Offline checks for the WB stock planner (own warehouse + WB warehouses)."""
import sys
import unittest
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT))
import wb_supply_plan as wb  # noqa: E402

TODAY = date(2026, 10, 9)
SETTINGS = {"lead_days": 14, "cover_days": 30, "service_level": 0.9, "wb_fbw_min_monthly": 3}


def history(rows):
    frame = pd.DataFrame(rows, columns=["date", "article", "qty"])
    frame["date"] = pd.to_datetime(frame["date"])
    return frame


def orders(rows):
    return pd.DataFrame(rows, columns=["date", "article", "qty", "cancelled", "own"]).assign(date=lambda f: pd.to_datetime(f["date"]))


def daily(start, days, article="A", qty=1):
    return [(start + timedelta(days=i), article, qty) for i in range(days)]


class SeasonTests(unittest.TestCase):
    def test_window_twice_as_busy_last_year(self):
        shift = timedelta(days=364)
        base_from = TODAY - timedelta(days=90) - shift
        rows = daily(base_from, 90 + 14 + 30, qty=1)                     # обычный темп год назад, 1 шт. в день
        window_from = TODAY + timedelta(days=14) - shift
        rows += daily(window_from, 30, qty=1)                            # +1 в день в окне — итого вдвое больше
        result = wb.season_factor(history(rows), TODAY, TODAY + timedelta(days=14), TODAY + timedelta(days=44))
        self.assertEqual(result["method"], "season")
        self.assertAlmostEqual(result["raw"], 2.0, delta=0.15)
        # Истории чуть больше года — коэффициент ослаблен вдвое: 1 + (2 − 1) / 2.
        self.assertEqual(result["weight"], 0.5)
        self.assertAlmostEqual(result["factor"], 1.5, delta=0.08)

    def test_full_season_after_two_years_of_history(self):
        shift = timedelta(days=364)
        rows = daily(TODAY - timedelta(days=760), 2)                      # история с двух лет назад
        base_from = TODAY - timedelta(days=90) - shift
        rows += daily(base_from, 90 + 14 + 30, qty=1)
        rows += daily(TODAY + timedelta(days=14) - shift, 30, qty=1)
        result = wb.season_factor(history(rows), TODAY, TODAY + timedelta(days=14), TODAY + timedelta(days=44))
        self.assertEqual(result["weight"], 1.0)
        self.assertAlmostEqual(result["factor"], 2.0, delta=0.15)

    def test_without_last_year_history_factor_is_one(self):
        rows = daily(TODAY - timedelta(days=60), 60)
        result = wb.season_factor(history(rows), TODAY, TODAY + timedelta(days=14), TODAY + timedelta(days=44))
        self.assertEqual((result["method"], result["factor"]), ("none", 1.0))


class RateTests(unittest.TestCase):
    def test_rate_uses_days_wb_returned_not_90(self):
        # WB отдал заказы только за 40 дней (как метод статистики на деле): 1 заказ в день, а не 40 / 90.
        recent = orders([(TODAY - timedelta(days=i), "A", 1, False, False) for i in range(40)])
        rates = wb.order_rate(recent, TODAY)
        self.assertEqual(rates["days"], 40)
        self.assertAlmostEqual(rates["rate"], 1.0)

    def test_rate_is_average_of_last_30_days_and_whole_period(self):
        # 90 дней: первые 60 — по 1 заказу в день, последние 30 — по 3 (рост продаж).
        rows = [(TODAY - timedelta(days=i), "A", 3 if i < 30 else 1, False, False) for i in range(90)]
        rates = wb.order_rate(orders(rows), TODAY)
        self.assertAlmostEqual(rates["rate_recent"], 3.0)
        self.assertAlmostEqual(rates["rate_all"], 150 / 90)
        self.assertAlmostEqual(rates["rate"], (3.0 + 150 / 90) / 2)

    def test_cancelled_orders_count_for_period_but_not_for_rate(self):
        rows = [(TODAY - timedelta(days=9), "A", 1, True, False), (TODAY, "A", 1, False, False)]
        rates = wb.order_rate(orders(rows), TODAY)
        self.assertEqual(rates["days"], 10)
        self.assertAlmostEqual(rates["rate"], 0.1)


class PlanTests(unittest.TestCase):
    def plan(self, own_stock=None, wb_stock=None):
        recent = [(TODAY - timedelta(days=i), "A", 1, False, True) for i in range(1, 46)]     # A: 45 заказов за 90 дн.
        recent += [(TODAY - timedelta(days=i * 30), "B", 1, False, True) for i in range(1, 3)]  # B: 2 заказа
        recent += [(TODAY - timedelta(days=3), "B", 1, True, True)]                           # отмена не считается
        return wb.build_wb_plan(orders(recent), history([]), own_stock or {}, wb_stock or {}, SETTINGS, TODAY)

    def test_steady_article_goes_to_wb_rare_stays_on_own_warehouse(self):
        result = self.plan(own_stock={"A": 1, "B": 0}, wb_stock={"A": {"quantity": 2, "from_client": 1, "to_client": 0}})
        rows = result["details"].set_index("Артикул")
        self.assertEqual(result["recent_units"], 47)
        self.assertEqual(rows.loc["A", "Возить на WB"], "да")
        self.assertEqual(rows.loc["B", "Возить на WB"], "нет")
        self.assertEqual(rows.loc["B", "К поставке на WB"], 0)
        # A: спрос до прихода сверх остатка WB (2 + 1 возврат) уходит на свой склад.
        expected_spill = max(0.0, rows.loc["A", "Спрос до прихода, шт."] - 3)
        self.assertAlmostEqual(rows.loc["A", "Спрос на своём складе, шт."], round(expected_spill, 1), delta=0.11)
        self.assertGreater(rows.loc["A", "К поставке на WB"], 0)
        self.assertEqual(rows.loc["B", "Не хватает FBS WB"], rows.loc["B", "Нужно на своём складе"])

    def test_settings_are_separate_from_ozon(self):
        import json
        import tempfile
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "fbo_supply_settings.json").write_text(json.dumps({"lead_days": 30, "cover_days": 60}), encoding="utf-8")
            self.assertEqual(wb.load_settings(root)["lead_days"], 14)                  # настройки Ozon не подхватываются
            wb.save_settings(root, {**wb.WB_DEFAULTS, "lead_days": 21})
            self.assertEqual(wb.load_settings(root)["lead_days"], 21)
            self.assertEqual(json.loads((root / "fbo_supply_settings.json").read_text(encoding="utf-8"))["lead_days"], 30)

    def test_costs_come_from_wb_costs_by_nm_id(self):
        cards = {"1101": {"nm_ids": [111], "chrt_ids": []}, "1102": {"nm_ids": [222], "chrt_ids": []}}
        self.assertEqual(wb.unit_costs_by_article({"111": 250, "222": None}, cards), {"1101": 250.0})
        result = self.plan(own_stock={"B": 0})
        costs = wb.plan_costs(result, {"B": 100.0})
        details = result["details"].set_index("Артикул")
        self.assertEqual(costs["fbs"], details.loc["B", "Не хватает FBS WB"] * 100.0)
        self.assertIn("A", costs["missing"])

    def test_sales_rows_from_finance_details(self):
        details = [
            {"sellerOperName": "Продажа", "quantity": 1, "orderDt": "2025-11-02T10:00:00Z", "vendorCode": "1101"},
            {"sellerOperName": "Логистика", "quantity": 0, "orderDt": "2025-11-02T10:00:00Z", "vendorCode": "1101"},
            {"sellerOperName": "Продажа", "quantity": 2, "orderDt": "", "saleDt": "2025-11-05T10:00:00Z", "vendorCode": "1102.0"},
        ]
        self.assertEqual(wb._sales_from_details(details), [["2025-11-02", "1101", 1], ["2025-11-05", "1102", 2]])


if __name__ == "__main__":
    unittest.main()
