"""Offline checks for WB prices: payout formula, logistics fallback, commission tariffs, margin settings."""
import json
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT))
import wb_prices as wp  # noqa: E402

SETTINGS = {"min_margin": 0.25, "desired_margin": 0.30}


def sale(article, price=1000, quantity=1, acquiring=30, spp=25):
    return {"vendorCode": article, "sellerOperName": "Продажа", "quantity": quantity,
            "retailPriceWithDisc": str(price), "acquiringFee": str(acquiring), "spp": spp, "deliveryService": 0}


def delivery(article, rub):
    return {"vendorCode": article, "sellerOperName": "Логистика", "quantity": 0, "deliveryService": rub}


def response(body, status=200):
    result = mock.Mock(status_code=status)
    result.json.return_value = body
    return result


class FormulaTests(unittest.TestCase):
    def test_price_for_margin(self):
        # (250 + 150) / (1 − 0,40 − 0,03 − 0,25) = 400 / 0,32 = 1250
        self.assertEqual(wp.price_for_margin(250, 150, 0.40, 0.03, 0.25), 1250)

    def test_no_price_without_cost_or_when_fees_eat_everything(self):
        self.assertIsNone(wp.price_for_margin(0, 150, 0.40, 0.03, 0.25))
        self.assertIsNone(wp.price_for_margin(250, 150, 0.70, 0.05, 0.25))

    def test_profit_matches_wb_payout(self):
        # Реальная продажа 1205: 1440 × (1 − 0,38) − 40,16 = 852,64 — выплата WB; минус логистика и себестоимость.
        acquiring = 40.16 / 1440
        self.assertAlmostEqual(wp.unit_profit(1440, 400, 117, 0.38, acquiring), 852.64 - 117 - 400, places=2)


class ReportCostsTests(unittest.TestCase):
    def test_acquiring_and_spp_weighted_by_quantity(self):
        costs = wp.report_costs([sale("1101", price=1000, quantity=2, acquiring=60, spp=20),
                                 sale("1102", price=2000, quantity=1, acquiring=60, spp=50)])
        self.assertAlmostEqual(costs["acquiring"], 120 / 4000)
        self.assertAlmostEqual(costs["spp"], (20 * 2 + 50) / 3 / 100)

    def test_no_sales_gives_no_shares(self):
        costs = wp.report_costs([delivery("1101", 100)])
        self.assertIsNone(costs["acquiring"])
        self.assertIsNone(costs["spp"])

    def test_logistics_own_article(self):
        costs = wp.report_costs([sale("1101"), sale("1101"), delivery("1101", 150), delivery("1101", 50)])
        self.assertEqual(wp.logistics_per_unit("1101", costs), {"value": 100, "source": "артикул"})

    def test_logistics_series_when_article_sold_once(self):
        rows = [sale("1101"), delivery("1101", 90), sale("1102"), sale("1102"), delivery("1102", 210),
                sale("7001"), sale("7001"), delivery("7001", 20)]
        costs = wp.report_costs(rows)
        self.assertEqual(wp.logistics_per_unit("1101", costs), {"value": 100, "source": "серия"})
        self.assertEqual(wp.logistics_per_unit("1199", costs), {"value": 100, "source": "серия"})

    def test_logistics_store_when_series_is_thin(self):
        costs = wp.report_costs([sale("1101"), delivery("1101", 100), sale("7001"), sale("7001"), delivery("7001", 20)])
        self.assertEqual(wp.logistics_per_unit("5003", costs), {"value": 40, "source": "магазин"})

    def test_logistics_none_without_sales(self):
        self.assertIsNone(wp.logistics_per_unit("1101", wp.report_costs([])))


class CommissionTests(unittest.TestCase):
    REPORT = [{"subjectID": 824, "subjectName": "Разделочные доски", "kgvpMarketplace": 43, "paidStorageKgvp": 38.5},
              {"subjectID": 1, "kgvpMarketplace": None, "paidStorageKgvp": 10}]

    def test_parse_takes_fbs_and_fbw_tariffs(self):
        self.assertEqual(wp.parse_commission(self.REPORT), {824: {"fbs": 0.43, "fbw": 0.385}})

    def test_cached_for_a_day(self):
        with tempfile.TemporaryDirectory() as tmp, mock.patch.object(wp.requests, "request",
                                                                     return_value=response({"report": self.REPORT})) as request:
            first = wp.load_commission(Path(tmp), "token", date(2026, 10, 27))
            second = wp.load_commission(Path(tmp), "token", date(2026, 10, 27))
            self.assertEqual(first, second)
            self.assertEqual(request.call_count, 1)
            wp.load_commission(Path(tmp), "token", date(2026, 10, 28))
            self.assertEqual(request.call_count, 2)

    def test_http_error_raises(self):
        with tempfile.TemporaryDirectory() as tmp, mock.patch.object(wp.requests, "request", return_value=response({}, 401)):
            with self.assertRaises(wp.WbPricesError):
                wp.load_commission(Path(tmp), "token", date(2026, 10, 27))


class LoadPricesTests(unittest.TestCase):
    def test_pages_until_empty(self):
        good = {"nmID": 1, "vendorCode": "1205", "discount": 64,
                "sizes": [{"price": 4000, "discountedPrice": 1440}]}
        pages = [response({"data": {"listGoods": [good]}}), response({"data": {"listGoods": []}})]
        with mock.patch.object(wp.requests, "request", side_effect=pages) as request:
            prices = wp.load_prices("token")
        self.assertEqual(prices, {1: {"article": "1205", "price": 4000.0, "discount": 64, "discounted_price": 1440.0}})
        self.assertEqual(request.call_args_list[1].kwargs["params"]["offset"], 1)


class BuildRowsTests(unittest.TestCase):
    PRICES = {1: {"article": "1101", "price": 2500.0, "discount": 60, "discounted_price": 1000.0},
              2: {"article": "1102", "price": 2500.0, "discount": 60, "discounted_price": 1000.0},
              3: {"article": "7001", "price": 1000.0, "discount": 5, "discounted_price": 950.0}}
    SUBJECTS = {1: {"id": 824, "name": "Разделочные доски"}, 2: {"id": 824, "name": "Разделочные доски"},
                3: {"id": 955, "name": "Обереги"}}
    COMMISSION = {824: {"fbs": 0.43, "fbw": 0.385}, 955: {"fbs": 0.33, "fbw": 0.295}}

    def costs(self):
        rows = [dict(sale("1101", acquiring=30, spp=20), commissionPercent=40),
                dict(sale("1101", acquiring=30, spp=20), commissionPercent=40), delivery("1101", 300)]
        return wp.report_costs(rows)

    def test_row_numbers_use_actual_commission(self):
        row = wp.build_rows(self.PRICES, {"1": 250}, self.SUBJECTS, self.COMMISSION, self.costs(), SETTINGS)[0]
        self.assertEqual((row["commission"], row["commission_source"]), (0.40, "артикул"))
        self.assertEqual(row["logistics"], 150)
        self.assertAlmostEqual(row["profit"], 1000 * (1 - 0.40 - 0.03) - 150 - 250)
        self.assertAlmostEqual(row["margin"], row["profit"] / 1000)
        self.assertEqual(row["buyer_price"], 800)
        self.assertEqual(row["min_price"], 1250)          # 400 / (1 − 0,40 − 0,03 − 0,25)
        self.assertEqual(row["desired_price"], round(400 / (1 - 0.40 - 0.03 - 0.30)))

    def test_series_commission_for_article_without_sales(self):
        rows = [dict(sale("1101"), commissionPercent=38), dict(sale("1102"), commissionPercent=40),
                dict(sale("1102"), commissionPercent=40)]
        rate = wp.commission_share("1199", wp.report_costs(rows), {"fbs": 0.43})
        self.assertAlmostEqual(rate["value"], (0.38 + 0.80) / 3)
        self.assertEqual(rate["source"], "серия")

    def test_tariff_when_store_has_no_sales(self):
        self.assertEqual(wp.commission_share("7001", wp.report_costs([]), {"fbs": 0.33, "fbw": 0.295}),
                         {"value": 0.33, "source": "тариф WB"})
        self.assertIsNone(wp.commission_share("7001", wp.report_costs([]), None))

    def test_tariff_before_store_average_of_other_category(self):
        costs = wp.report_costs([dict(sale("1101"), commissionPercent=40), dict(sale("1101"), commissionPercent=40)])
        self.assertEqual(wp.commission_share("5003", costs, {"fbs": 0.33, "fbw": 0.295})["source"], "тариф WB")
        self.assertEqual(wp.commission_share("5003", costs, None), {"value": 0.40, "source": "магазин"})

    def test_missing_cost_leaves_numbers_empty(self):
        row = wp.build_rows(self.PRICES, {"1": 250}, self.SUBJECTS, self.COMMISSION, self.costs(), SETTINGS)[1]
        self.assertEqual(row["notes"], ["нет себестоимости в wb_costs.xlsx"])
        self.assertIsNone(row["min_price"])
        self.assertIsNone(row["profit"])


class DiscountRangeTests(unittest.TestCase):
    def test_discount_keeps_price_not_below_target(self):
        # 1117: цена до скидки 1750, минимальная 1037, желательная 1222 → 30–40 %.
        self.assertEqual(wp.discount_for(1750, 1037), 40)       # 1750 × 0,60 = 1050 ≥ 1037
        self.assertEqual(wp.discount_for(1750, 1222), 30)       # 1750 × 0,70 = 1225 ≥ 1222
        self.assertEqual(wp.discount_for(1000, 750), 25)        # ровно 750 — целая скидка без запаса
        self.assertEqual(wp.discount_for(1000, 1000), 0)
        self.assertIsNone(wp.discount_for(1000, 1200))          # поднимать надо цену до скидки
        self.assertIsNone(wp.discount_for(1000, None))

    def test_row_has_range_and_view_colors_current_discount(self):
        from wb_prices_view import DARK_GREEN, GREEN, RED, WbPricesView
        prices = {1: {"article": "1117", "price": 1750.0, "discount": 71, "discounted_price": 507.5}}
        costs = wp.report_costs([dict(sale("1117", price=578), commissionPercent=39),
                                 dict(sale("1117", price=578), commissionPercent=39), delivery("1117", 184)])
        row = wp.build_rows(prices, {"1": 250}, {}, {}, costs, SETTINGS)[0]
        self.assertLess(row["discount_desired"], row["discount_max"])
        self.assertGreaterEqual(1750 * (100 - row["discount_max"]) / 100, row["min_price"])
        text, fill = WbPricesView.discount_range(row)
        self.assertEqual((text, fill), (f"{row['discount_desired']}–{row['discount_max']} %", RED))
        self.assertEqual(WbPricesView.discount_range(dict(row, discount=row["discount_max"]))[1], GREEN)
        self.assertEqual(WbPricesView.discount_range(dict(row, discount=row["discount_desired"]))[1], DARK_GREEN)
        self.assertEqual(WbPricesView.discount_range(dict(row, discount_max=None, discount_desired=None))[0],
                         "поднять цену до скидки")


class SnapshotTests(unittest.TestCase):
    def test_sync_saves_snapshot_and_compute_reads_it(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / ".env").write_text("WB_API_TOKEN=token\n", encoding="utf-8")
            prices = {1: {"article": "1101", "price": 2500.0, "discount": 60, "discounted_price": 1000.0}}
            with mock.patch.object(wp, "load_prices", return_value=prices), \
                    mock.patch.object(wp, "load_subjects", return_value={1: {"id": 824, "name": "Доски"}}), \
                    mock.patch.object(wp, "load_commission", return_value={}):
                wp.sync(root, "token")
            snapshot = wp.load_snapshot(root)
            self.assertEqual(snapshot["prices"]["1"]["article"], "1101")
            result = wp.compute(root, snapshot)
            self.assertEqual(result["rows"][0]["article"], "1101")
            self.assertIn("нет себестоимости в wb_costs.xlsx", result["rows"][0]["notes"])
            (root / ".env").write_text("WB_API_TOKEN=other\n", encoding="utf-8")
            self.assertIsNone(wp.load_snapshot(root))


class SettingsTests(unittest.TestCase):
    def test_defaults_when_missing(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(wp.load_settings(Path(tmp)), SETTINGS)

    def test_save_and_load_own_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            wp.save_settings(Path(tmp), "0.2", 0.35)
            self.assertEqual(json.loads((Path(tmp) / "wb_price_settings.json").read_text()),
                             {"min_margin": 0.2, "desired_margin": 0.35})
            self.assertEqual(wp.load_settings(Path(tmp)), {"min_margin": 0.2, "desired_margin": 0.35})
            self.assertFalse((Path(tmp) / "margin_settings.json").exists())

    def test_rejects_bad_margins(self):
        for low, high in ((0.3, 0.3), (0, 0.3), (0.2, 1), ("x", 0.3)):
            with self.assertRaises(wp.WbPricesError):
                wp.validate_margins(low, high)

    def test_broken_file_falls_back_to_defaults(self):
        with tempfile.TemporaryDirectory() as tmp:
            (Path(tmp) / "wb_price_settings.json").write_text('{"min_margin": 0.5, "desired_margin": 0.4}')
            with mock.patch("builtins.print"):
                self.assertEqual(wp.load_settings(Path(tmp)), SETTINGS)


if __name__ == "__main__":
    unittest.main()
