"""WB promotions: boost tiers by participation, discount history, boost payback per article."""
import json
import sys
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))
from scripts import wb_promotions as wp  # noqa: E402

RANGING = [{"condition": "productsInPromotion", "participationRate": 1, "boost": 25},
           {"condition": "productsInPromotion", "participationRate": 50, "boost": 30},
           {"condition": "productsInPromotion", "participationRate": 80, "boost": 35}]


class TierTests(unittest.TestCase):
    def test_11_11_after_leaving_four_products(self):
        # 13 из 35 с остатком (37%): буст 25%; до 50% нужно ещё 5 товаров (18/35 = 51%).
        promo = {"inPromoActionLeftovers": 13, "notInPromoActionLeftovers": 22, "participationPercentage": 37,
                 "ranging": RANGING}
        t = wp.tiers(promo)
        self.assertEqual((t["boost"], t["next_boost"], t["need"]), (25, 30, 5))
        self.assertGreaterEqual((13 + t["need"]) / 35 * 100, 50)
        self.assertLess((13 + t["need"] - 1) / 35 * 100, 50)

    def test_top_tier_and_no_participants(self):
        top = wp.tiers({"inPromoActionLeftovers": 9, "notInPromoActionLeftovers": 1, "participationPercentage": 90,
                        "ranging": RANGING})
        self.assertEqual((top["boost"], top["next_boost"], top["need"]), (35, None, None))
        none = wp.tiers({"inPromoActionLeftovers": 0, "notInPromoActionLeftovers": 4, "participationPercentage": 0,
                         "ranging": RANGING})
        self.assertEqual((none["boost"], none["next_boost"], none["need"]), (0, 25, 1))


class HistoryTests(unittest.TestCase):
    def test_records_only_changes(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            item = {"1": {"article": "1117", "price": 1750.0, "discount": 71, "discounted_price": 507.5}}
            wp.record_history(root, item, datetime(2026, 10, 27, 10))
            wp.record_history(root, item, datetime(2026, 10, 27, 11))
            wp.record_history(root, {"1": dict(item["1"], discount=40)}, datetime(2026, 10, 27, 12))
            events = wp.load_history(root)["items"]["1"]
        self.assertEqual(events, [["2026-10-27T10:00:00", 1750.0, 71], ["2026-10-27T12:00:00", 1750.0, 40]])

    def test_promo_discount_is_largest_in_30_days(self):
        history = {"items": {"1": [["2026-09-01T10:00:00", 1750.0, 80], ["2026-10-20T10:00:00", 1750.0, 71],
                                   ["2026-10-27T12:00:00", 1750.0, 40]]}}
        orders = {"1": [["2026-10-08", 1750.0, 69]]}
        found = wp.promo_discount("1", history, orders, date(2026, 10, 27))
        self.assertEqual((found["discount"], found["source"]), (71, "история скидок"))   # 80% — старше 30 дней
        self.assertEqual(wp.promo_discount("1", {"items": {}}, orders, date(2026, 10, 27))["source"], "заказы")
        self.assertIsNone(wp.promo_discount("2", history, orders, date(2026, 10, 27)))


def price_row(**updates):
    # 1117 после ручного повышения: 1750 −40% = 1050, себестоимость 250, логистика 92, комиссия 39%, эквайринг 3%.
    row = {"nm_id": 1, "article": "1117", "price": 1750.0, "discount": 40, "discounted_price": 1050.0, "min_price": 1036,
           "cost": 250.0, "logistics": 92.0, "commission": 0.39, "acquiring": 0.03, "profit": 267.0, "notes": []}
    row.update(updates)
    return row


class ArticleTests(unittest.TestCase):
    HISTORY = {"items": {"1": [["2026-10-27T10:00:00", 1750.0, 71], ["2026-10-27T12:00:00", 1750.0, 40]]}}

    def test_left_promo_with_loss(self):
        row = wp.article_rows([price_row()], self.HISTORY, {}, date(2026, 10, 27))[0]
        self.assertEqual((row["promo_discount"], row["promo_price"], row["status"]), (71, 507.5, "вышел из акции"))
        self.assertLess(row["profit_promo"], 0)
        self.assertEqual(row["verdict"], "выйти: в акции убыток")

    def test_payback_multiplier(self):
        row = wp.article_rows([price_row(discount=63, discounted_price=647.5)], {"items": {}}, {}, date(2026, 10, 27))[0]
        expected = wp.unit_profit(1036, 250, 92, 0.39, 0.03) / wp.unit_profit(647.5, 250, 92, 0.39, 0.03)
        self.assertAlmostEqual(row["multiplier"], expected)
        self.assertEqual(row["status"], "по истории не видно")
        self.assertEqual(row["verdict"], "выйти: буст не окупится")

    def test_no_conflict_when_promo_price_above_min(self):
        row = wp.article_rows([price_row(discount=30, discounted_price=1225, min_price=1036)], {"items": {}}, {},
                              date(2026, 10, 27))[0]
        self.assertEqual(row["verdict"], "конфликта нет: цена акции не ниже минимума")

    def test_verdict_thresholds(self):
        self.assertEqual(wp.verdict(100, 1.2), "держать в акции")
        self.assertEqual(wp.verdict(100, 2.0), "смотреть по продажам")
        self.assertEqual(wp.verdict(100, 4.0), "выйти: буст не окупится")
        self.assertEqual(wp.verdict(None, None), "нет данных")


class SyncTests(unittest.TestCase):
    def test_sync_keeps_old_order_discounts_when_statistics_is_rate_limited(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / ".env").write_text("WB_API_TOKEN=token\n", encoding="utf-8")
            with mock.patch.object(wp, "load_promotions", return_value=[]), \
                    mock.patch.object(wp, "load_order_discounts", return_value={"1": [["2026-10-08", 1750.0, 69]]}):
                wp.sync(root, "token")
            with mock.patch.object(wp, "load_promotions", return_value=[{"id": 1}]), \
                    mock.patch.object(wp, "load_order_discounts", side_effect=wp.WbPromotionsError("HTTP 429")):
                data = wp.sync(root, "token")
            self.assertEqual(data["order_discounts"]["1"][0][2], 69)
            self.assertIn("429", data["order_discounts_error"])
            self.assertEqual(wp.load(root)["promotions"], [{"id": 1}])


if __name__ == "__main__":
    unittest.main()


class DailyEffectTests(unittest.TestCase):
    HISTORY = {"items": {"1": [["2026-10-25T10:00:00", 1750.0, 71], ["2026-10-28T12:00:00", 1750.0, 40]]}}
    DAILY = {"items": {"1": {"2026-10-22": [10, 1, 1], "2026-10-23": [8, 0, 0], "2026-10-25": [12, 0, 0],
                             "2026-10-28": [3, 0, 0], "2026-10-29": [2, 0, 1], "2026-10-30": [4, 0, 0]}}}

    def test_discount_on_day(self):
        events = self.HISTORY["items"]["1"]
        self.assertEqual(wp.discount_on("2026-10-27", events, {}), {"discount": 71, "estimated": False})
        self.assertEqual(wp.discount_on("2026-10-28", events, {}), {"discount": 40, "estimated": False})
        self.assertEqual(wp.discount_on("2026-10-20", events, {"2026-10-20": 65}), {"discount": 65, "estimated": False})
        self.assertEqual(wp.discount_on("2026-10-20", events, {}), {"discount": 71, "estimated": True})
        self.assertIsNone(wp.discount_on("2026-10-20", [], {}))

    def test_views_per_day_with_and_without_promo(self):
        effect = wp.daily_effect(["1"], {"1": 71}, self.HISTORY, {}, self.DAILY)
        self.assertEqual((effect["promo"]["days"], effect["plain"]["days"]), (3, 3))
        self.assertAlmostEqual(effect["promo"]["views"], 10)        # (10 + 8 + 12) / 3
        self.assertAlmostEqual(effect["plain"]["views"], 3)         # (3 + 2 + 4) / 3
        self.assertAlmostEqual(effect["views_uplift"], 10 / 3)
        self.assertTrue(effect["enough"] and effect["estimated"])   # 22–23.10 — до начала истории

    def test_fact_verdict(self):
        effect = wp.daily_effect(["1"], {"1": 71}, self.HISTORY, {}, self.DAILY)
        self.assertEqual(wp.fact_verdict(effect, 2.0, 50), "держать: буст окупается")
        self.assertEqual(wp.fact_verdict(effect, 5.0, 50), "выйти: буст не окупается")
        self.assertEqual(wp.fact_verdict(effect, 2.0, -5), "выйти: в акции убыток")
        thin = wp.daily_effect(["1"], {"1": 71}, self.HISTORY, {}, {"items": {"1": {"2026-10-28": [3, 0, 0]}}})
        self.assertEqual(wp.fact_verdict(thin, 2.0, 50), "мало данных (0 | 1 дн.)")

    def test_series(self):
        self.assertEqual((wp.series_of("1117"), wp.series_of("7005"), wp.series_of("A")), ("11xx", "70xx", "A"))

    def test_daily_funnel_chunks_and_merge(self):
        responses = [mock.Mock(status_code=200, json=lambda: [{"product": {"nmId": 1}, "history": [
            {"date": "2026-10-27", "openCount": 5, "cartCount": 1, "orderCount": 0}]}]),
            mock.Mock(status_code=200, json=lambda: [])]
        pauses = []
        with mock.patch.object(wp.requests, "post", side_effect=responses) as post:
            fresh = wp.load_daily_funnel("token", list(range(25)), date(2026, 10, 27), pause=pauses.append)
        self.assertEqual([len(c.kwargs["json"]["nmIds"]) for c in post.call_args_list], [20, 5])
        self.assertEqual(pauses, [wp.FUNNEL_PAUSE])
        self.assertEqual(fresh, {"1": {"2026-10-27": [5, 1, 0]}})
        with tempfile.TemporaryDirectory() as tmp:
            wp.merge_daily(Path(tmp), {"1": {"2026-10-26": [1, 0, 0]}})
            wp.merge_daily(Path(tmp), fresh)
            self.assertEqual(wp.load_daily(Path(tmp))["items"]["1"], {"2026-10-26": [1, 0, 0], "2026-10-27": [5, 1, 0]})


class AdviceTests(unittest.TestCase):
    def test_advice_depends_on_status(self):
        self.assertEqual(wp.advice("выйти: буст не окупится", True), {"text": "не возвращать: буст не окупится", "tone": "ok"})
        self.assertEqual(wp.advice("выйти: в акции убыток", True)["text"], "не возвращать: в акции убыток")
        self.assertEqual(wp.advice("выйти: в акции убыток", False), {"text": "выйти: в акции убыток — поднять цену", "tone": "act"})
        self.assertEqual(wp.advice("держать в акции", True)["tone"], "consider")
        self.assertEqual(wp.advice("держать в акции", False)["tone"], "ok")
        self.assertEqual(wp.advice("мало данных (6 | 2 дн.)", True), {"text": "мало данных (6 | 2 дн.)", "tone": None})

    def test_left_promo_row_is_not_a_contradiction(self):
        row = wp.article_rows([price_row()], ArticleTests.HISTORY, {}, date(2026, 10, 27))[0]
        self.assertEqual(row["status"], "вышел из акции")
        self.assertEqual(row["advice"], {"text": "не возвращать: в акции убыток", "tone": "ok"})
