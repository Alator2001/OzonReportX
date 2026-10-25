"""Offline checks for the data behind the «Поиск» screen: windows, KPIs, store diagnosis, gaps, actions, positions table."""
import shutil
import sys
import unittest
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import search_view as view  # noqa: E402

# search_view импортирует соседей как scripts.x или как x — берём те же объекты, что видит он.
db, sem = view.db, view.semantics

TMP = ROOT / ".cache" / "search-view-tests"
LAST = date(2026, 10, 21)


def day(offset: int) -> str:
    return (LAST - timedelta(days=offset)).isoformat()


class ViewTestCase(unittest.TestCase):
    def setUp(self):
        shutil.rmtree(TMP, ignore_errors=True)
        TMP.mkdir(parents=True)
        self.path = TMP / "search.db"
        self.conn = db.connect(self.path)
        self.addCleanup(shutil.rmtree, TMP, True)
        self.addCleanup(self.conn.close)
        self.w = view.windows(LAST)
        self.settings = sem.load_settings(TMP / "none.json")

    def stats(self, offset, sku, query, searched):
        with self.conn:
            db.upsert_query_stats(self.conn, day(offset), [{"sku": sku, "query": query, "searched_users": searched}])
            current = self.conn.execute("SELECT searched_users FROM sku_daily WHERE date = ? AND sku = ?", (day(offset), sku)).fetchone()
            db.upsert_sku_daily(self.conn, [{"date": day(offset), "sku": sku, "searched_users": (current[0] or 0 if current else 0) + searched}])

    def sales(self, offset, sku, units, revenue=0.0):
        with self.conn:
            db.upsert_sku_daily(self.conn, [{"date": day(offset), "sku": sku, "ordered_units": units, "revenue": revenue}])

    def card(self, sku, name, description=""):
        with self.conn:
            db.upsert_products(self.conn, [{"sku": sku, "offer_id": f"A{sku}", "name": name}])
            db.upsert_card_snapshots(self.conn, "2026-10-23", [{"sku": sku, "name": name, "description": description,
                                                                "attributes": {}, "images": ["1.jpg"]}])


class WindowAndKpiTests(ViewTestCase):
    def test_windows_are_seven_days_back_to_back(self):
        self.assertEqual(self.w["7"], ("2026-10-15", "2026-10-21"))
        self.assertEqual(self.w["prev7"], ("2026-10-08", "2026-10-14"))
        self.assertEqual(self.w["28"], ("2026-09-24", "2026-10-21"))

    def test_store_demand_takes_max_over_skus_per_query(self):
        self.stats(0, 1, "доска", 100)
        self.stats(0, 2, "доска", 80)
        self.stats(0, 1, "доска из кедра", 10)
        self.stats(8, 1, "доска", 50)
        k = view.kpis(self.conn, self.w)
        self.assertEqual(k["searches"]["value"], 110)
        self.assertEqual(k["searches"]["previous"], 50)
        self.assertAlmostEqual(k["searches"]["change"], 1.2)
        self.assertEqual(k["queries"]["value"], 2)

    def test_change_without_base_is_none(self):
        self.assertIsNone(view.change(10, 0))

    def test_daily_series_fills_missing_days_with_zero(self):
        self.stats(0, 1, "доска", 100)
        series = view.daily(self.conn, self.w)
        self.assertEqual(len(series), 28)
        self.assertEqual(series[-1], {"date": "2026-10-21", "searches": 100.0, "orders": 0.0})
        self.assertEqual(series[0]["searches"], 0.0)


class DiagnosisTests(unittest.TestCase):
    @staticmethod
    def k(searches, orders, previous_orders=20):
        return {"searches": {"change": searches}, "orders": {"change": orders, "previous": previous_orders}}

    def test_orders_drop_with_stable_demand_is_conversion_problem(self):
        self.assertEqual(view.store_diagnosis(self.k(-0.06, -0.57))["kind"], "conversion")

    def test_both_drop_is_demand_problem(self):
        self.assertEqual(view.store_diagnosis(self.k(-0.4, -0.5))["kind"], "demand")

    def test_small_shop_or_stable_orders_give_no_verdict(self):
        self.assertIsNone(view.store_diagnosis(self.k(-0.06, -0.57, previous_orders=3)))
        self.assertIsNone(view.store_diagnosis(self.k(-0.06, -0.1)))


class GapAndActionTests(ViewTestCase):
    def setUp(self):
        super().setUp()
        self.card(1, "Доска разделочная из кедра", "Деревянная доска для кухни, для сервировки стола")
        self.card(2, "Доска разделочная круглая", "Деревянная доска")
        for offset in range(14):
            self.stats(offset, 1, "доска разделочная", 300)
            self.stats(offset, 1, "доска разделочная деревянная", 200)
            self.stats(offset, 1, "стол", 500)       # Ozon показал доску по чужому запросу
            self.stats(offset, 2, "доска разделочная деревянная", 150)
        sem.rebuild(self.conn, self.settings)

    def test_gap_only_for_clusters_whose_words_are_mostly_in_title(self):
        rows = view.coverage(self.conn, 1, self.w)
        gaps = {g["cluster"]: g for g in view.gaps(rows)}
        self.assertIn("доска разделочная деревянная", gaps)
        self.assertEqual(gaps["доска разделочная деревянная"]["kind"], "not_in_title")
        self.assertNotIn("стол", gaps)  # «стол» есть в описании, но в названии нет ни слова — это не про товар
        self.assertNotIn("доска разделочная", gaps)  # уже в названии

    def test_coverage_sorted_by_this_products_demand(self):
        rows = view.coverage(self.conn, 1, self.w)
        self.assertEqual(rows[0]["cluster"], "стол")
        self.assertEqual(rows[0]["demand"], 500 * 14)

    def test_title_coverage_share(self):
        rows = view.coverage(self.conn, 1, self.w)
        self.assertAlmostEqual(view.title_coverage(rows), 300 / (300 + 200 + 500))

    def test_same_gap_in_many_products_is_one_action(self):
        products = view.products(self.conn, self.w)
        items = [i for i in view.actions(self.conn, self.w, products) if i["kind"] == "gap"]
        self.assertEqual(len(items), 1)
        self.assertEqual(sorted(items[0]["offers"]), ["A1", "A2"])
        self.assertIn("у 2 товара", items[0]["text"])

    def test_rising_query_must_match_product_titles(self):
        for offset in range(7):
            self.stats(offset, 1, "доска разделочная большая", 100)
            self.stats(offset + 7, 1, "доска разделочная большая", 10)
            self.stats(offset, 1, "стол обеденный", 400)
            self.stats(offset + 7, 1, "стол обеденный", 10)
        sem.rebuild(self.conn, self.settings)
        rising = [i["text"] for i in view.actions(self.conn, self.w, view.products(self.conn, self.w)) if i["kind"] == "rising"]
        self.assertTrue(any("доска разделочная большая" in t for t in rising))
        self.assertFalse(any("стол" in t for t in rising))

    def test_product_list_and_detail(self):
        products = {p["sku"]: p for p in view.products(self.conn, self.w)}
        self.assertEqual(products[1]["gaps"], 1)
        detail = view.product_detail(self.conn, 1, self.w)
        self.assertEqual(detail["queries"][0]["query"], "стол")
        self.assertEqual(detail["card"]["images"], 1)
        self.assertEqual(detail["changes"], [])

    def test_card_changes_list_changed_parts(self):
        with self.conn:
            db.upsert_card_snapshots(self.conn, "2026-10-24", [{"sku": 1, "name": "Доска разделочная деревянная из кедра",
                                                                "description": "Деревянная доска для кухни, для сервировки стола",
                                                                "attributes": {}, "images": ["1.jpg", "2.jpg"]}])
        self.assertEqual(view.card_changes(self.conn, 1), [{"date": "2026-10-24", "parts": ["название", "фото"]}])

    def test_set_target_marks_all_linked_products(self):
        keyword_id = self.conn.execute("SELECT id FROM keywords WHERE query = 'доска разделочная деревянная'").fetchone()[0]
        self.assertEqual(view.set_target(self.conn, keyword_id, True), 2)
        self.assertTrue(all(r[0] for r in self.conn.execute("SELECT is_target FROM sku_keywords WHERE keyword_id = ?", (keyword_id,))))


class StableTrendTests(ViewTestCase):
    def test_query_seen_on_few_days_has_no_trend(self):
        for offset in range(14):
            self.stats(offset, 1, "доска разделочная", 100)
        self.stats(9, 1, "подставка для досок", 300)   # прошлую неделю один день, на этой — пропуск
        sem.rebuild(self.conn, self.settings)
        stable = view.stable_keywords(self.conn, self.w)
        names = dict(self.conn.execute("SELECT query, id FROM keywords").fetchall())
        self.assertIn(names["доска разделочная"], stable)
        self.assertNotIn(names["подставка для досок"], stable)
        clusters = {c["name"]: c for c in view.clusters(self.conn, self.w)}
        self.assertIsNone(clusters["подставка для досок"]["trend"])  # а не −100%
        self.assertEqual(clusters["подставка для досок"]["demand_28"], 300)

    def test_drops_are_one_action(self):
        for sku in (1, 2):
            self.card(sku, f"Доска {sku}")
            for offset in range(7, 14):
                self.stats(offset, sku, "доска", 1000)
            for offset in range(7):
                self.stats(offset, sku, "доска", 100)
        sem.rebuild(self.conn, self.settings)
        drops = [i for i in view.actions(self.conn, self.w, view.products(self.conn, self.w)) if i["kind"] == "drop"]
        self.assertEqual(len(drops), 1)
        self.assertEqual(sorted(drops[0]["offers"]), ["A1", "A2"])
        self.assertIn("2 товара", drops[0]["text"])

    def test_searched_but_never_ordered_is_flagged(self):
        self.card(1, "Доска разделочная")
        self.card(3, "Доска разделочная большая")
        for offset in range(28):
            self.stats(offset, 1, "доска разделочная", 1000)
            self.stats(offset, 3, "доска разделочная большая", 1000)
        self.sales(0, 1, 2)
        sem.rebuild(self.conn, self.settings)
        products = {p["sku"]: p for p in view.products(self.conn, self.w)}
        self.assertTrue(products[3]["no_sales"])
        self.assertFalse(products[1]["no_sales"])
        silent = [i for i in view.actions(self.conn, self.w, list(products.values())) if i["kind"] == "no_sales"]
        self.assertEqual(silent[0]["offers"], ["A3"])
        self.assertIn("7\u00a0000 за неделю", silent[0]["text"])

    def test_plural(self):
        self.assertEqual([view.goods(n) for n in (1, 2, 5, 11, 21, 32)],
                         ["1 товар", "2 товара", "5 товаров", "11 товаров", "21 товар", "32 товара"])


class PositionsTableTests(ViewTestCase):
    def test_regions_side_by_side_with_spread_and_change(self):
        msk = db.upsert_region(self.conn, "msk", "Москва")
        spb = db.upsert_region(self.conn, "spb", "Санкт-Петербург")
        self.card(1, "Доска разделочная")
        keyword_id = db.ensure_keyword(self.conn, "доска разделочная", "api")
        with self.conn:
            for region, pos_now, pos_week_ago in [(msk, 5, 9), (spb, 40, 40)]:
                db.upsert_positions(self.conn, "2026-10-21", region, [{"sku": 1, "keyword_id": keyword_id, "position": pos_now, "found": 1}])
                db.upsert_positions(self.conn, "2026-10-14", region, [{"sku": 1, "keyword_id": keyword_id, "position": pos_week_ago, "found": 1}])
        rows = view.positions_table(self.conn)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["regions"]["Москва"]["position"], 5.0)
        self.assertEqual(rows[0]["regions"]["Москва"]["change"], 4.0)  # поднялся на 4 места
        self.assertEqual(rows[0]["spread"], 35.0)
        self.assertEqual(rows[0]["offer_id"], "A1")

    def test_empty_database_loads_as_empty(self):
        self.assertEqual(view.load(TMP / "missing.db"), {"empty": True})
        self.assertEqual(view.load(self.path), {"empty": True})


if __name__ == "__main__":
    unittest.main()
