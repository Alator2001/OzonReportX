"""Offline checks for search semantics: lemmas, clusters, coverage, scoring, manual input and CSV import."""
import json
import shutil
import sys
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import search_collect as collect  # noqa: E402

# Сборщик импортирует соседей как scripts.x или как x — в зависимости от sys.path. Патчим те же объекты, что видит он.
api, db, sem = collect.api, collect.db, collect.semantics

FIXTURES = ROOT / "tests" / "fixtures" / "search_api"
TMP = ROOT / ".cache" / "search-semantics-tests"


class DbTestCase(unittest.TestCase):
    def setUp(self):
        shutil.rmtree(TMP, ignore_errors=True)
        TMP.mkdir(parents=True)
        self.conn = db.connect(TMP / "search.db")
        self.settings = sem.load_settings(TMP / "missing.json")

    def tearDown(self):
        self.conn.close()
        shutil.rmtree(TMP, ignore_errors=True)

    def stats(self, day, sku, query, searched, position=None):
        with self.conn:
            db.upsert_query_stats(self.conn, day, [{"sku": sku, "query": query, "searched_users": searched, "position": position}])

    def card(self, sku, name, description="", attributes=None):
        with self.conn:
            db.upsert_card_snapshots(self.conn, "2026-10-23", [{"sku": sku, "name": name, "description": description,
                                                                "attributes": attributes or {}, "images": []}])

    def keyword(self, query):
        return self.conn.execute("SELECT * FROM keywords WHERE query = ?", (query,)).fetchone()


class LemmaTests(unittest.TestCase):
    def test_word_forms_and_order_give_same_key(self):
        self.assertEqual(sem.lemma_key("доски разделочные"), sem.lemma_key("Разделочная доска"))

    def test_prepositions_dropped_and_yo_normalized(self):
        self.assertEqual(sem.lemmas("Доски из кедра для подачи"), ["доска", "кедр", "подача"])
        self.assertEqual(sem.lemmas("ёлочные игрушки"), ["елочный", "игрушка"])

    def test_numbers_and_units_are_separate_tokens(self):
        self.assertEqual(sem.lemmas("доска 28см 1,5 кг"), ["доска", "28", "см", "1.5", "кг"])


class ClusterTests(unittest.TestCase):
    def test_extra_meaningful_word_opens_own_cluster(self):
        mapping = sem.build_clusters([("доска разделочный", 100), ("деревянный доска разделочный", 50)], 0.75)
        self.assertEqual(mapping["деревянный доска разделочный"], "деревянный доска разделочный")

    def test_long_tail_joins_closest_head_without_chaining(self):
        groups = [("деревянный доска разделочный", 100), ("большой деревянный доска разделочный", 10),
                  ("большой деревянный доска", 5)]
        mapping = sem.build_clusters(groups, 0.75)
        self.assertEqual(mapping["большой деревянный доска разделочный"], "деревянный доска разделочный")
        # 2/4 с главной группой — не присоединяется, даже если близка к уже присоединённому запросу.
        self.assertEqual(mapping["большой деревянный доска"], "большой деревянный доска")

    def test_result_does_not_depend_on_input_order(self):
        groups = [("a b", 10), ("a b c", 5), ("c d", 5)]
        self.assertEqual(sem.build_clusters(groups, 0.6), sem.build_clusters(list(reversed(groups)), 0.6))


class FrequencyAndClusterDbTests(DbTestCase):
    def test_frequency_is_daily_max_over_skus_not_sum(self):
        self.stats("2026-10-20", 1, "доска разделочная", 100)
        self.stats("2026-10-20", 2, "доска разделочная", 80)
        self.stats("2026-10-21", 1, "доска разделочная", 50)
        sem.rebuild(self.conn, self.settings)
        self.assertEqual(self.keyword("доска разделочная")["frequency"], 150)
        self.assertEqual(self.keyword("доска разделочная")["frequency_source"], "api")

    def test_cluster_ids_are_stable_and_locked_keywords_kept(self):
        self.stats("2026-10-20", 1, "доска разделочная", 100)
        self.stats("2026-10-20", 1, "разделочные доски", 10)
        self.stats("2026-10-20", 1, "доска из кедра", 5)
        sem.rebuild(self.conn, self.settings)
        main = self.keyword("доска разделочная")["cluster_id"]
        self.assertEqual(self.keyword("разделочные доски")["cluster_id"], main)
        sem.lock_cluster(self.conn, self.keyword("доска из кедра")["id"], main)
        sem.rebuild(self.conn, self.settings)
        self.assertEqual(self.keyword("доска разделочная")["cluster_id"], main)
        self.assertEqual(self.keyword("доска из кедра")["cluster_id"], main)
        name = self.conn.execute("SELECT name FROM clusters WHERE id = ?", (main,)).fetchone()[0]
        self.assertEqual(name, "доска разделочная")


class CoverageTests(DbTestCase):
    def test_coverage_by_field_and_missing_words(self):
        self.stats("2026-10-20", 1, "доска разделочная", 100)
        self.stats("2026-10-20", 1, "доска разделочная деревянная", 60)
        self.stats("2026-10-20", 1, "набор досок", 20)
        self.card(1, "Доска разделочная из кедра", "Разделочная доска для кухни", {"6383": ["Деревянный"]})
        sem.rebuild(self.conn, self.settings)
        matrix = {row["cluster"]: row for row in sem.coverage_matrix(self.conn, 1)}
        self.assertTrue(matrix["доска разделочная"]["in_title"])
        self.assertTrue(matrix["доска разделочная"]["in_description"])
        wooden = matrix["доска разделочная деревянная"]
        # Ни одно поле не содержит всех трёх слов, но вместе они покрывают кластер: пропущенных слов нет.
        self.assertEqual((wooden["in_title"], wooden["in_attributes"], wooden["missing"]), (False, False, []))
        self.assertEqual(wooden["shares"], {"title": 0.67, "description": 0.67, "attributes": 0.33})
        self.assertEqual(matrix["набор досок"]["missing"], ["набор"])
        self.assertEqual(list(matrix)[0], "доска разделочная")  # по частотности

    def test_card_parsing_splits_name_description_attributes(self):
        item = json.loads((FIXTURES / "product_attributes.json").read_text(encoding="utf-8"))["result"][0]
        card = api.parse_card(item)
        self.assertEqual(card["description"], "Круглая разделочная доска из массива кедра для подачи.")
        self.assertEqual(card["attributes"], {"85": ["Нет бренда"], "6383": ["Кедр"], "23171": ["#Подарок #разделочные_доски"]})
        self.assertEqual(card["rich_content"], '{"content": []}')
        self.assertEqual(card["images"][0], "https://ir.ozone.ru/s3/multimedia-1-2/main.jpg")
        self.assertEqual(len(card["images"]), 2)

    def test_snapshot_hashes_change_only_with_content(self):
        self.card(1, "Доска")
        before = self.conn.execute("SELECT name_hash, attributes_hash FROM card_snapshots").fetchone()
        self.card(1, "Доска", attributes={"6383": ["Кедр"]})
        after = self.conn.execute("SELECT name_hash, attributes_hash FROM card_snapshots").fetchone()
        self.assertEqual(before[0], after[0])
        self.assertIsNone(before[1])
        self.assertIsNotNone(after[1])


class ScoreTests(DbTestCase):
    def test_ctr_curve_interpolation(self):
        curve = {"1": 0.1, "10": 0.04, "50": 0.01}
        self.assertAlmostEqual(sem.ctr_at(30, curve), 0.025)
        self.assertEqual(sem.ctr_at(0.5, curve), 0.1)
        self.assertEqual(sem.ctr_at(500, curve), 0.01)
        self.assertEqual(sem.ctr_at(None, curve), 0.0)

    def test_score_is_frequency_times_ctr_gain_times_conversion(self):
        settings = {**self.settings, "ctr_curve": {"1": 0.1, "10": 0.04, "50": 0.01}, "unknown_position": 10}
        known = sem.score_pair(1000, 50, 2.0, settings)
        self.assertAlmostEqual(known["score"], 1000 * 0.09 * 2.0)
        self.assertTrue(known["position_known"])
        unknown = sem.score_pair(1000, None, 1.0, settings)
        self.assertAlmostEqual(unknown["score"], 1000 * 0.06)
        self.assertEqual(sem.score_pair(1000, 1, 1.0, settings)["score"], 0)

    def test_growth_zone_needs_known_position_in_range_and_frequent_query(self):
        self.stats("2026-10-20", 1, "доска разделочная", 1000, position=25)
        self.stats("2026-10-20", 1, "доска большая", 10, position=25)
        self.stats("2026-10-20", 1, "доска из кедра", 1000, position=5)
        self.stats("2026-10-20", 1, "доска круглая", 1000)
        sem.rebuild(self.conn, self.settings)
        self.assertEqual([r["query"] for r in sem.growth_zone(self.conn)], ["доска разделочная"])

    def test_conversion_factor_is_relative_and_clamped(self):
        with self.conn:
            db.upsert_sku_daily(self.conn, [
                {"date": "2026-10-20", "sku": 1, "searched_users": 1000, "ordered_units": 10},
                {"date": "2026-10-20", "sku": 2, "searched_users": 1000, "ordered_units": 5},
                {"date": "2026-10-20", "sku": 3, "searched_users": 1000, "ordered_units": 100},
                {"date": "2026-10-20", "sku": 4, "searched_users": 1000, "ordered_units": 0},
            ])
        factors = sem.conversion_factors(self.conn, 28)
        self.assertEqual(factors[1], 1.0)
        self.assertEqual(factors[3], 2.0)
        self.assertEqual(factors[4], 0.5)


class ManualInputTests(DbTestCase):
    def test_manual_keyword_is_target_and_does_not_override_api_frequency(self):
        self.stats("2026-10-20", 1, "доска разделочная", 100)
        sem.rebuild(self.conn, self.settings)
        sem.add_keywords(self.conn, ["доска разделочная", "доска для хлеба"], [1], frequency=5, priority=3)
        self.assertEqual(self.keyword("доска разделочная")["frequency"], 100)
        self.assertEqual(self.keyword("доска для хлеба")["frequency"], 5)
        link = self.conn.execute("SELECT is_target, manual_priority FROM sku_keywords sk JOIN keywords k ON k.id = sk.keyword_id "
                                 "WHERE k.query = 'доска для хлеба'").fetchone()
        self.assertEqual(tuple(link), (1, 3))

    def test_csv_cp1251_with_russian_headers_and_offer_id(self):
        with self.conn:
            db.upsert_products(self.conn, [{"sku": 77, "offer_id": "1302", "name": "Доска"}])
        self.assertEqual(sem.import_csv(self.conn, FIXTURES / "keywords_cp1251.csv"), 2)
        cedar = self.keyword("доска из кедра")
        self.assertEqual((cedar["frequency"], cedar["priority"], cedar["source"]), (1200, 2, "csv"))
        self.assertEqual(self.conn.execute("SELECT sku FROM sku_keywords WHERE keyword_id = ?", (cedar["id"],)).fetchone()[0], 77)
        self.assertIsNotNone(self.keyword("подарок повару"))

    def test_csv_comma_delimited_english_headers(self):
        rows = sem.parse_csv("query,sku,frequency\nдоска,123,10\n,,\n")
        self.assertEqual(rows, [{"query": "доска", "sku": "123", "offer_id": None, "frequency": 10.0, "priority": None}])

    def test_csv_without_query_column_is_rejected(self):
        with self.assertRaises(ValueError):
            sem.parse_csv("артикул;частотность\n1302;10\n")

    def test_unknown_offer_id_is_an_error(self):
        with self.assertRaises(ValueError):
            sem.resolve_skus(self.conn, [], ["нет-такого"])


class MigrationTwoTests(DbTestCase):
    def test_down_to_one_keeps_stage_one_data(self):
        self.stats("2026-10-20", 1, "доска", 10)
        sem.rebuild(self.conn, self.settings)
        db.migrate(self.conn, 1)
        self.assertEqual(self.conn.execute("SELECT COUNT(*) FROM query_stats").fetchone()[0], 1)
        columns = {r[1] for r in self.conn.execute("PRAGMA table_info(keywords)")}
        self.assertNotIn("lemma_key", columns)
        db.migrate(self.conn)
        self.assertIsNone(self.keyword("доска")["lemma_key"])


class CollectorCardsTests(DbTestCase):
    def test_card_failure_is_reported_but_not_fatal(self):
        with patch.object(api, "fetch_cards", side_effect=RuntimeError("HTTP 500")):
            self.assertFalse(collect.collect_cards(self.conn, Mock(), [1], collect.api.utc_today()))
        self.assertEqual(db.last_run(self.conn, collect.SOURCE_CARDS)["status"], "error")

    def test_semantics_failure_is_logged(self):
        with patch.object(sem, "rebuild", side_effect=ImportError("No module named 'pymorphy3'")):
            self.assertFalse(collect.update_semantics(self.conn))
        self.assertIn("pymorphy3", db.last_run(self.conn, collect.SOURCE_SEMANTICS)["message"])


if __name__ == "__main__":
    unittest.main()
