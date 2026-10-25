"""Offline checks for the search analytics section: DB migrations, upserts, Ozon response parsing, collector."""
import json
import shutil
import sqlite3
import sys
import unittest
from contextlib import closing
from datetime import date
from pathlib import Path
from unittest.mock import Mock, patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import search_collect as collect  # noqa: E402

# Сборщик импортирует соседей как scripts.x или как x — в зависимости от sys.path. Патчим те же объекты, что видит он.
api, db = collect.api, collect.db

FIXTURES = ROOT / "tests" / "fixtures" / "search_api"
TMP = ROOT / ".cache" / "search-tests"


def fixture(name):
    return json.loads((FIXTURES / name).read_text(encoding="utf-8"))


def tables(conn):
    return {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%'")}


class DbTestCase(unittest.TestCase):
    def setUp(self):
        shutil.rmtree(TMP, ignore_errors=True)
        TMP.mkdir(parents=True)
        self.path = TMP / "search.db"
        self.conn = db.connect(self.path)

    def tearDown(self):
        self.conn.close()
        shutil.rmtree(TMP, ignore_errors=True)


class MigrationTests(DbTestCase):
    def test_migrate_down_and_up_again(self):
        self.assertEqual(db.current_version(self.conn), db.LATEST_VERSION)
        self.assertIn("query_stats", tables(self.conn))
        self.assertEqual(db.migrate(self.conn, 0), 0)
        self.assertEqual(tables(self.conn), {"schema_version"})
        self.assertEqual(db.migrate(self.conn), db.LATEST_VERSION)
        self.assertIn("query_stats", tables(self.conn))

    def test_failed_migration_rolls_back_and_keeps_version(self):
        db.migrate(self.conn, 0)
        broken = [(1, "broken", "CREATE TABLE a (x INTEGER); CREATE TABLE a (x INTEGER)", "DROP TABLE a")]
        with patch.object(db, "MIGRATIONS", broken), patch.object(db, "LATEST_VERSION", 1):
            with self.assertRaises(sqlite3.OperationalError):
                db.migrate(self.conn)
        self.assertEqual(db.current_version(self.conn), 0)
        self.assertNotIn("a", tables(self.conn))

    def test_unknown_target_is_rejected(self):
        with self.assertRaises(ValueError):
            db.migrate(self.conn, db.LATEST_VERSION + 1)


class UpsertTests(DbTestCase):
    def test_query_stats_upsert_is_idempotent_and_updates_values(self):
        row = {"sku": 1, "query": "доска разделочная", "query_index": 1, "searched_users": 10}
        with self.conn:
            db.upsert_query_stats(self.conn, "2026-10-20", [row])
            db.upsert_query_stats(self.conn, "2026-10-20", [{**row, "searched_users": 12}])
        rows = self.conn.execute("SELECT * FROM query_stats").fetchall()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["searched_users"], 12)
        self.assertEqual(self.conn.execute("SELECT COUNT(*) FROM keywords").fetchone()[0], 1)
        self.assertEqual(self.conn.execute("SELECT COUNT(*) FROM sku_keywords").fetchone()[0], 1)

    def test_keyword_keeps_original_text_and_normalized_form(self):
        keyword_id = db.ensure_keyword(self.conn, "  Ёлочная   Игрушка ", "manual")
        row = self.conn.execute("SELECT query, normalized FROM keywords WHERE id = ?", (keyword_id,)).fetchone()
        self.assertEqual((row["query"], row["normalized"]), ("Ёлочная Игрушка", "елочная игрушка"))
        self.assertEqual(db.ensure_keyword(self.conn, "Ёлочная Игрушка", "api"), keyword_id)

    def test_sku_daily_sources_do_not_overwrite_each_other(self):
        with self.conn:
            db.upsert_sku_daily(self.conn, [{"date": "2026-10-20", "sku": 1, "searched_users": 100, "position": None}])
            db.upsert_sku_daily(self.conn, [{"date": "2026-10-20", "sku": 1, "ordered_units": 2.0, "revenue": 3000.0}])
        row = self.conn.execute("SELECT * FROM sku_daily").fetchone()
        self.assertEqual((row["searched_users"], row["ordered_units"], row["revenue"]), (100, 2.0, 3000.0))

    def test_products_upsert_keeps_known_fields(self):
        with self.conn:
            db.upsert_products(self.conn, [{"sku": 5, "offer_id": "A-1", "name": "Доска"}])
            db.upsert_products(self.conn, [{"sku": 5, "offer_id": None, "name": "Доска 2"}])
        row = self.conn.execute("SELECT offer_id, name FROM products WHERE sku = 5").fetchone()
        self.assertEqual((row["offer_id"], row["name"]), ("A-1", "Доска 2"))


class ParsingTests(unittest.TestCase):
    def test_product_queries_without_premium_keep_nulls(self):
        rows = api.parse_product_queries(fixture("product_queries_no_premium.json")["items"], "2026-10-20")
        self.assertEqual(rows[0], {"date": "2026-10-20", "sku": 1850332702, "searched_users": 13648, "viewed_users": None,
                                   "view_conversion": None, "position": None, "gmv": 0.0})
        self.assertEqual(rows[1]["searched_users"], 0)

    def test_query_details_collapse_spaces_and_skip_empty(self):
        rows = api.parse_query_details(fixture("product_queries_details_no_premium.json")["queries"])
        self.assertEqual([r["query"] for r in rows], ["доска разделочная", "Доска из кедра"])
        self.assertEqual(rows[0]["searched_users"], 3772)
        self.assertIsNone(rows[0]["position"])
        self.assertEqual(rows[0]["orders"], 0)

    def test_analytics_data_rows(self):
        rows = api.parse_analytics_data(fixture("analytics_data_sku_day.json"), ["ordered_units", "revenue"])
        self.assertEqual(rows[1], {"sku": 1850332668, "date": "2026-10-18", "ordered_units": 1.0, "revenue": 990.5})

    def test_analytics_data_with_dropped_metric_is_an_error(self):
        # Без подписки Ozon молча убирает недоступные метрики — нельзя записать их значения не в те столбцы.
        with self.assertRaises(RuntimeError):
            api.parse_analytics_data(fixture("analytics_data_sku_day.json"), ["ordered_units", "revenue", "hits_view_search"])

    def test_day_bounds_are_half_open_utc(self):
        self.assertEqual(api.day_bounds(date(2026, 10, 20)),
                         {"date_from": "2026-10-20T00:00:00Z", "date_to": "2026-10-21T00:00:00Z"})


class FetchTests(unittest.TestCase):
    @staticmethod
    def response(status, data):
        return Mock(status_code=status, json=Mock(return_value=data), text=json.dumps(data, ensure_ascii=False))

    def setUp(self):
        patcher = patch.object(api, "_headers", return_value={"Client-Id": "1", "Api-Key": "k"})
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_details_requested_only_for_skus_found_in_search(self):
        session = Mock()
        session.post.side_effect = [self.response(200, fixture("product_queries_no_premium.json")),
                                    self.response(200, fixture("product_queries_details_no_premium.json"))]
        data = api.fetch_product_queries_day(session, [1850332702, 1850332681], date(2026, 10, 20))
        details_body = session.post.call_args_list[1].kwargs["json"]
        self.assertEqual(details_body["skus"], ["1850332702"])
        self.assertEqual(details_body["limit_by_sku"], 15)
        self.assertEqual(len(data["sku_daily"]), 2)
        self.assertEqual(len(data["query_stats"]), 2)

    def test_403_raises_access_denied(self):
        session = Mock()
        session.post.return_value = self.response(403, {"code": 7, "message": "Method available with Premium Pro subscription"})
        with self.assertRaises(api.AccessDenied) as ctx:
            api.fetch_product_queries_day(session, [1], date(2026, 10, 20))
        self.assertIn("Premium Pro", str(ctx.exception))

    def test_other_errors_raise(self):
        session = Mock()
        session.post.return_value = self.response(500, {"message": "internal"})
        with self.assertRaises(RuntimeError):
            api.fetch_product_queries_day(session, [1], date(2026, 10, 20))

    def test_product_list_detects_repeated_cursor(self):
        session = Mock()
        page = {"result": {"items": [{"product_id": 1}], "last_id": "same"}}
        session.post.side_effect = [self.response(200, page), self.response(200, page)]
        with self.assertRaises(RuntimeError):
            api.load_products(session)


class CollectorTests(DbTestCase):
    def test_days_to_collect_skips_done_but_refreshes_last_days(self):
        done = {"2026-10-18", "2026-10-19", "2026-10-20", "2026-10-21", "2026-10-22"}
        days = collect.days_to_collect(date(2026, 10, 22), done, days=6, refresh=2)
        self.assertEqual(days, [date(2026, 10, 17), date(2026, 10, 21), date(2026, 10, 22)])

    def test_failed_day_is_not_written_and_not_marked_done(self):
        good = {"sku_daily": [{"date": "2026-10-20", "sku": 1, "searched_users": 5}],
                "query_stats": [{"sku": 1, "query": "доска", "searched_users": 5}]}
        with patch.object(api, "fetch_product_queries_day", side_effect=[good, RuntimeError("HTTP 500")]):
            ok = collect.collect_queries(self.conn, Mock(), [1], [date(2026, 10, 20), date(2026, 10, 21)])
        self.assertFalse(ok)
        self.assertEqual(db.collected_days(self.conn, collect.SOURCE_QUERIES), {"2026-10-20"})
        self.assertEqual(self.conn.execute("SELECT COUNT(*) FROM query_stats").fetchone()[0], 1)

    def test_access_denied_is_logged_and_not_an_error(self):
        with patch.object(api, "fetch_product_queries_day", side_effect=api.AccessDenied("нужна подписка")):
            ok = collect.collect_queries(self.conn, Mock(), [1], [date(2026, 10, 20)])
        self.assertTrue(ok)
        self.assertEqual(db.last_run(self.conn, collect.SOURCE_QUERIES)["status"], "denied")

    def test_sales_quota_stops_before_request(self):
        for _ in range(api.ANALYTICS_DATA_DAILY_QUOTA):
            db.start_run(self.conn, collect.SOURCE_SALES_REQUEST)
        session = Mock()
        ok = collect.collect_sales(self.conn, session, [date(2026, 10, 20)])
        self.assertFalse(ok)
        session.post.assert_not_called()
        self.assertEqual(db.collected_days(self.conn, collect.SOURCE_SALES), set())

    def test_sales_marks_every_day_of_range(self):
        rows = [{"sku": 1, "date": "2026-10-20", "ordered_units": 1.0, "revenue": 100.0}]
        with patch.object(api, "fetch_sales_by_day", return_value=rows):
            self.assertTrue(collect.collect_sales(self.conn, Mock(), [date(2026, 10, 20), date(2026, 10, 21)]))
        self.assertEqual(db.collected_days(self.conn, collect.SOURCE_SALES), {"2026-10-20", "2026-10-21"})


if __name__ == "__main__":
    unittest.main()
