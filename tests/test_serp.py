"""Offline checks for the Ozon search results parser: response parsing, ranking, region check, run control."""
import json
import shutil
import sys
import unittest
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import serp_collect as collect  # noqa: E402

# serp_collect импортирует соседей как scripts.x или как x — патчим те же объекты, что видит он.
db, sem, browser, parser = collect.db, collect.semantics, collect.browser, collect.parser
import serp_profile as profile  # noqa: E402

FIXTURES = ROOT / "tests" / "fixtures" / "serp"
TMP = ROOT / ".cache" / "serp-tests"
MSK_FIAS = "0c5b2444-70a0-4932-980c-b4dc0d3f02b5"
NO_PAUSE = lambda _seconds: None  # noqa: E731


def fixture(name):
    return json.loads((FIXTURES / name).read_text(encoding="utf-8"))


def response(name):
    data = fixture(name)
    return {"data": data, "raw": json.dumps(data, ensure_ascii=False)}


SETTINGS = {**sem.DEFAULT_SETTINGS["serp"], "page_delay_seconds": [0, 0], "query_delay_seconds": [0, 0]}
REGION = {"id": 1, "code": "msk", "name": "Москва", "fias": MSK_FIAS, "profile_path": "p.json", "proxy_env": None}


class ResponseTests(unittest.TestCase):
    def test_captcha_page_blocks(self):
        html = (FIXTURES / "captcha.html").read_text(encoding="utf-8")
        with self.assertRaises(parser.SerpBlocked):
            parser.check_response(403, "text/html; charset=utf-8", html)
        with self.assertRaises(parser.SerpBlocked):
            parser.check_response(200, "text/html", html)

    def test_rate_limit_blocks_without_retry(self):
        with self.assertRaises(parser.SerpBlocked):
            parser.check_response(429, "application/json", "{}")

    def test_unexpected_html_is_format_error(self):
        with self.assertRaises(parser.SerpFormatError):
            parser.check_response(200, "text/html", "<html><title>Ozon</title></html>")

    def test_json_is_parsed(self):
        self.assertEqual(parser.check_response(200, "application/json", '{"a": 1}'), {"a": 1})


class ParseTests(unittest.TestCase):
    def test_first_page_items_next_page_and_location(self):
        page = parser.parse_page(fixture("page1_msk.json"))
        self.assertEqual([i["sku"] for i in page["items"]], [1085894091, 4344859507, 1850332702])
        self.assertFalse(any(i["is_ad"] for i in page["items"]))  # advert/advertLite — токены у всех, не реклама
        self.assertTrue(page["next_page"].startswith("/category/"))
        self.assertEqual(page["location"], {"fias": MSK_FIAS, "city": "Москва"})

    def test_ad_label_and_last_page(self):
        page = parser.parse_page(fixture("page2_synthetic_ad_label.json"))
        self.assertEqual([i["is_ad"] for i in page["items"]], [True, False, False])
        self.assertIsNone(page["next_page"])
        self.assertEqual(page["location"], {"fias": None, "city": None})

    def test_missing_tiles_widget_is_format_error(self):
        with self.assertRaises(parser.SerpFormatError):
            parser.parse_page(fixture("no_tiles.json"))

    def test_rank_overall_and_organic_first_occurrence(self):
        pages = [parser.parse_page(fixture("page1_msk.json"))["items"], parser.parse_page(fixture("page2_synthetic_ad_label.json"))["items"]]
        ranked = parser.rank(pages)
        self.assertEqual(ranked[777], {"position": 4, "organic_position": None, "page": 2, "is_ad": 1})
        self.assertEqual(ranked[1850332681], {"position": 5, "organic_position": 4, "page": 2, "is_ad": 0})
        self.assertEqual(ranked[1085894091]["position"], 1)  # повтор на второй странице не перебивает первое место

    def test_not_found_is_null_not_999(self):
        rows = parser.positions_for([1, 1850332702], parser.rank([parser.parse_page(fixture("page1_msk.json"))["items"]]), 3)
        self.assertEqual(rows[0], {"sku": 1, "position": None, "organic_position": None, "page": None, "is_ad": None, "found": 0, "depth": 3})
        self.assertEqual((rows[1]["position"], rows[1]["found"]), (3, 1))


class MeasureTests(unittest.TestCase):
    def setUp(self):
        shutil.rmtree(TMP, ignore_errors=True)
        patcher = patch.object(collect, "RAW_DIR", TMP / "raw")
        patcher.start()
        self.addCleanup(patcher.stop)
        self.addCleanup(shutil.rmtree, TMP, True)

    def measure(self, responses, skus, settings=SETTINGS):
        query = {"keyword_id": 5, "query": "доска разделочная", "skus": skus}
        with patch.object(browser, "fetch_page_json", side_effect=responses) as fetch:
            rows = collect.measure_query(object(), REGION, query, settings, "2026-10-23", NO_PAUSE)
        return rows, fetch

    def test_pages_until_all_skus_found(self):
        rows, fetch = self.measure([response("page1_msk.json"), response("page2_synthetic_ad_label.json")], [1850332702, 1850332681])
        self.assertEqual(fetch.call_count, 2)
        self.assertEqual([(r["sku"], r["position"], r["organic_position"]) for r in rows], [(1850332702, 3, 3), (1850332681, 5, 4)])
        self.assertTrue((TMP / "raw" / "2026-10-23" / "msk-5-p2.json.gz").exists())

    def test_stops_early_when_everything_found_on_first_page(self):
        _, fetch = self.measure([response("page1_msk.json")], [1850332702])
        self.assertEqual(fetch.call_count, 1)

    def test_depth_limit_gives_not_found(self):
        rows, fetch = self.measure([response("page1_msk.json")], [999], {**SETTINGS, "max_pages": 1})
        self.assertEqual(fetch.call_count, 1)
        self.assertEqual((rows[0]["found"], rows[0]["position"], rows[0]["depth"]), (0, None, 3))

    def test_other_region_is_rejected_before_writing(self):
        with self.assertRaises(collect.RegionMismatch):
            self.measure([response("page1_spb.json")], [1850332702])

    def test_repeated_next_page_stops_paging(self):
        page = fixture("page1_msk.json")
        page["location"] = None
        first = response("page1_msk.json")
        loop = {"data": page, "raw": "{}"}
        state = json.loads(page["widgetStates"]["infiniteVirtualPaginator-3618992-default-1"])
        state["nextPage"] = "/search/?text=%D0%B4%D0%BE%D1%81%D0%BA%D0%B0%20%D1%80%D0%B0%D0%B7%D0%B4%D0%B5%D0%BB%D0%BE%D1%87%D0%BD%D0%B0%D1%8F&from_global=true"
        page["widgetStates"]["infiniteVirtualPaginator-3618992-default-1"] = json.dumps(state)
        with self.assertRaises(parser.SerpFormatError):
            self.measure([first, loop, loop], [999])

    def test_captcha_during_paging_propagates(self):
        with self.assertRaises(parser.SerpBlocked):
            self.measure([response("page1_msk.json"), parser.SerpBlocked("HTTP 403")], [999])


class DbTestCase(unittest.TestCase):
    def setUp(self):
        shutil.rmtree(TMP, ignore_errors=True)
        TMP.mkdir(parents=True)
        self.conn = db.connect(TMP / "search.db")
        self.addCleanup(shutil.rmtree, TMP, True)
        self.addCleanup(self.conn.close)

    def keyword(self, query, skus, target=False, score=0.0):
        keyword_id = db.ensure_keyword(self.conn, query, "api")
        for sku in skus:
            db.link_sku_keyword(self.conn, sku, keyword_id, "api")
            self.conn.execute("UPDATE sku_keywords SET is_target = ?, score = ? WHERE sku = ? AND keyword_id = ?", (int(target), score, sku, keyword_id))
        self.conn.commit()
        return keyword_id


class PositionsDbTests(DbTestCase):
    def test_upsert_positions_is_idempotent(self):
        region = db.upsert_region(self.conn, "msk", "Москва", fias=MSK_FIAS)
        kw = self.keyword("доска", [1])
        with self.conn:
            db.upsert_positions(self.conn, "2026-10-23", region, [{"sku": 1, "keyword_id": kw, "position": 7, "found": 1}])
            db.upsert_positions(self.conn, "2026-10-23", region, [{"sku": 1, "keyword_id": kw, "position": None, "found": 0}])
        rows = self.conn.execute("SELECT position, found FROM positions").fetchall()
        self.assertEqual([tuple(r) for r in rows], [(None, 0)])

    def test_median_high_treats_not_found_as_worst(self):
        msk = db.upsert_region(self.conn, "msk", "Москва")
        spb = db.upsert_region(self.conn, "spb", "Санкт-Петербург")
        kw = self.keyword("доска", [1, 2])
        with self.conn:
            for day, region, sku, pos in [("2026-10-21", msk, 1, 10), ("2026-10-22", msk, 1, 30), ("2026-10-23", msk, 1, 20),
                                          ("2026-10-19", msk, 1, 1),  # старше окна в 3 дня
                                          ("2026-10-23", msk, 2, 3), ("2026-10-23", spb, 2, None)]:
                db.upsert_positions(self.conn, day, region, [{"sku": sku, "keyword_id": kw, "position": pos, "found": int(pos is not None)}])
        result = sem.serp_positions(self.conn, 3)
        self.assertEqual(result[(1, kw)], {"position": 20.0, "measurements": 3, "found": 3})
        self.assertEqual(result[(2, kw)]["position"], None)
        self.assertEqual(sem.serp_positions(self.conn, 3, region_id=msk)[(2, kw)]["position"], 3.0)

    def test_not_found_gets_full_ctr_gain_unknown_gets_default(self):
        settings = {**sem.DEFAULT_SETTINGS, "ctr_curve": {"1": 0.1, "30": 0.02}, "unknown_position": 30}
        self.assertAlmostEqual(sem.score_pair(100, None, 1.0, settings, known=True)["ctr_gain"], 0.1)
        self.assertAlmostEqual(sem.score_pair(100, None, 1.0, settings, known=False)["ctr_gain"], 0.08)

    def test_plan_targets_first_then_fill(self):
        self.keyword("топ по приоритету", [1], score=500)
        self.keyword("целевой", [1, 2], target=True, score=1)
        self.keyword("второй", [2], score=100)
        plan = collect.plan_queries(self.conn, 2, True)
        self.assertEqual([q["query"] for q in plan], ["целевой", "топ по приоритету"])
        self.assertEqual(plan[0]["skus"], [1, 2])
        self.assertEqual([q["query"] for q in collect.plan_queries(self.conn, 15, False)], ["целевой"])

    def test_migration_three_down_and_up(self):
        db.migrate(self.conn, 2)
        tables = {r[0] for r in self.conn.execute("SELECT name FROM sqlite_master WHERE type = 'table'")}
        self.assertNotIn("positions", tables)
        db.migrate(self.conn)
        self.assertEqual(db.current_version(self.conn), db.LATEST_VERSION)


class RunControlTests(DbTestCase):
    def setUp(self):
        super().setUp()
        for target, value in [(db, ("DB_PATH", TMP / "search.db")), (db, ("ROOT", TMP)), (collect, ("LOCK_PATH", TMP / "serp.lock")),
                              (collect, ("RAW_DIR", TMP / "raw")), (sem, ("SETTINGS_PATH", TMP / "none.json"))]:
            patcher = patch.object(target, *value)
            patcher.start()
            self.addCleanup(patcher.stop)
        (TMP / "msk.json").write_text("{}", encoding="utf-8")
        with self.conn:
            db.upsert_region(self.conn, "msk", "Москва", fias=MSK_FIAS, profile_path="msk.json")
        self.keyword("доска", [1], target=True)

    def test_disabled_flag_does_nothing(self):
        with patch.object(collect, "enabled", return_value=False), patch.object(collect, "measure_region") as measure:
            self.assertEqual(collect.run(sleep=NO_PAUSE), 0)
        measure.assert_not_called()

    def test_captcha_stops_run_and_starts_cooldown(self):
        with patch.object(collect, "enabled", return_value=True), patch.object(browser, "playwright_available", return_value=True), \
                patch.object(collect, "measure_region", side_effect=parser.SerpBlocked("HTTP 403")) as measure:
            self.assertEqual(collect.run(sleep=NO_PAUSE), collect.EXIT_BLOCKED)
            self.assertEqual(collect.run(sleep=NO_PAUSE), 0)  # второй запуск — пауза после капчи, браузер не открывается
        self.assertEqual(measure.call_count, 1)
        self.assertEqual(db.last_run(self.conn, collect.SOURCE)["status"], "blocked")
        self.assertEqual(self.conn.execute("SELECT COUNT(*) FROM positions").fetchone()[0], 0)

    def test_cooldown_expires(self):
        old = (datetime.now(timezone.utc) - timedelta(hours=25)).replace(microsecond=0).isoformat()
        self.conn.execute("INSERT INTO collect_runs (source, started_at, status) VALUES (?, ?, 'blocked')", (collect.SOURCE, old))
        self.conn.commit()
        self.assertIsNone(collect.blocked_until(self.conn, 24))

    def test_missing_playwright_is_reported(self):
        with patch.object(collect, "enabled", return_value=True), patch.object(browser, "playwright_available", return_value=False):
            self.assertEqual(collect.run(sleep=NO_PAUSE), 1)


class ProfileAndProxyTests(unittest.TestCase):
    def test_proxy_from_env_with_credentials(self):
        with patch.dict("os.environ", {"SERP_PROXY_SPB": "http://user:p%40ss@10.0.0.1:3128"}), patch.object(browser, "load_dotenv"):
            self.assertEqual(browser.proxy_settings("SERP_PROXY_SPB"),
                             {"server": "http://10.0.0.1:3128", "username": "user", "password": "p@ss"})
        self.assertIsNone(browser.proxy_settings(None))

    def test_missing_proxy_variable_is_an_error(self):
        with patch.dict("os.environ", {}, clear=True), patch.object(browser, "load_dotenv"):
            with self.assertRaises(RuntimeError):
                browser.proxy_settings("SERP_PROXY_SPB")

    def test_cookie_diff_hides_values(self):
        a = profile.cookie_fingerprint({"cookies": [{"domain": ".ozon.ru", "name": "xcid", "value": "secret-1"},
                                                    {"domain": ".ozon.ru", "name": "only_a", "value": "1"}]})
        b = profile.cookie_fingerprint({"cookies": [{"domain": ".ozon.ru", "name": "xcid", "value": "secret-2"}]})
        self.assertNotIn("secret-1", json.dumps(a))
        self.assertEqual(profile.diff_fingerprints(a, b), {"only_a": [".ozon.ru|only_a"], "only_b": [], "different": [".ozon.ru|xcid"]})

    def test_nested_serp_settings_are_merged(self):
        path = TMP / "settings.json"
        TMP.mkdir(parents=True, exist_ok=True)
        self.addCleanup(shutil.rmtree, TMP, True)
        path.write_text(json.dumps({"serp": {"max_queries": 5}}), encoding="utf-8")
        settings = sem.load_settings(path)
        self.assertEqual(settings["serp"]["max_queries"], 5)
        self.assertEqual(settings["serp"]["median_days"], 3)


if __name__ == "__main__":
    unittest.main()
