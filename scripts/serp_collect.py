"""
Замер позиций в выдаче Ozon по регионам (за флагом SERP_PARSER_ENABLED=1 в .env, по умолчанию выключен).

    python scripts/serp_collect.py            -> один замер в день: до 15 запросов в каждом активном регионе
    python scripts/serp_collect.py --dry-run  -> показать, какие запросы и регионы будут замерены, без браузера

Порядок: для каждого региона — главная Ozon с профилем региона → по каждому запросу порции выдачи через
JSON-эндпоинт, пока не найдены все наши SKU или не достигнута глубина → запись позиций. Перед записью регион
проверяется по коду ФИАС из ответа Ozon (location.current.uid): не совпал — замер региона не пишется.
Капча или блокировка — прогон останавливается целиком, повтор не раньше чем через cooldown_hours.

Код выхода: 0 — сделано или делать нечего, 1 — ошибки, 2 — уже идёт, 3 — Ozon показал капчу.
"""
import argparse
import gzip
import os
import random
import shutil
import sys
import time
from contextlib import closing
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

from dotenv import load_dotenv
try:
    from scripts.app_paths import ENV_FILE
except ModuleNotFoundError:
    from app_paths import ENV_FILE

try:
    from scripts import search_db as db
    from scripts import search_semantics as semantics
    from scripts import serp_browser as browser
    from scripts import serp_parser as parser
    from scripts import serp_paths as paths
    from scripts.file_lock import FileBusyError, exclusive_file
except ModuleNotFoundError:
    import search_db as db
    import search_semantics as semantics
    import serp_browser as browser
    import serp_parser as parser
    import serp_paths as paths
    from file_lock import FileBusyError, exclusive_file

SOURCE = "serp"
RAW_DIR = db.DATA_DIR / "serp_raw"
LOCK_PATH = db.DATA_DIR / "serp.lock"
EXIT_BLOCKED = 3


class RegionMismatch(Exception):
    """Ответ Ozon пришёл для другого региона, чем в профиле."""


def enabled() -> bool:
    load_dotenv(ENV_FILE)
    return (os.getenv("SERP_PARSER_ENABLED") or "").strip().lower() in ("1", "true", "yes", "да")


def blocked_until(conn, cooldown_hours: float) -> Optional[datetime]:
    row = conn.execute("SELECT started_at FROM collect_runs WHERE source = ? AND status = 'blocked' ORDER BY id DESC LIMIT 1",
                       (SOURCE,)).fetchone()
    if not row:
        return None
    until = datetime.fromisoformat(row["started_at"]) + timedelta(hours=cooldown_hours)
    return until if until > datetime.now(timezone.utc) else None


def plan_queries(conn, max_queries: int, fill_with_top: bool) -> List[Dict[str, Any]]:
    """
    Запросы для замера: сначала целевые (is_target) по ручному приоритету, затем — если fill_with_top — остальные
    по приоритету score. Для каждого запроса замеряются все связанные с ним товары: выдача одна, лишних запросов нет.
    """
    rows = conn.execute("""
        SELECT k.id, k.query, MAX(sk.is_target) AS target, MAX(COALESCE(sk.manual_priority, 0)) AS manual,
               k.priority, MAX(COALESCE(sk.score, 0)) AS score, COALESCE(k.frequency, 0) AS frequency,
               GROUP_CONCAT(sk.sku) AS skus
        FROM sku_keywords sk JOIN keywords k ON k.id = sk.keyword_id
        GROUP BY k.id
    """).fetchall()
    candidates = [r for r in rows if r["target"] or fill_with_top]
    candidates.sort(key=lambda r: (-r["target"], -r["manual"], -r["priority"], -r["score"], -r["frequency"], r["query"]))
    return [{"keyword_id": r["id"], "query": r["query"], "target": bool(r["target"]),
             "skus": sorted({int(s) for s in str(r["skus"]).split(",")})} for r in candidates[:max_queries]]


def measured_keywords(conn, day: str, region_id: int) -> set:
    return {r[0] for r in conn.execute("SELECT DISTINCT keyword_id FROM positions WHERE date = ? AND region_id = ? AND source = ?",
                                       (day, region_id, SOURCE))}


def save_raw(day: str, region_code: str, keyword_id: int, page_no: int, text: str) -> None:
    folder = RAW_DIR / day
    folder.mkdir(parents=True, exist_ok=True)
    with gzip.open(folder / f"{region_code}-{keyword_id}-p{page_no}.json.gz", "wt", encoding="utf-8") as stream:
        stream.write(text)


def prune_raw(today: date, keep_days: int) -> None:
    if not RAW_DIR.exists():
        return
    for folder in RAW_DIR.iterdir():
        try:
            folder_day = date.fromisoformat(folder.name)
        except ValueError:
            continue
        if folder_day < today - timedelta(days=keep_days):
            shutil.rmtree(folder, ignore_errors=True)


def pause(bounds, sleep=time.sleep) -> None:
    low, high = bounds
    sleep(random.uniform(float(low), float(high)))


def measure_query(page, region, query: Dict[str, Any], settings: Dict[str, Any], day: str, sleep=time.sleep) -> List[Dict[str, Any]]:
    """Листает выдачу запроса до нахождения всех SKU или до глубины. Возвращает строки позиций."""
    pages: List[List[Dict[str, Any]]] = []
    wanted = set(query["skus"])
    path: Optional[str] = paths.search_path(query["query"])
    seen_paths = set()
    while path and len(pages) < int(settings["max_pages"]) and sum(map(len, pages)) < int(settings["max_positions"]):
        if pages:
            pause(settings["page_delay_seconds"], sleep)
        if path in seen_paths:
            raise parser.SerpFormatError("Ozon вернул ту же ссылку на следующую порцию выдачи — листание остановлено.")
        seen_paths.add(path)
        response = browser.fetch_page_json(page, path)
        parsed = parser.parse_page(response["data"])
        save_raw(day, region["code"], query["keyword_id"], len(pages) + 1, response["raw"])
        if not pages:
            got = parsed["location"]
            if not got["fias"] or got["fias"] != region["fias"]:
                raise RegionMismatch(f"Ozon отдал выдачу для «{got['city'] or 'неизвестного города'}», "
                                     f"а профиль «{region['name']}». Пересоздайте профиль региона.")
        pages.append(parsed["items"])
        if wanted <= {item["sku"] for items in pages for item in items}:
            break
        path = parsed["next_page"]
    depth = sum(map(len, pages))
    return parser.positions_for(query["skus"], parser.rank(pages), depth)


def measure_region(conn, region, queries: List[Dict[str, Any]], settings: Dict[str, Any], day: str, sleep=time.sleep) -> int:
    """Замер одного региона. Капча — SerpBlocked наружу; несовпадение региона — RegionMismatch наружу (без записи)."""
    profile = db.ROOT / region["profile_path"]
    written = 0
    with browser.open_page(profile, region["proxy_env"], bool(settings["headless"]), float(settings["timeout_seconds"])) as page:
        browser.open_site(page)
        for i, query in enumerate(queries):
            if i:
                pause(settings["query_delay_seconds"], sleep)
            try:
                rows = measure_query(page, region, query, settings, day, sleep)
            except parser.SerpFormatError as e:
                print(f"⚠️ {region['name']}, «{query['query']}»: {e}", flush=True)
                continue
            with conn:
                db.upsert_positions(conn, day, region["id"], [{**r, "keyword_id": query["keyword_id"]} for r in rows], SOURCE)
            written += 1
            found = [f"{r['sku']} → {r['position']}" for r in rows if r["found"]]
            print(f"✅ {region['name']}, «{query['query']}»: " + (", ".join(found) if found else
                  f"наших товаров нет в первых {rows[0]['depth'] if rows else 0} позициях"), flush=True)
        browser.save_state(page, profile)
    return written


def run(dry_run: bool = False, sleep=time.sleep) -> int:
    settings = semantics.load_settings()["serp"]
    if not dry_run and not enabled():
        print("Парсер выдачи выключен (SERP_PARSER_ENABLED в .env). Позиции не замеряются.", flush=True)
        return 0
    LOCK_PATH.parent.mkdir(parents=True, exist_ok=True)
    try:
        with exclusive_file(LOCK_PATH), closing(db.connect()) as conn:
            until = blocked_until(conn, float(settings["cooldown_hours"]))
            if until and not dry_run:
                print(f"⛔ Ozon недавно показал капчу — парсер не запускается до {until.astimezone():%d.%m %H:%M}.", flush=True)
                return 0
            today = datetime.now(timezone.utc).date()
            day = today.isoformat()
            regions = [r for r in db.active_regions(conn) if r["profile_path"] and (db.ROOT / r["profile_path"]).exists()]
            if not regions:
                print("Нет регионов с профилем. Создайте профиль: scripts/serp_profile.py create msk \"Москва\".", flush=True)
                return 0
            queries = plan_queries(conn, int(settings["max_queries"]), bool(settings["fill_with_top"]))
            if not queries:
                print("Нет запросов для замера: сначала соберите поисковую аналитику или добавьте целевые запросы.", flush=True)
                return 0
            if dry_run:
                for r in regions:
                    print(f"Регион: {r['name']} ({r['code']}), ФИАС {r['fias']}")
                for q in queries:
                    print(f"{'★' if q['target'] else ' '} {q['query']} — товаров: {len(q['skus'])}")
                return 0
            if not browser.playwright_available():
                print(browser.INSTALL_HINT, flush=True)
                return 1
            prune_raw(today, int(settings["keep_raw_days"]))
            ok, measured_any = True, False
            for index, region in enumerate(regions):
                todo = [q for q in queries if q["keyword_id"] not in measured_keywords(conn, day, region["id"])]
                if not todo:
                    print(f"{region['name']}: сегодня уже замерено.", flush=True)
                    continue
                if index:
                    pause(settings["region_delay_seconds"], sleep)
                run_id = db.start_run(conn, SOURCE, day)
                try:
                    written = measure_region(conn, region, todo, settings, day, sleep)
                except parser.SerpBlocked as e:
                    db.finish_run(conn, run_id, "blocked", f"{region['code']}: {e}")
                    print(f"⛔ Ozon показал капчу ({region['name']}): {e} Замеры остановлены, повтор не раньше чем через "
                          f"{settings['cooldown_hours']} ч. Уменьшите число запросов или увеличьте паузы в search_settings.json.", flush=True)
                    return EXIT_BLOCKED
                except RegionMismatch as e:
                    db.finish_run(conn, run_id, "error", f"{region['code']}: {e}")
                    print(f"❌ {e}", flush=True)
                    ok = False
                    continue
                except Exception as e:
                    db.finish_run(conn, run_id, "error", f"{region['code']}: {e}")
                    print(f"❌ {region['name']}: замер прерван — {e}", flush=True)
                    ok = False
                    continue
                db.finish_run(conn, run_id, "ok", f"{region['code']}: запросов {written} из {len(todo)}")
                measured_any = measured_any or written > 0
            if measured_any:
                # Новые позиции сразу попадают в приоритет и зону роста.
                result = semantics.rebuild(conn)
                print(f"✅ Приоритеты пересчитаны: пар «товар — запрос» — {result['scored']}.", flush=True)
            return 0 if ok else 1
    except FileBusyError:
        print("Замер позиций уже идёт — этот запуск пропущен.", flush=True)
        return 2


def main(argv: Optional[List[str]] = None) -> int:
    parser_ = argparse.ArgumentParser(description="Замер позиций в выдаче Ozon по регионам.")
    parser_.add_argument("--dry-run", action="store_true", help="показать план замера без браузера")
    args = parser_.parse_args(argv)
    return run(dry_run=args.dry_run)


if __name__ == "__main__":
    sys.exit(main())
