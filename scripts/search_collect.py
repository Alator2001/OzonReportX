"""
Ежедневный сбор поисковой аналитики Ozon в `search data/search.db`.

    python scripts/search_collect.py                 -> догрузить пропущенные дни (28 дней), обновить 3 последних
    python scripts/search_collect.py --days 7 --refresh 1 --skip-sales

Запуск идемпотентен: каждый день пишется upsert-ом в одной транзакции, повторный запуск ничего не дублирует.
Последние --refresh дней перезагружаются всегда — Ozon досчитывает данные 1–2 дня.
Приложение запускает сборщик само при старте, если сегодня он ещё не отработал (см. docs/dev/17-search.md).

Код выхода: 0 — всё собрано (или источник недоступен без подписки), 1 — были ошибки, 2 — сбор уже идёт.
"""
import argparse
import json
import sys
import time
from contextlib import closing
from datetime import date, datetime, timedelta, timezone
from typing import List, Optional

try:
    from scripts import search_api as api
    from scripts import search_db as db
    from scripts import search_semantics as semantics
    from scripts.file_lock import FileBusyError, exclusive_file
except ModuleNotFoundError:
    import search_api as api
    import search_db as db
    import search_semantics as semantics
    from file_lock import FileBusyError, exclusive_file

SOURCE_PRODUCTS = "products"
SOURCE_CARDS = "cards"
SOURCE_SEMANTICS = "semantics"
SOURCE_QUERIES = "product_queries"
SOURCE_SALES = "sales"
SOURCE_SALES_RANGE = "sales_range"          # один запуск за диапазон; по дням — SOURCE_SALES
SOURCE_SALES_REQUEST = "analytics_data_request"


# Без подписки product-queries отдаёт день D только с полуночи D+2 (UTC): за вчера — 403 «premium subscription».
QUERIES_LAG_DAYS = 2
SALES_LAG_DAYS = 1


def days_to_collect(last_day: date, done: set, days: int, refresh: int) -> List[date]:
    """Дни окна [last_day − days + 1, last_day], которых ещё нет в журнале, плюс последние refresh дней — всегда."""
    window = [last_day - timedelta(days=n) for n in range(days - 1, -1, -1)]
    fresh = set(window[-refresh:]) if refresh > 0 else set()
    return [d for d in window if d.isoformat() not in done or d in fresh]


def collect_queries(conn, session, skus: List[int], days: List[date]) -> bool:
    """Запросы по SKU за каждый день. Возвращает False, если были ошибки (не считая отказа в доступе)."""
    for day in days:
        run_id = db.start_run(conn, SOURCE_QUERIES, day.isoformat())
        try:
            data = api.fetch_product_queries_day(session, skus, day)
            with conn:
                db.upsert_sku_daily(conn, data["sku_daily"])
                db.upsert_query_stats(conn, day.isoformat(), data["query_stats"])
        except api.AccessDenied as e:
            db.finish_run(conn, run_id, "denied", str(e))
            print(f"⚠️ Поисковая аналитика недоступна: {e}. Продолжаем без неё.", flush=True)
            return True
        except Exception as e:
            db.finish_run(conn, run_id, "error", str(e))
            print(f"❌ Запросы за {day:%d.%m.%Y} не загружены: {e}", flush=True)
            return False
        db.finish_run(conn, run_id, "ok", f"SKU: {len(data['sku_daily'])}, строк запросов: {len(data['query_stats'])}")
        print(f"✅ {day:%d.%m.%Y}: товаров в поиске — {len(data['sku_daily'])}, пар «товар — запрос» — {len(data['query_stats'])}", flush=True)
    return True


def _sales_request_gate(conn):
    """Перед каждым вызовом /v1/analytics/data: суточная квота и пауза минута между вызовами (в т.ч. из прошлых запусков)."""
    def gate() -> None:
        midnight = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0).isoformat()
        runs = db.runs_started_since(conn, SOURCE_SALES_REQUEST, midnight)
        if len(runs) >= api.ANALYTICS_DATA_DAILY_QUOTA:
            raise RuntimeError(f"сегодня уже сделано {len(runs)} запросов к /v1/analytics/data — "
                               "без Premium Plus Ozon разрешает 50 в сутки. Продажи догрузятся завтра.")
        if runs:
            last = datetime.fromisoformat(runs[-1]["started_at"]).timestamp()
            api.sleep_until(last + api.ANALYTICS_DATA_INTERVAL)
        run_id = db.start_run(conn, SOURCE_SALES_REQUEST)
        db.finish_run(conn, run_id, "ok")
    return gate


def collect_sales(conn, session, days: List[date]) -> bool:
    if not days:
        return True
    start, end = min(days), max(days)
    run_id = db.start_run(conn, SOURCE_SALES_RANGE, f"{start.isoformat()}..{end.isoformat()}")
    try:
        rows = api.fetch_sales_by_day(session, start, end, before_request=_sales_request_gate(conn))
        with conn:
            db.upsert_sku_daily(conn, rows)
            for day in days:
                conn.execute("INSERT INTO collect_runs (source, day, started_at, finished_at, status) VALUES (?, ?, ?, ?, 'ok')",
                             (SOURCE_SALES, day.isoformat(), db.now_iso(), db.now_iso()))
    except api.AccessDenied as e:
        db.finish_run(conn, run_id, "denied", str(e))
        print(f"⚠️ Продажи по дням недоступны: {e}. Продолжаем без них.", flush=True)
        return True
    except Exception as e:
        db.finish_run(conn, run_id, "error", str(e))
        print(f"❌ Продажи по дням не загружены: {e}", flush=True)
        return False
    db.finish_run(conn, run_id, "ok", f"строк: {len(rows)}")
    print(f"✅ Продажи по дням {start:%d.%m}–{end:%d.%m.%Y}: строк — {len(rows)}", flush=True)
    return True


def collect_cards(conn, session, skus: List[int], day: date) -> bool:
    """Снимок карточек на сегодня: тексты для матрицы покрытия и хеши для поиска изменений карточки."""
    run_id = db.start_run(conn, SOURCE_CARDS, day.isoformat())
    try:
        cards = api.fetch_cards(session, skus)
        with conn:
            db.upsert_card_snapshots(conn, day.isoformat(), cards)
    except Exception as e:
        db.finish_run(conn, run_id, "error", str(e))
        print(f"❌ Карточки не загружены: {e}", flush=True)
        return False
    db.finish_run(conn, run_id, "ok", f"карточек: {len(cards)}")
    print(f"✅ Карточки: {len(cards)}", flush=True)
    return True


def update_semantics(conn) -> bool:
    run_id = db.start_run(conn, SOURCE_SEMANTICS)
    try:
        result = semantics.rebuild(conn)
    except Exception as e:
        db.finish_run(conn, run_id, "error", str(e))
        print(f"❌ Семантика не пересчитана: {e}", flush=True)
        return False
    db.finish_run(conn, run_id, "ok", json.dumps(result, ensure_ascii=False))
    print(f"✅ Семантика: кластеров — {result['clusters']}, пар «товар — запрос» — {result['scored']}", flush=True)
    return True


def run(days: int = 28, refresh: int = 3, skip_sales: bool = False, db_path=None) -> int:
    path = db_path or db.DB_PATH
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with exclusive_file(path), closing(db.connect(path)) as conn, api.create_session() as session:
            started = time.monotonic()
            today = api.utc_today()
            run_id = db.start_run(conn, SOURCE_PRODUCTS, today.isoformat())
            try:
                products = api.load_products(session)
                with conn:
                    db.upsert_products(conn, products)
            except Exception as e:
                db.finish_run(conn, run_id, "error", str(e))
                print(f"❌ Не удалось загрузить список товаров: {e}", flush=True)
                return 1
            db.finish_run(conn, run_id, "ok", f"товаров: {len(products)}")
            skus = sorted({p["sku"] for p in products})
            print(f"Товаров с SKU: {len(skus)}", flush=True)
            query_days = days_to_collect(today - timedelta(days=QUERIES_LAG_DAYS), db.collected_days(conn, SOURCE_QUERIES), days, refresh)
            ok = collect_queries(conn, session, skus, query_days)
            ok = collect_cards(conn, session, skus, today) and ok
            if not skip_sales:
                sales_days = days_to_collect(today - timedelta(days=SALES_LAG_DAYS), db.collected_days(conn, SOURCE_SALES), days, refresh)
                ok = collect_sales(conn, session, sales_days) and ok
            ok = update_semantics(conn) and ok
            print(f"Сбор поисковой аналитики завершён за {time.monotonic() - started:.0f} с.", flush=True)
            return 0 if ok else 1
    except FileBusyError:
        print("Сбор поисковой аналитики уже идёт в другом окне — этот запуск пропущен.", flush=True)
        return 2


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Сбор поисковой аналитики Ozon.")
    parser.add_argument("--days", type=int, default=28, help="сколько дней назад догружать (Ozon без подписки хранит ~месяц)")
    parser.add_argument("--refresh", type=int, default=3, help="сколько последних дней перезагружать всегда")
    parser.add_argument("--skip-sales", action="store_true", help="не загружать продажи по дням (/v1/analytics/data)")
    args = parser.parse_args(argv)
    return run(max(1, args.days), max(0, args.refresh), args.skip_sales)


if __name__ == "__main__":
    sys.exit(main())
