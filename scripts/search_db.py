"""
База раздела «Поиск и ключевые слова»: SQLite-файл `search data/search.db`.

Схема меняется только миграциями из MIGRATIONS: у каждой есть шаг вперёд (up) и откат (down),
номер применённой версии хранится в таблице schema_version. Запись — только через upsert-функции
ниже, поэтому повторный сбор того же дня ничего не дублирует.

    python scripts/search_db.py migrate            -> до последней версии
    python scripts/search_db.py migrate --to 0     -> откатить всё
    python scripts/search_db.py version
"""
import argparse
import hashlib
import json
import sqlite3
from contextlib import closing
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
try:
    from scripts.app_paths import DATA_ROOT
except ModuleNotFoundError:
    from app_paths import DATA_ROOT

ROOT = DATA_ROOT
DATA_DIR = ROOT / "search data"
DB_PATH = DATA_DIR / "search.db"

# (версия, описание, up, down). Новые миграции — только в конец списка, старые не править.
MIGRATIONS: List[Tuple[int, str, str, str]] = [
    (1, "Товары, запросы, статистика запросов из Seller API, журнал сбора", """
        CREATE TABLE products (
            sku INTEGER PRIMARY KEY,
            offer_id TEXT,
            name TEXT,
            category TEXT,
            updated_at TEXT NOT NULL
        );
        CREATE TABLE keywords (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            query TEXT NOT NULL UNIQUE,
            normalized TEXT NOT NULL,
            cluster_id INTEGER,
            frequency REAL,
            priority INTEGER NOT NULL DEFAULT 0,
            source TEXT NOT NULL,
            created_at TEXT NOT NULL
        );
        CREATE INDEX keywords_normalized ON keywords(normalized);
        CREATE TABLE sku_keywords (
            sku INTEGER NOT NULL,
            keyword_id INTEGER NOT NULL REFERENCES keywords(id) ON DELETE CASCADE,
            source TEXT NOT NULL,
            created_at TEXT NOT NULL,
            PRIMARY KEY (sku, keyword_id)
        );
        CREATE TABLE query_stats (
            date TEXT NOT NULL,
            sku INTEGER NOT NULL,
            keyword_id INTEGER NOT NULL REFERENCES keywords(id) ON DELETE CASCADE,
            query_index INTEGER,
            searched_users INTEGER,
            viewed_users INTEGER,
            view_conversion REAL,
            position REAL,
            orders INTEGER,
            gmv REAL,
            collected_at TEXT NOT NULL,
            PRIMARY KEY (date, sku, keyword_id)
        );
        CREATE TABLE sku_daily (
            date TEXT NOT NULL,
            sku INTEGER NOT NULL,
            searched_users INTEGER,
            viewed_users INTEGER,
            view_conversion REAL,
            position REAL,
            gmv REAL,
            ordered_units REAL,
            revenue REAL,
            collected_at TEXT NOT NULL,
            PRIMARY KEY (date, sku)
        );
        CREATE TABLE collect_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source TEXT NOT NULL,
            day TEXT,
            started_at TEXT NOT NULL,
            finished_at TEXT,
            status TEXT NOT NULL,
            message TEXT
        );
        CREATE INDEX collect_runs_source_day ON collect_runs(source, day);
    """, """
        DROP TABLE collect_runs;
        DROP TABLE sku_daily;
        DROP TABLE query_stats;
        DROP TABLE sku_keywords;
        DROP TABLE keywords;
        DROP TABLE products;
    """),
    (2, "Семантика: леммы и кластеры запросов, целевые запросы товаров, снимки карточек", """
        ALTER TABLE products ADD COLUMN product_id INTEGER;
        ALTER TABLE keywords ADD COLUMN lemma_key TEXT;
        ALTER TABLE keywords ADD COLUMN cluster_locked INTEGER NOT NULL DEFAULT 0;
        ALTER TABLE keywords ADD COLUMN frequency_source TEXT;
        CREATE INDEX keywords_lemma_key ON keywords(lemma_key);
        CREATE INDEX keywords_cluster ON keywords(cluster_id);
        CREATE TABLE clusters (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            lemma_key TEXT NOT NULL UNIQUE,
            name TEXT NOT NULL,
            created_at TEXT NOT NULL
        );
        ALTER TABLE sku_keywords ADD COLUMN is_target INTEGER NOT NULL DEFAULT 0;
        ALTER TABLE sku_keywords ADD COLUMN manual_priority INTEGER;
        ALTER TABLE sku_keywords ADD COLUMN score REAL;
        ALTER TABLE sku_keywords ADD COLUMN score_details TEXT;
        CREATE TABLE card_snapshots (
            date TEXT NOT NULL,
            sku INTEGER NOT NULL,
            name TEXT,
            description TEXT,
            attributes TEXT,
            images TEXT,
            name_hash TEXT,
            description_hash TEXT,
            attributes_hash TEXT,
            images_hash TEXT,
            rich_content_hash TEXT,
            collected_at TEXT NOT NULL,
            PRIMARY KEY (date, sku)
        )
    """, """
        DROP TABLE card_snapshots;
        ALTER TABLE sku_keywords DROP COLUMN score_details;
        ALTER TABLE sku_keywords DROP COLUMN score;
        ALTER TABLE sku_keywords DROP COLUMN manual_priority;
        ALTER TABLE sku_keywords DROP COLUMN is_target;
        DROP TABLE clusters;
        DROP INDEX keywords_cluster;
        DROP INDEX keywords_lemma_key;
        ALTER TABLE keywords DROP COLUMN frequency_source;
        ALTER TABLE keywords DROP COLUMN cluster_locked;
        ALTER TABLE keywords DROP COLUMN lemma_key;
        ALTER TABLE products DROP COLUMN product_id
    """),
    (3, "Парсер выдачи: регионы и позиции", """
        CREATE TABLE regions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT NOT NULL UNIQUE,
            name TEXT NOT NULL,
            fias TEXT,
            profile_path TEXT,
            proxy_env TEXT,
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        CREATE TABLE positions (
            date TEXT NOT NULL,
            measured_at TEXT NOT NULL,
            sku INTEGER NOT NULL,
            keyword_id INTEGER NOT NULL REFERENCES keywords(id) ON DELETE CASCADE,
            region_id INTEGER NOT NULL REFERENCES regions(id) ON DELETE CASCADE,
            position INTEGER,
            organic_position INTEGER,
            page INTEGER,
            is_ad INTEGER,
            found INTEGER NOT NULL,
            depth INTEGER,
            source TEXT NOT NULL,
            PRIMARY KEY (date, sku, keyword_id, region_id, source)
        );
        CREATE INDEX positions_keyword ON positions(keyword_id, date)
    """, """
        DROP INDEX positions_keyword;
        DROP TABLE positions;
        DROP TABLE regions
    """),
]

LATEST_VERSION = MIGRATIONS[-1][0]


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def normalize_query(text: str) -> str:
    """Базовая нормализация: нижний регистр, ё → е, лишние пробелы. Леммы — на этапе семантики."""
    return " ".join(str(text or "").lower().replace("ё", "е").split())


def connect(path: Optional[Path] = None, migrate_to_latest: bool = True) -> sqlite3.Connection:
    path = Path(path or DB_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    if migrate_to_latest:
        migrate(conn)
    return conn


def current_version(conn: sqlite3.Connection) -> int:
    conn.execute("CREATE TABLE IF NOT EXISTS schema_version (version INTEGER NOT NULL)")
    row = conn.execute("SELECT version FROM schema_version").fetchone()
    return int(row[0]) if row else 0


def _set_version(conn: sqlite3.Connection, version: int) -> None:
    conn.execute("DELETE FROM schema_version")
    conn.execute("INSERT INTO schema_version (version) VALUES (?)", (version,))


def _run_script(conn: sqlite3.Connection, script: str) -> None:
    # executescript делает COMMIT перед выполнением, поэтому идём по операторам внутри своей транзакции.
    for statement in script.split(";"):
        if statement.strip():
            conn.execute(statement)


def migrate(conn: sqlite3.Connection, target: Optional[int] = None) -> int:
    """Приводит схему к версии target (по умолчанию — последней). Каждый шаг — отдельная транзакция."""
    target = LATEST_VERSION if target is None else target
    if not 0 <= target <= LATEST_VERSION:
        raise ValueError(f"Нет миграции с номером {target}; последняя — {LATEST_VERSION}.")
    version = current_version(conn)
    conn.commit()
    by_version = {m[0]: m for m in MIGRATIONS}
    while version != target:
        step = version + 1 if target > version else version
        _, _, up, down = by_version[step]
        try:
            conn.execute("BEGIN")
            _run_script(conn, up if target > version else down)
            version = step if target > version else step - 1
            _set_version(conn, version)
            conn.commit()
        except Exception:
            conn.rollback()
            raise
    return version


# --- upsert -----------------------------------------------------------------------------------

def upsert_products(conn: sqlite3.Connection, items: Iterable[Dict[str, Any]]) -> int:
    rows = [(int(i["sku"]), i.get("product_id"), i.get("offer_id"), i.get("name"), i.get("category"), now_iso()) for i in items]
    conn.executemany("""
        INSERT INTO products (sku, product_id, offer_id, name, category, updated_at) VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(sku) DO UPDATE SET
            product_id = COALESCE(excluded.product_id, product_id),
            offer_id = COALESCE(excluded.offer_id, offer_id),
            name = COALESCE(excluded.name, name),
            category = COALESCE(excluded.category, category),
            updated_at = excluded.updated_at
    """, rows)
    return len(rows)


def ensure_keyword(conn: sqlite3.Connection, query: str, source: str) -> int:
    """id запроса; новый запрос создаётся. Ключ — текст запроса как его отдал Ozon или ввёл пользователь."""
    query = " ".join(str(query or "").split())
    if not query:
        raise ValueError("Пустой поисковый запрос.")
    conn.execute("""
        INSERT INTO keywords (query, normalized, source, created_at) VALUES (?, ?, ?, ?)
        ON CONFLICT(query) DO NOTHING
    """, (query, normalize_query(query), source, now_iso()))
    return int(conn.execute("SELECT id FROM keywords WHERE query = ?", (query,)).fetchone()[0])


def link_sku_keyword(conn: sqlite3.Connection, sku: int, keyword_id: int, source: str) -> None:
    conn.execute("""
        INSERT INTO sku_keywords (sku, keyword_id, source, created_at) VALUES (?, ?, ?, ?)
        ON CONFLICT(sku, keyword_id) DO NOTHING
    """, (int(sku), int(keyword_id), source, now_iso()))


QUERY_STAT_FIELDS = ("query_index", "searched_users", "viewed_users", "view_conversion", "position", "orders", "gmv")


def upsert_query_stats(conn: sqlite3.Connection, day: str, rows: Sequence[Dict[str, Any]], source: str = "api") -> int:
    """rows: {sku, query, ...QUERY_STAT_FIELDS}. Запрос и связь с SKU создаются автоматически."""
    stamp = now_iso()
    for row in rows:
        keyword_id = ensure_keyword(conn, row["query"], source)
        link_sku_keyword(conn, row["sku"], keyword_id, source)
        values = [row.get(field) for field in QUERY_STAT_FIELDS]
        conn.execute(f"""
            INSERT INTO query_stats (date, sku, keyword_id, {", ".join(QUERY_STAT_FIELDS)}, collected_at)
            VALUES (?, ?, ?, {", ".join("?" * len(QUERY_STAT_FIELDS))}, ?)
            ON CONFLICT(date, sku, keyword_id) DO UPDATE SET
                {", ".join(f"{f} = excluded.{f}" for f in QUERY_STAT_FIELDS)},
                collected_at = excluded.collected_at
        """, [day, int(row["sku"]), keyword_id, *values, stamp])
    return len(rows)


SKU_DAILY_FIELDS = ("searched_users", "viewed_users", "view_conversion", "position", "gmv", "ordered_units", "revenue")


def upsert_sku_daily(conn: sqlite3.Connection, rows: Sequence[Dict[str, Any]]) -> int:
    """
    rows: {date, sku, <часть SKU_DAILY_FIELDS>}. Обновляются только переданные поля: product-queries
    и analytics/data пишут в одну строку разные столбцы и не затирают друг друга.
    """
    stamp = now_iso()
    for row in rows:
        fields = [f for f in SKU_DAILY_FIELDS if f in row]
        columns = ", ".join(["date", "sku", *fields, "collected_at"])
        updates = ", ".join([*(f"{f} = excluded.{f}" for f in fields), "collected_at = excluded.collected_at"])
        conn.execute(f"""
            INSERT INTO sku_daily ({columns}) VALUES ({", ".join("?" * (len(fields) + 3))})
            ON CONFLICT(date, sku) DO UPDATE SET {updates}
        """, [row["date"], int(row["sku"]), *(row[f] for f in fields), stamp])
    return len(rows)


def _hash(value: Any) -> Optional[str]:
    if value in (None, "", [], {}):
        return None
    return hashlib.sha1(json.dumps(value, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()


def upsert_card_snapshots(conn: sqlite3.Connection, day: str, cards: Sequence[Dict[str, Any]]) -> int:
    """Снимок карточек на день: тексты и хеши частей — по хешам этап экспериментов находит изменения карточки."""
    stamp = now_iso()
    for card in cards:
        attributes = json.dumps(card.get("attributes") or {}, ensure_ascii=False, sort_keys=True)
        images = json.dumps(card.get("images") or [], ensure_ascii=False)
        conn.execute("""
            INSERT INTO card_snapshots (date, sku, name, description, attributes, images, name_hash, description_hash,
                                        attributes_hash, images_hash, rich_content_hash, collected_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(date, sku) DO UPDATE SET
                name = excluded.name, description = excluded.description, attributes = excluded.attributes,
                images = excluded.images, name_hash = excluded.name_hash, description_hash = excluded.description_hash,
                attributes_hash = excluded.attributes_hash, images_hash = excluded.images_hash,
                rich_content_hash = excluded.rich_content_hash, collected_at = excluded.collected_at
        """, (day, int(card["sku"]), card.get("name"), card.get("description"), attributes, images,
              _hash(card.get("name")), _hash(card.get("description")), _hash(card.get("attributes")),
              _hash(card.get("images")), _hash(card.get("rich_content")), stamp))
    return len(cards)


def latest_cards(conn: sqlite3.Connection) -> Dict[int, sqlite3.Row]:
    rows = conn.execute("""
        SELECT c.* FROM card_snapshots c
        JOIN (SELECT sku, MAX(date) AS date FROM card_snapshots GROUP BY sku) last USING (sku, date)
    """).fetchall()
    return {int(r["sku"]): r for r in rows}


# --- регионы и позиции -----------------------------------------------------------------------

def upsert_region(conn: sqlite3.Connection, code: str, name: str, **fields: Any) -> int:
    """Регион по коду (msk, spb…). Переданные поля (fias, profile_path, proxy_env, active) обновляются."""
    allowed = {k: v for k, v in fields.items() if k in ("fias", "profile_path", "proxy_env", "active")}
    stamp = now_iso()
    conn.execute("""
        INSERT INTO regions (code, name, created_at, updated_at) VALUES (?, ?, ?, ?)
        ON CONFLICT(code) DO UPDATE SET name = excluded.name, updated_at = excluded.updated_at
    """, (code, name, stamp, stamp))
    for key, value in allowed.items():
        conn.execute(f"UPDATE regions SET {key} = ? WHERE code = ?", (value, code))
    return int(conn.execute("SELECT id FROM regions WHERE code = ?", (code,)).fetchone()[0])


def active_regions(conn: sqlite3.Connection) -> List[sqlite3.Row]:
    return conn.execute("SELECT * FROM regions WHERE active = 1 ORDER BY id").fetchall()


POSITION_FIELDS = ("position", "organic_position", "page", "is_ad", "found", "depth")


def upsert_positions(conn: sqlite3.Connection, day: str, region_id: int, rows: Sequence[Dict[str, Any]], source: str = "serp") -> int:
    """rows: {sku, keyword_id, position, organic_position, page, is_ad, found, depth}. Не найден — position NULL, found 0."""
    stamp = now_iso()
    for row in rows:
        conn.execute(f"""
            INSERT INTO positions (date, measured_at, sku, keyword_id, region_id, {", ".join(POSITION_FIELDS)}, source)
            VALUES (?, ?, ?, ?, ?, {", ".join("?" * len(POSITION_FIELDS))}, ?)
            ON CONFLICT(date, sku, keyword_id, region_id, source) DO UPDATE SET
                measured_at = excluded.measured_at, {", ".join(f"{f} = excluded.{f}" for f in POSITION_FIELDS)}
        """, [day, stamp, int(row["sku"]), int(row["keyword_id"]), int(region_id),
              *(row.get(f) for f in POSITION_FIELDS), source])
    return len(rows)


# --- журнал сбора -----------------------------------------------------------------------------

def start_run(conn: sqlite3.Connection, source: str, day: Optional[str] = None) -> int:
    cursor = conn.execute("INSERT INTO collect_runs (source, day, started_at, status) VALUES (?, ?, ?, 'running')",
                          (source, day, now_iso()))
    conn.commit()
    return int(cursor.lastrowid)


def finish_run(conn: sqlite3.Connection, run_id: int, status: str, message: str = "") -> None:
    conn.execute("UPDATE collect_runs SET status = ?, message = ?, finished_at = ? WHERE id = ?",
                 (status, message[:2000], now_iso(), run_id))
    conn.commit()


def collected_days(conn: sqlite3.Connection, source: str) -> set:
    return {r[0] for r in conn.execute("SELECT DISTINCT day FROM collect_runs WHERE source = ? AND status = 'ok'", (source,))}


def runs_started_since(conn: sqlite3.Connection, source: str, since_iso: str) -> List[sqlite3.Row]:
    return conn.execute("SELECT * FROM collect_runs WHERE source = ? AND started_at >= ? ORDER BY started_at",
                        (source, since_iso)).fetchall()


def last_run(conn: sqlite3.Connection, source: Optional[str] = None) -> Optional[sqlite3.Row]:
    if source:
        return conn.execute("SELECT * FROM collect_runs WHERE source = ? ORDER BY id DESC LIMIT 1", (source,)).fetchone()
    return conn.execute("SELECT * FROM collect_runs ORDER BY id DESC LIMIT 1").fetchone()


def collected_today(path: Optional[Path] = None) -> bool:
    """Успешно ли сегодня (по UTC) отработал сборщик — по журналу загрузки товаров. Базы нет — False."""
    path = Path(path or DB_PATH)
    if not path.exists():
        return False
    with closing(connect(path)) as conn:
        row = conn.execute("SELECT 1 FROM collect_runs WHERE source = 'products' AND status = 'ok' AND day = ?",
                           (datetime.now(timezone.utc).date().isoformat(),)).fetchone()
    return row is not None


def main(argv: Optional[List[str]] = None) -> None:
    parser = argparse.ArgumentParser(description="Миграции базы раздела «Поиск».")
    sub = parser.add_subparsers(dest="command", required=True)
    m = sub.add_parser("migrate")
    m.add_argument("--to", type=int, default=None)
    sub.add_parser("version")
    args = parser.parse_args(argv)
    with closing(connect(migrate_to_latest=False)) as conn:
        if args.command == "migrate":
            print(f"Версия схемы: {migrate(conn, args.to)}")
        else:
            print(f"Версия схемы: {current_version(conn)} (последняя — {LATEST_VERSION})")


if __name__ == "__main__":
    main()
