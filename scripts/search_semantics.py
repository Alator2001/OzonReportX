"""
Семантика раздела «Поиск и ключевые слова»: леммы запросов, частотность, кластеры, покрытие карточки, приоритет.

Всё считается локально по данным `search data/search.db` (сеть не нужна) и пересчитывается сборщиком
после каждой загрузки. Подробности и формулы — docs/dev/17-search.md.

    python scripts/search_semantics.py rebuild                     -> леммы, частотность, кластеры, приоритеты
    python scripts/search_semantics.py add "доска из кедра" --offer 1302 [--frequency 500] [--priority 1]
    python scripts/search_semantics.py import запросы.csv
    python scripts/search_semantics.py coverage --offer 1302
    python scripts/search_semantics.py growth
"""
import argparse
import csv
import json
import re
import statistics
import sys
from bisect import bisect_left
from contextlib import closing
from functools import lru_cache
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set, Tuple

try:
    from scripts import search_db as db
except ModuleNotFoundError:
    import search_db as db

SETTINGS_PATH = db.ROOT / "search_settings.json"

DEFAULT_SETTINGS: Dict[str, Any] = {
    # Запрос входит в кластер, если доля общих лемм с главным запросом кластера (Жаккар) не меньше порога.
    "cluster_jaccard": 0.75,
    # За сколько последних дней считать частотность и конверсию товара.
    "frequency_days": 28,
    # «Зона роста»: позиции, с которых реально подняться в топ, и минимальная частотность (доля от медианы по товару).
    "growth_zone": [10, 50],
    "growth_min_frequency_ratio": 1.0,
    # Позиция, которую подставляем в расчёт, пока она неизвестна (позиции придут из парсера выдачи).
    "unknown_position": 30,
    # Ожидаемая доля кликов по позиции в выдаче; между точками — линейно, после последней — последнее значение.
    "ctr_curve": {"1": 0.10, "2": 0.08, "3": 0.07, "4": 0.06, "5": 0.055, "10": 0.04, "20": 0.025,
                  "36": 0.015, "50": 0.01, "100": 0.004, "200": 0.001},
    # Парсер выдачи (scripts/serp_collect.py). Объём осторожный: Ozon показывает капчу уже на ~5 запросах в минуту.
    "serp": {
        "max_queries": 15,             # запросов за прогон: сначала целевые, затем (fill_with_top) лучшие по приоритету
        "fill_with_top": True,
        "max_pages": 10,               # глубина: порций выдачи на запрос…
        "max_positions": 300,          # …и позиций — что наступит раньше
        "query_delay_seconds": [20, 60],
        "page_delay_seconds": [3, 8],
        "region_delay_seconds": [60, 120],
        "median_days": 3,              # позиция для приоритета — медиана последних замеров за столько дней
        "cooldown_hours": 24,          # после капчи парсер не запускается столько часов
        "headless": True,
        "timeout_seconds": 45,
        "keep_raw_days": 7,            # сколько дней хранить сырые ответы выдачи
    },
}

# Части речи, которые не несут смысла запроса: предлоги, союзы, частицы, междометия.
SKIP_POS = {"PREP", "CONJ", "PRCL", "INTJ"}
# Единицы измерения: pymorphy читает их как сокращения слов («см» → «смотреть»), поэтому оставляем как есть.
UNITS = {"мм", "см", "дм", "м", "км", "г", "гр", "кг", "мл", "л", "шт", "уп", "вт", "квт"}
TOKEN_RE = re.compile(r"[a-zа-яё]+|\d+(?:[.,]\d+)?", re.IGNORECASE)


def load_settings(path: Optional[Path] = None) -> Dict[str, Any]:
    """Настройки по умолчанию, поверх — `search_settings.json`, если он есть. Повреждённый файл — ошибка, а не тихий откат."""
    settings = json.loads(json.dumps(DEFAULT_SETTINGS))
    path = Path(path or SETTINGS_PATH)
    if path.exists():
        data = json.loads(path.read_text(encoding="utf-8"))
        for key, value in data.items():
            if key not in DEFAULT_SETTINGS:
                continue
            # Вложенные настройки (serp) дополняются, а не заменяются целиком.
            settings[key] = {**settings[key], **value} if isinstance(settings[key], dict) and key != "ctr_curve" else value
    return settings


# --- леммы ------------------------------------------------------------------------------------

@lru_cache(maxsize=1)
def _morph():
    import pymorphy3  # импорт здесь: без зависимости сбор данных работает, не считается только семантика
    return pymorphy3.MorphAnalyzer()


@lru_cache(maxsize=200_000)
def _lemma(token: str) -> Optional[str]:
    if token[0].isdigit():
        return token.replace(",", ".")
    if token in UNITS:
        return token
    parse = _morph().parse(token)[0]
    if parse.tag.POS in SKIP_POS:
        return None
    return parse.normal_form.replace("ё", "е")


def lemmas(text: str) -> List[str]:
    """Леммы слов текста по порядку, без предлогов, союзов и частиц: «Доски из кедра» → [доска, кедр]."""
    result = []
    for token in TOKEN_RE.findall(str(text or "").lower()):
        lemma = _lemma(token)
        if lemma:
            result.append(lemma)
    return result


def lemma_key(text: str) -> str:
    """Ключ смысла: множество лемм по алфавиту. «разделочная доска» и «доски разделочные» дают один ключ."""
    return " ".join(sorted(set(lemmas(text))))


def jaccard(a: Set[str], b: Set[str]) -> float:
    return len(a & b) / len(a | b) if a or b else 0.0


def update_lemmas(conn) -> int:
    """Пересчитывает ключи всех запросов: после изменения правил лемм старые ключи не должны остаться. Возвращает число изменённых."""
    rows = conn.execute("SELECT id, query, lemma_key FROM keywords").fetchall()
    changed = [(key, r["id"]) for r in rows if (key := lemma_key(r["query"])) != r["lemma_key"]]
    conn.executemany("UPDATE keywords SET lemma_key = ? WHERE id = ?", changed)
    return len(changed)


# --- частотность ------------------------------------------------------------------------------

def update_frequency(conn, days: int) -> Optional[Tuple[str, str]]:
    """
    Частотность запроса за последние days дней данных: сумма по дням максимума searched_users по товарам.
    Сумма по товарам завысила бы частотность: один покупатель видит несколько товаров магазина. Максимум — оценка снизу.
    Запросы без данных Ozon сохраняют частотность из ручного ввода или CSV.
    """
    last = conn.execute("SELECT MAX(date) FROM query_stats").fetchone()[0]
    if not last:
        return None
    first = conn.execute("SELECT date(?, ?)", (last, f"-{days - 1} days")).fetchone()[0]
    conn.execute("""
        UPDATE keywords SET frequency = f.value, frequency_source = 'api'
        FROM (
            SELECT keyword_id, SUM(day_max) AS value FROM (
                SELECT date, keyword_id, MAX(COALESCE(searched_users, 0)) AS day_max
                FROM query_stats WHERE date BETWEEN ? AND ? GROUP BY date, keyword_id
            ) GROUP BY keyword_id
        ) AS f
        WHERE keywords.id = f.keyword_id
    """, (first, last))
    return first, last


# --- кластеры ---------------------------------------------------------------------------------

def build_clusters(groups: Sequence[Tuple[str, float]], threshold: float) -> Dict[str, str]:
    """
    groups — (lemma_key, частотность группы). Возвращает {lemma_key: lemma_key главной группы кластера}.

    Жадно, от частых к редким: группа присоединяется к кластеру, с главной группой которого у неё наибольший
    Жаккар не ниже порога, иначе открывает свой кластер. Сравнение только с главной группой (без цепочек),
    поэтому кластер не «расползается»: «доска разделочная» и «доска разделочная деревянная» при пороге 0,75
    остаются разными кластерами — у них разные слова, которые нужно покрыть в карточке.
    """
    heads: List[Tuple[str, Set[str]]] = []
    by_lemma: Dict[str, List[int]] = {}
    result: Dict[str, str] = {}
    for key, _ in sorted(groups, key=lambda g: (-g[1], g[0])):
        words = set(key.split())
        candidates = {i for word in words for i in by_lemma.get(word, [])}
        best, best_score = None, 0.0
        for i in sorted(candidates):
            score = jaccard(words, heads[i][1])
            if score > best_score:
                best, best_score = i, score
        if best is not None and best_score >= threshold:
            result[key] = heads[best][0]
            continue
        heads.append((key, words))
        for word in words:
            by_lemma.setdefault(word, []).append(len(heads) - 1)
        result[key] = key
    return result


def update_clusters(conn, threshold: float) -> int:
    """Пересобирает кластеры. Запросы с cluster_locked (кластер выбран вручную) не трогаются. id кластера стабилен по главной группе."""
    groups = conn.execute("""
        SELECT lemma_key, SUM(COALESCE(frequency, 0)) AS freq FROM keywords
        WHERE lemma_key IS NOT NULL AND lemma_key != '' AND cluster_locked = 0 GROUP BY lemma_key
    """).fetchall()
    mapping = build_clusters([(g["lemma_key"], g["freq"]) for g in groups], threshold)
    stamp = db.now_iso()
    for head in sorted(set(mapping.values())):
        name = conn.execute("""
            SELECT query FROM keywords WHERE lemma_key = ? ORDER BY COALESCE(frequency, 0) DESC, length(query), query LIMIT 1
        """, (head,)).fetchone()["query"]
        conn.execute("""
            INSERT INTO clusters (lemma_key, name, created_at) VALUES (?, ?, ?)
            ON CONFLICT(lemma_key) DO UPDATE SET name = excluded.name
        """, (head, name, stamp))
    ids = {r["lemma_key"]: r["id"] for r in conn.execute("SELECT id, lemma_key FROM clusters")}
    conn.executemany("UPDATE keywords SET cluster_id = ? WHERE lemma_key = ? AND cluster_locked = 0",
                     [(ids[head], key) for key, head in mapping.items()])
    conn.execute("UPDATE keywords SET cluster_id = NULL WHERE (lemma_key IS NULL OR lemma_key = '') AND cluster_locked = 0")
    conn.execute("DELETE FROM clusters WHERE id NOT IN (SELECT cluster_id FROM keywords WHERE cluster_id IS NOT NULL)")
    return len(set(mapping.values()))


def lock_cluster(conn, keyword_id: int, cluster_id: Optional[int]) -> None:
    """Ручной перенос запроса в кластер (None — снять ручную привязку, кластер посчитается заново)."""
    conn.execute("UPDATE keywords SET cluster_id = ?, cluster_locked = ? WHERE id = ?",
                 (cluster_id, 0 if cluster_id is None else 1, keyword_id))


# --- покрытие карточки ------------------------------------------------------------------------

def card_lemmas(card) -> Dict[str, Set[str]]:
    attributes = json.loads(card["attributes"] or "{}")
    return {
        "title": set(lemmas(card["name"] or "")),
        "description": set(lemmas(card["description"] or "")),
        "attributes": set(lemmas(" ".join(v for values in attributes.values() for v in values))),
    }


def coverage_matrix(conn, sku: int) -> List[Dict[str, Any]]:
    """
    Для каждого кластера запросов товара: есть ли все слова главного запроса кластера в названии, описании,
    характеристиках, и каких слов нет нигде. Сортировка — по частотности кластера для этого товара.
    """
    card = db.latest_cards(conn).get(int(sku))
    fields = card_lemmas(card) if card else {"title": set(), "description": set(), "attributes": set()}
    everywhere = set().union(*fields.values())
    rows = conn.execute("""
        SELECT c.id, c.name, c.lemma_key, SUM(COALESCE(k.frequency, 0)) AS frequency, COUNT(*) AS queries,
               MAX(sk.is_target) AS is_target
        FROM sku_keywords sk JOIN keywords k ON k.id = sk.keyword_id JOIN clusters c ON c.id = k.cluster_id
        WHERE sk.sku = ? GROUP BY c.id ORDER BY frequency DESC, c.name
    """, (int(sku),)).fetchall()
    matrix = []
    for r in rows:
        words = set(r["lemma_key"].split())
        matrix.append({
            "cluster_id": r["id"], "cluster": r["name"], "words": sorted(words), "frequency": r["frequency"], "queries": r["queries"],
            "is_target": bool(r["is_target"]),
            "in_title": words <= fields["title"], "in_description": words <= fields["description"],
            "in_attributes": words <= fields["attributes"], "missing": sorted(words - everywhere),
            # Доля слов кластера в каждом поле: частичное покрытие видно, даже когда поле не содержит всех слов.
            "shares": {name: round(len(words & found) / len(words), 2) for name, found in fields.items()},
            "has_card": card is not None,
        })
    return matrix


# --- приоритет --------------------------------------------------------------------------------

def ctr_at(position: Optional[float], curve: Dict[str, float]) -> float:
    """Ожидаемый CTR позиции по кривой из настроек. Не найден в выдаче (None) — 0."""
    if position is None:
        return 0.0
    points = sorted((float(k), float(v)) for k, v in curve.items())
    xs = [p[0] for p in points]
    if position <= xs[0]:
        return points[0][1]
    if position >= xs[-1]:
        return points[-1][1]
    i = bisect_left(xs, position)
    (x0, y0), (x1, y1) = points[i - 1], points[i]
    return y0 + (y1 - y0) * (position - x0) / (x1 - x0)


def serp_positions(conn, days: int, region_id: Optional[int] = None) -> Dict[Tuple[int, int], Dict[str, Any]]:
    """
    Позиции из парсера выдачи за последние days дней замеров (по всем регионам или одному): медиана по замерам.
    «Не найден» считается позицией хуже любой, медиана — верхняя (statistics.median_high): при двух замерах
    «3 и не найден» итог — «не найден». Осторожная оценка: одна удачная выдача не делает запрос «в топе».
    Возвращает {(sku, keyword_id): {position (None — не найден), measurements, found}}.
    """
    last = conn.execute("SELECT MAX(date) FROM positions WHERE source = 'serp'").fetchone()[0]
    if not last:
        return {}
    query = "SELECT sku, keyword_id, position, found FROM positions WHERE source = 'serp' AND date > date(?, ?)"
    params: List[Any] = [last, f"-{days} days"]
    if region_id is not None:
        query += " AND region_id = ?"
        params.append(region_id)
    values: Dict[Tuple[int, int], List[float]] = {}
    for r in conn.execute(query, params):
        values.setdefault((r["sku"], r["keyword_id"]), []).append(float(r["position"]) if r["found"] else float("inf"))
    result = {}
    for key, items in values.items():
        median = statistics.median_high(items)
        result[key] = {"position": None if median == float("inf") else median, "measurements": len(items),
                       "found": sum(1 for v in items if v != float("inf"))}
    return result


def latest_positions(conn, serp_days: int = 3) -> Dict[Tuple[int, int], Optional[float]]:
    """
    Текущая позиция пары (sku, keyword_id): медиана выдачи за serp_days дней; если замеров выдачи нет — последняя
    позиция из API (только с Premium). Значение None — товар искали и не нашли; пары нет в словаре — позиция неизвестна.
    """
    rows = conn.execute("""
        SELECT q.sku, q.keyword_id, q.position FROM query_stats q
        JOIN (SELECT sku, keyword_id, MAX(date) AS date FROM query_stats WHERE position IS NOT NULL GROUP BY sku, keyword_id) last
          USING (sku, keyword_id, date)
    """).fetchall()
    result: Dict[Tuple[int, int], Optional[float]] = {(r["sku"], r["keyword_id"]): float(r["position"]) for r in rows}
    result.update({key: value["position"] for key, value in serp_positions(conn, serp_days).items()})
    return result


def conversion_factors(conn, days: int) -> Dict[int, float]:
    """
    Насколько товар лучше превращает поиск в заказы, чем средний товар магазина: (заказы / искавшие) товара,
    делённые на медиану по товарам с заказами. Ограничено 0,5–2, чтобы один товар не перекрывал частотность.
    """
    last = conn.execute("SELECT MAX(date) FROM sku_daily").fetchone()[0]
    if not last:
        return {}
    rows = conn.execute("""
        SELECT sku, SUM(COALESCE(ordered_units, 0)) AS orders, SUM(COALESCE(searched_users, 0)) AS searched
        FROM sku_daily WHERE date > date(?, ?) GROUP BY sku
    """, (last, f"-{days} days")).fetchall()
    rates = {r["sku"]: r["orders"] / r["searched"] for r in rows if r["searched"] > 0}
    positive = [v for v in rates.values() if v > 0]
    if not positive:
        return {}
    median = statistics.median(positive)
    return {sku: min(2.0, max(0.5, rate / median)) for sku, rate in rates.items()}


def score_pair(frequency: float, position: Optional[float], conversion: float, settings: Dict[str, Any],
               known: Optional[bool] = None) -> Dict[str, Any]:
    """
    Приоритет = частотность × прирост CTR при подъёме в топ-1 × коэффициент конверсии товара.
    Это оценка дополнительного трафика с поправкой на то, как товар продаёт: высокий приоритет — частый запрос,
    по которому товар далеко от топа, а сам товар хорошо конвертирует.

    known=False — позиция неизвестна (подставляется unknown_position); known=True и position=None — товар не найден
    в выдаче (CTR 0, прирост максимальный). По умолчанию known = position is not None.
    """
    curve = settings["ctr_curve"]
    known = position is not None if known is None else known
    if not known:
        current_ctr = ctr_at(float(settings["unknown_position"]), curve)
    else:
        current_ctr = ctr_at(position, curve)
    gain = max(0.0, ctr_at(1, curve) - current_ctr)
    return {"score": round(frequency * gain * conversion, 4), "frequency": frequency, "position": position,
            "position_known": known, "found": None if not known else position is not None,
            "ctr_gain": round(gain, 5), "conversion_factor": round(conversion, 3)}


def update_scores(conn, settings: Dict[str, Any]) -> int:
    positions = latest_positions(conn, int(settings["serp"]["median_days"]))
    conversion = conversion_factors(conn, int(settings["frequency_days"]))
    low, high = settings["growth_zone"]
    pairs = conn.execute("""
        SELECT sk.sku, sk.keyword_id, COALESCE(k.frequency, 0) AS frequency
        FROM sku_keywords sk JOIN keywords k ON k.id = sk.keyword_id
    """).fetchall()
    by_sku: Dict[int, List[float]] = {}
    for p in pairs:
        if p["frequency"] > 0:
            by_sku.setdefault(p["sku"], []).append(p["frequency"])
    medians = {sku: statistics.median(values) for sku, values in by_sku.items()}
    updates = []
    for p in pairs:
        key = (p["sku"], p["keyword_id"])
        position = positions.get(key)
        details = score_pair(p["frequency"], position, conversion.get(p["sku"], 1.0), settings, known=key in positions)
        frequent = p["frequency"] > 0 and p["frequency"] >= medians.get(p["sku"], 0) * float(settings["growth_min_frequency_ratio"])
        details["growth_zone"] = bool(position is not None and low <= position <= high and frequent)
        updates.append((details["score"], json.dumps(details, ensure_ascii=False), p["sku"], p["keyword_id"]))
    conn.executemany("UPDATE sku_keywords SET score = ?, score_details = ? WHERE sku = ? AND keyword_id = ?", updates)
    return len(updates)


def growth_zone(conn, limit: int = 100) -> List[Dict[str, Any]]:
    rows = conn.execute("""
        SELECT sk.sku, p.offer_id, k.query, sk.score, sk.score_details FROM sku_keywords sk
        JOIN keywords k ON k.id = sk.keyword_id LEFT JOIN products p ON p.sku = sk.sku
        WHERE json_extract(sk.score_details, '$.growth_zone') = 1 ORDER BY sk.score DESC LIMIT ?
    """, (limit,)).fetchall()
    return [{**dict(r), "score_details": json.loads(r["score_details"])} for r in rows]


def rebuild(conn, settings: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Полный пересчёт семантики одной транзакцией."""
    settings = settings or load_settings()
    with conn:
        lemmatized = update_lemmas(conn)
        window = update_frequency(conn, int(settings["frequency_days"]))
        clusters = update_clusters(conn, float(settings["cluster_jaccard"]))
        scored = update_scores(conn, settings)
    return {"lemmatized": lemmatized, "window": window, "clusters": clusters, "scored": scored}


# --- ручной ввод и CSV ------------------------------------------------------------------------

def resolve_skus(conn, skus: Iterable[Any] = (), offer_ids: Iterable[str] = ()) -> List[int]:
    result = [int(s) for s in skus]
    for offer_id in offer_ids:
        rows = conn.execute("SELECT sku FROM products WHERE offer_id = ?", (str(offer_id).strip(),)).fetchall()
        if not rows:
            raise ValueError(f"Артикул «{offer_id}» не найден среди товаров — сначала запустите сбор данных.")
        result += [r["sku"] for r in rows]
    return result


def add_keywords(conn, queries: Iterable[str], skus: Iterable[int] = (), source: str = "manual", target: bool = True,
                 frequency: Optional[float] = None, priority: Optional[int] = None) -> List[int]:
    """Добавляет запросы (и связи с товарами). Частотность из ввода не затирает данные Ozon."""
    ids = []
    for query in queries:
        keyword_id = db.ensure_keyword(conn, query, source)
        conn.execute("UPDATE keywords SET lemma_key = COALESCE(lemma_key, ?) WHERE id = ?", (lemma_key(query), keyword_id))
        if frequency is not None:
            conn.execute("""UPDATE keywords SET frequency = ?, frequency_source = ?
                            WHERE id = ? AND COALESCE(frequency_source, '') != 'api'""", (float(frequency), source, keyword_id))
        if priority is not None:
            conn.execute("UPDATE keywords SET priority = ? WHERE id = ?", (int(priority), keyword_id))
        for sku in skus:
            db.link_sku_keyword(conn, sku, keyword_id, source)
            if target:
                conn.execute("UPDATE sku_keywords SET is_target = 1 WHERE sku = ? AND keyword_id = ?", (int(sku), keyword_id))
            if priority is not None:
                conn.execute("UPDATE sku_keywords SET manual_priority = ? WHERE sku = ? AND keyword_id = ?", (int(priority), int(sku), keyword_id))
        ids.append(keyword_id)
    return ids


CSV_COLUMNS = {
    "query": {"query", "запрос", "поисковый запрос", "ключ", "ключевое слово", "keyword"},
    "sku": {"sku", "ску"},
    "offer_id": {"offer_id", "артикул"},
    "frequency": {"frequency", "частотность", "частота"},
    "priority": {"priority", "приоритет"},
}


def _read_csv_text(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "cp1251"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"{path.name}: не удалось прочитать файл — сохраните его в UTF-8.")


def _number(value: str) -> Optional[float]:
    value = (value or "").strip().replace("\xa0", "").replace(" ", "").replace(",", ".")
    return float(value) if value else None


def parse_csv(text: str) -> List[Dict[str, Any]]:
    """Строки CSV → [{query, sku, offer_id, frequency, priority}]. Разделитель ; , или табуляция, заголовки по-русски или по-английски."""
    dialect = csv.Sniffer().sniff(text.splitlines()[0] if text.strip() else ",", delimiters=";,\t")
    reader = csv.reader(text.splitlines(), dialect)
    header = [h.strip().lower() for h in next(reader, [])]
    columns = {name: header.index(h) for name, aliases in CSV_COLUMNS.items() for h in header if h in aliases}
    if "query" not in columns:
        raise ValueError("В CSV нет колонки с запросом: назовите её «запрос» или «query».")
    rows = []
    for line_no, cells in enumerate(reader, start=2):
        get = lambda name: cells[columns[name]].strip() if name in columns and columns[name] < len(cells) else ""
        if not get("query"):
            continue
        try:
            rows.append({"query": get("query"), "sku": get("sku") or None, "offer_id": get("offer_id") or None,
                         "frequency": _number(get("frequency")), "priority": _number(get("priority"))})
        except ValueError:
            raise ValueError(f"Строка {line_no}: частотность и приоритет должны быть числами.") from None
    return rows


def import_csv(conn, path: Path) -> int:
    rows = parse_csv(_read_csv_text(Path(path)))
    with conn:
        for row in rows:
            skus = resolve_skus(conn, [row["sku"]] if row["sku"] else [], [row["offer_id"]] if row["offer_id"] else [])
            priority = int(row["priority"]) if row["priority"] is not None else None
            add_keywords(conn, [row["query"]], skus, source="csv", target=bool(skus), frequency=row["frequency"], priority=priority)
    return len(rows)


# --- CLI --------------------------------------------------------------------------------------

def _print_coverage(conn, sku: int) -> None:
    matrix = coverage_matrix(conn, sku)
    if matrix and not matrix[0]["has_card"]:
        print("⚠️ Снимка карточки ещё нет — запустите scripts/search_collect.py.")
    mark = lambda flag: "✅" if flag else "—"
    print(f"{'Кластер':48} {'Част.':>9}  Назв  Опис  Хар-ки  Нет слов")
    for row in matrix:
        print(f"{row['cluster'][:48]:48} {row['frequency']:>9.0f}  {mark(row['in_title']):4}  {mark(row['in_description']):4}  "
              f"{mark(row['in_attributes']):6}  {', '.join(row['missing'])}")


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Семантика поисковых запросов.")
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("rebuild")
    add = sub.add_parser("add")
    add.add_argument("queries", nargs="+")
    add.add_argument("--sku", type=int, action="append", default=[])
    add.add_argument("--offer", action="append", default=[])
    add.add_argument("--frequency", type=float)
    add.add_argument("--priority", type=int)
    imp = sub.add_parser("import")
    imp.add_argument("path", type=Path)
    cov = sub.add_parser("coverage")
    cov.add_argument("--sku", type=int)
    cov.add_argument("--offer")
    sub.add_parser("growth")
    args = parser.parse_args(argv)
    with closing(db.connect()) as conn:
        if args.command == "add":
            with conn:
                ids = add_keywords(conn, args.queries, resolve_skus(conn, args.sku, args.offer),
                                   frequency=args.frequency, priority=args.priority)
            print(f"Добавлено запросов: {len(ids)}")
        elif args.command == "import":
            print(f"Импортировано строк: {import_csv(conn, args.path)}")
        elif args.command == "coverage":
            for sku in resolve_skus(conn, [args.sku] if args.sku else [], [args.offer] if args.offer else []):
                _print_coverage(conn, sku)
            return 0
        elif args.command == "growth":
            rows = growth_zone(conn)
            if not rows:
                print("Зона роста пуста: позиции ещё неизвестны (без подписки они появятся после включения парсера выдачи).")
            for r in rows:
                print(f"{r['offer_id'] or r['sku']:>10}  поз. {r['score_details']['position']:>5.0f}  {r['query']}")
            return 0
        result = rebuild(conn)
        print(f"Семантика пересчитана: изменённых ключей лемм — {result['lemmatized']}, кластеров — {result['clusters']}, "
              f"пар «товар — запрос» — {result['scored']}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
