"""
Данные для экрана «Поиск» в приложении: всё считается из `search data/search.db` без сети.

Окна считаются от последнего дня с данными запросов D (Ozon без подписки отдаёт день с задержкой в 2 дня):
«7 дней» = [D−6, D], «прошлые 7 дней» = [D−13, D−7], «28 дней» = [D−27, D]. Продажи берутся за те же даты,
чтобы спрос и заказы сравнивались честно.

Спрос по запросу за день — максимум searched_users по товарам (один покупатель видит несколько товаров магазина),
спрос магазина — сумма по запросам. Спрос товара по запросу — searched_users этого товара.
"""
import json
import os
import statistics
from contextlib import closing
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

try:
    from scripts import search_db as db
    from scripts import search_semantics as semantics
except ModuleNotFoundError:
    import search_db as db
    import search_semantics as semantics

# Пороги подсказок «что сделать»: ниже — шум, а не сигнал.
RISING_MIN_SEARCHES = 200        # запрос растёт: не меньше стольких поисков за 7 дней…
RISING_MIN_GROWTH = 0.5          # …и на 50% больше прошлой недели
DROP_MIN_PREVIOUS = 500          # товар стали реже находить: было от 500 искавших за неделю…
DROP_MAX_CHANGE = -0.3           # …и стало на 30% меньше
GAP_MIN_DEMAND = 100             # пробел в карточке: спрос товара по кластеру за 28 дней не меньше
GAP_MIN_TITLE_SHARE = 0.5        # слова кластера хотя бы наполовину есть в НАЗВАНИИ — иначе кластер, скорее всего, не про этот товар
NO_SALES_MIN_DEMAND = 20000      # «находят, но не заказывают»: столько искавших за 28 дней и ни одного заказа
MISSING_WEIGHT = 1 / 3            # вес пробела «слов нет нигде» против «есть в описании, но не в названии»
TREND_MIN_DAYS = 4               # тренд запроса — только если он был в данных хотя бы 4 дня из 7 в обеих неделях
DIAGNOSIS_MIN_ORDERS = 10        # вывод о конверсии магазина — только если на прошлой неделе было от 10 заказов
ACTIONS_LIMIT = 12


def plural(n: int, one: str, few: str, many: str) -> str:
    """1 товар, 2 товара, 5 товаров."""
    n = abs(int(n))
    if n % 10 == 1 and n % 100 != 11:
        return one
    return few if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14 else many


def num(value: float) -> str:
    """91734 → «91 734» (неразрывный пробел между разрядами, как в остальной программе)."""
    return f"{value:,.0f}".replace(",", "\u00a0")


def goods(n: int) -> str:
    return f"{n} {plural(n, 'товар', 'товара', 'товаров')}"


def windows(last: date) -> Dict[str, Tuple[str, str]]:
    iso = lambda d: d.isoformat()  # noqa: E731
    return {"7": (iso(last - timedelta(days=6)), iso(last)), "prev7": (iso(last - timedelta(days=13)), iso(last - timedelta(days=7))),
            "28": (iso(last - timedelta(days=27)), iso(last))}


def change(current: float, previous: float) -> Optional[float]:
    """Относительное изменение; без базы сравнения — None (а не бесконечность)."""
    if not previous:
        return None
    return (current - previous) / previous


def _last_date(conn) -> Optional[date]:
    value = conn.execute("SELECT MAX(date) FROM query_stats").fetchone()[0]
    return date.fromisoformat(value) if value else None


# --- магазин ----------------------------------------------------------------------------------

def store_demand(conn, start: str, end: str) -> Dict[str, float]:
    row = conn.execute("""
        SELECT COALESCE(SUM(day_max), 0) AS searches, COUNT(DISTINCT keyword_id) AS queries FROM (
            SELECT date, keyword_id, MAX(COALESCE(searched_users, 0)) AS day_max FROM query_stats
            WHERE date BETWEEN ? AND ? GROUP BY date, keyword_id HAVING day_max > 0)
    """, (start, end)).fetchone()
    sales = conn.execute("SELECT COALESCE(SUM(ordered_units), 0), COALESCE(SUM(revenue), 0) FROM sku_daily WHERE date BETWEEN ? AND ?",
                         (start, end)).fetchone()
    return {"searches": float(row["searches"]), "queries": int(row["queries"]), "orders": float(sales[0]), "revenue": float(sales[1])}


def kpis(conn, w: Dict[str, Tuple[str, str]]) -> Dict[str, Any]:
    now, prev = store_demand(conn, *w["7"]), store_demand(conn, *w["prev7"])
    per_1000 = lambda d: d["orders"] / d["searches"] * 1000 if d["searches"] else None  # noqa: E731
    result = {}
    for key in ("searches", "queries", "orders", "revenue"):
        result[key] = {"value": now[key], "previous": prev[key], "change": change(now[key], prev[key])}
    current, previous = per_1000(now), per_1000(prev)
    result["per_1000"] = {"value": current, "previous": previous,
                          "change": change(current, previous) if current is not None and previous else None}
    return result


def daily(conn, w: Dict[str, Tuple[str, str]], sku: Optional[int] = None) -> List[Dict[str, Any]]:
    """По дням 28-дневного окна: спрос (магазина или товара) и заказы. Дни без данных — нули, чтобы ось была ровной."""
    start, end = w["28"]
    if sku is None:
        demand = dict(conn.execute("""
            SELECT date, SUM(day_max) FROM (SELECT date, keyword_id, MAX(COALESCE(searched_users, 0)) AS day_max
            FROM query_stats WHERE date BETWEEN ? AND ? GROUP BY date, keyword_id) GROUP BY date
        """, (start, end)).fetchall())
        orders = dict(conn.execute("SELECT date, SUM(COALESCE(ordered_units, 0)) FROM sku_daily WHERE date BETWEEN ? AND ? GROUP BY date",
                                   (start, end)).fetchall())
    else:
        rows = conn.execute("SELECT date, searched_users, ordered_units FROM sku_daily WHERE sku = ? AND date BETWEEN ? AND ?",
                            (sku, start, end)).fetchall()
        demand = {r["date"]: r["searched_users"] or 0 for r in rows}
        orders = {r["date"]: r["ordered_units"] or 0 for r in rows}
    day, last = date.fromisoformat(start), date.fromisoformat(end)
    result = []
    while day <= last:
        key = day.isoformat()
        result.append({"date": key, "searches": float(demand.get(key) or 0), "orders": float(orders.get(key) or 0)})
        day += timedelta(days=1)
    return result


# --- кластеры и запросы -----------------------------------------------------------------------

def _keyword_demand(conn, start: str, end: str) -> Dict[int, float]:
    return dict(conn.execute("""
        SELECT keyword_id, SUM(day_max) FROM (SELECT date, keyword_id, MAX(COALESCE(searched_users, 0)) AS day_max
        FROM query_stats WHERE date BETWEEN ? AND ? GROUP BY date, keyword_id) GROUP BY keyword_id
    """, (start, end)).fetchall())


def stable_keywords(conn, w: Dict[str, Tuple[str, str]]) -> set:
    """
    Запросы, которые были в данных не меньше TREND_MIN_DAYS дней в каждой из двух последних недель. Ozon отдаёт
    по товару только 15 главных запросов в день, поэтому редкие запросы появляются урывками, и их «−100%» — пропуск
    в данных, а не падение спроса. Тренд считаем только по устойчивым запросам.
    """
    counts = lambda start, end: dict(conn.execute(  # noqa: E731
        "SELECT keyword_id, COUNT(DISTINCT date) FROM query_stats WHERE date BETWEEN ? AND ? GROUP BY keyword_id", (start, end)).fetchall())
    now, prev = counts(*w["7"]), counts(*w["prev7"])
    return {k for k, days in now.items() if days >= TREND_MIN_DAYS and prev.get(k, 0) >= TREND_MIN_DAYS}


def _best_positions(conn, median_days: int) -> Dict[int, float]:
    """Лучшая (минимальная) медианная позиция среди товаров магазина по каждому запросу."""
    best: Dict[int, float] = {}
    for (sku, keyword_id), value in semantics.serp_positions(conn, median_days).items():
        if value["position"] is not None:
            best[keyword_id] = min(best.get(keyword_id, value["position"]), value["position"])
    return best


def clusters(conn, w: Dict[str, Tuple[str, str]], median_days: int = 3) -> List[Dict[str, Any]]:
    """Семантическое ядро магазина: кластеры по спросу за 28 дней, с трендом недели и покрытием в названиях."""
    d28, d7, dprev = _keyword_demand(conn, *w["28"]), _keyword_demand(conn, *w["7"]), _keyword_demand(conn, *w["prev7"])
    best = _best_positions(conn, median_days)
    stable = stable_keywords(conn, w)
    rows = conn.execute("""
        SELECT c.id, c.name, c.lemma_key, k.id AS keyword_id, k.query, sk.sku, sk.is_target
        FROM clusters c JOIN keywords k ON k.cluster_id = c.id LEFT JOIN sku_keywords sk ON sk.keyword_id = k.id
    """).fetchall()
    titles = {sku: set(semantics.lemmas(card["name"] or "")) for sku, card in db.latest_cards(conn).items()}
    result: Dict[int, Dict[str, Any]] = {}
    for r in rows:
        entry = result.setdefault(r["id"], {"id": r["id"], "name": r["name"], "words": r["lemma_key"].split(), "keywords": set(),
                                            "skus": set(), "target": False, "demand_28": 0.0, "demand_7": 0.0, "demand_prev7": 0.0,
                                            "best_position": None})
        if r["keyword_id"] not in entry["keywords"]:
            entry["keywords"].add(r["keyword_id"])
            entry["demand_28"] += d28.get(r["keyword_id"], 0)
            if r["keyword_id"] in stable:
                entry["demand_7"] += d7.get(r["keyword_id"], 0)
                entry["demand_prev7"] += dprev.get(r["keyword_id"], 0)
            if r["keyword_id"] in best:
                entry["best_position"] = min(entry["best_position"] or best[r["keyword_id"]], best[r["keyword_id"]])
        if r["sku"] is not None:
            entry["skus"].add(r["sku"])
            entry["target"] = entry["target"] or bool(r["is_target"])
    out = []
    for entry in result.values():
        words = set(entry["words"])
        in_title = sum(1 for sku in entry["skus"] if sku in titles and words <= titles[sku])
        out.append({**{k: v for k, v in entry.items() if k not in ("keywords", "skus")},
                    "queries": len(entry["keywords"]), "products": len(entry["skus"]), "in_title": in_title,
                    "trend": change(entry["demand_7"], entry["demand_prev7"])})
    out.sort(key=lambda c: (-c["demand_28"], c["name"]))
    return out


def cluster_queries(conn, cluster_id: int, w: Dict[str, Tuple[str, str]]) -> List[Dict[str, Any]]:
    d28, d7, dprev = _keyword_demand(conn, *w["28"]), _keyword_demand(conn, *w["7"]), _keyword_demand(conn, *w["prev7"])
    stable = stable_keywords(conn, w)
    rows = conn.execute("""
        SELECT k.id, k.query, k.source, MAX(COALESCE(sk.is_target, 0)) AS target, GROUP_CONCAT(COALESCE(p.offer_id, sk.sku)) AS offers
        FROM keywords k LEFT JOIN sku_keywords sk ON sk.keyword_id = k.id LEFT JOIN products p ON p.sku = sk.sku
        WHERE k.cluster_id = ? GROUP BY k.id
    """, (cluster_id,)).fetchall()
    result = [{"id": r["id"], "query": r["query"], "source": r["source"], "target": bool(r["target"]),
               "offers": sorted(set(str(r["offers"]).split(","))) if r["offers"] else [],
               "demand_28": float(d28.get(r["id"], 0)),
               "trend": change(d7.get(r["id"], 0), dprev.get(r["id"], 0)) if r["id"] in stable else None} for r in rows]
    result.sort(key=lambda q: (-q["demand_28"], q["query"]))
    return result


def set_target(conn, keyword_id: int, target: bool) -> int:
    """Целевой запрос — для всех связанных с ним товаров. Целевые замеряются парсером в первую очередь."""
    with conn:
        return conn.execute("UPDATE sku_keywords SET is_target = ? WHERE keyword_id = ?", (int(target), keyword_id)).rowcount


# --- товары -----------------------------------------------------------------------------------

def _sku_demand(conn, start: str, end: str) -> Dict[int, float]:
    return dict(conn.execute("SELECT sku, SUM(COALESCE(searched_users, 0)) FROM sku_daily WHERE date BETWEEN ? AND ? GROUP BY sku",
                             (start, end)).fetchall())


def _sku_orders(conn, start: str, end: str) -> Dict[int, float]:
    return dict(conn.execute("SELECT sku, SUM(COALESCE(ordered_units, 0)) FROM sku_daily WHERE date BETWEEN ? AND ? GROUP BY sku",
                             (start, end)).fetchall())


def coverage(conn, sku: int, w: Dict[str, Tuple[str, str]], limit: int = 40) -> List[Dict[str, Any]]:
    """
    Матрица покрытия карточки, отсортированная по спросу ЭТОГО товара за 28 дней (а не общему спросу кластера):
    так наверху то, что реально приводит покупателей к товару.
    """
    start, end = w["28"]
    demand = dict(conn.execute("""
        SELECT k.cluster_id, SUM(COALESCE(q.searched_users, 0)) FROM query_stats q JOIN keywords k ON k.id = q.keyword_id
        WHERE q.sku = ? AND q.date BETWEEN ? AND ? GROUP BY k.cluster_id
    """, (sku, start, end)).fetchall())
    rows = []
    for row in semantics.coverage_matrix(conn, sku):
        rows.append({**row, "demand": float(demand.get(row["cluster_id"], 0))})
    rows.sort(key=lambda r: (-r["demand"], -r["frequency"], r["cluster"]))
    return rows[:limit]


def title_coverage(rows: List[Dict[str, Any]]) -> Optional[float]:
    """Доля спроса товара, чьи слова целиком есть в названии. Название Ozon учитывает в поиске сильнее всего."""
    total = sum(r["demand"] for r in rows)
    return sum(r["demand"] for r in rows if r["in_title"]) / total if total else None


def gaps(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Пробелы карточки: кластеры со спросом, слов которых нет в названии. Кластеры, у которых в названии меньше половины
    слов, пропускаем: скорее всего, Ozon показал товар по чужому запросу («стол» у разделочной доски).
    """
    result = []
    for r in rows:
        if r["in_title"] or r["demand"] < GAP_MIN_DEMAND:
            continue
        if r["shares"]["title"] < GAP_MIN_TITLE_SHARE:
            continue
        # Слов нет нигде в карточке — часто это другой товар («из нержавеющей стали» у деревянной доски).
        # Такой пробел показываем, только если название совпадает с кластером больше чем наполовину.
        if r["missing"] and r["shares"]["title"] <= GAP_MIN_TITLE_SHARE:
            continue
        result.append({**r, "kind": "missing" if r["missing"] else "not_in_title"})
    return result


def products(conn, w: Dict[str, Tuple[str, str]]) -> List[Dict[str, Any]]:
    d7, dprev, d28 = _sku_demand(conn, *w["7"]), _sku_demand(conn, *w["prev7"]), _sku_demand(conn, *w["28"])
    o7, o28 = _sku_orders(conn, *w["7"]), _sku_orders(conn, *w["28"])
    counts = dict(conn.execute("SELECT sku, COUNT(*) FROM sku_keywords GROUP BY sku").fetchall())
    result = []
    for p in conn.execute("SELECT sku, offer_id, name FROM products").fetchall():
        sku = p["sku"]
        if not d28.get(sku) and not o28.get(sku):
            continue
        rows = coverage(conn, sku, w)
        result.append({"sku": sku, "offer_id": p["offer_id"], "name": p["name"], "demand_7": float(d7.get(sku, 0)),
                       "demand_prev7": float(dprev.get(sku, 0)), "trend": change(d7.get(sku, 0), dprev.get(sku, 0)),
                       "orders_7": float(o7.get(sku, 0)), "orders_28": float(o28.get(sku, 0)),
                       "per_1000": float(o28.get(sku, 0)) / d28[sku] * 1000 if d28.get(sku) else None,
                       "queries": int(counts.get(sku, 0)), "title_coverage": title_coverage(rows), "gaps": len(gaps(rows)),
                       "no_sales": not o28.get(sku) and float(d28.get(sku, 0)) >= NO_SALES_MIN_DEMAND})
    result.sort(key=lambda p: (-p["demand_7"], -p["orders_28"], str(p["offer_id"])))
    return result


def card_changes(conn, sku: int) -> List[Dict[str, Any]]:
    """Что менялось в карточке по ежедневным снимкам: дата и части (название, описание, характеристики, фото, rich)."""
    parts = [("name_hash", "название"), ("description_hash", "описание"), ("attributes_hash", "характеристики"),
             ("images_hash", "фото"), ("rich_content_hash", "rich-контент")]
    rows = conn.execute("SELECT * FROM card_snapshots WHERE sku = ? ORDER BY date", (sku,)).fetchall()
    result = []
    for previous, current in zip(rows, rows[1:]):
        changed = [label for key, label in parts if previous[key] != current[key]]
        if changed:
            result.append({"date": current["date"], "parts": changed})
    return list(reversed(result))


def product_detail(conn, sku: int, w: Dict[str, Tuple[str, str]], median_days: int = 3) -> Dict[str, Any]:
    product = conn.execute("SELECT * FROM products WHERE sku = ?", (sku,)).fetchone()
    card = db.latest_cards(conn).get(sku)
    rows = coverage(conn, sku, w)
    start, end = w["28"]
    s7, sp = w["7"], w["prev7"]
    positions = semantics.serp_positions(conn, median_days)
    stable = stable_keywords(conn, w)
    queries = conn.execute("""
        SELECT k.id, k.query, sk.is_target,
               SUM(CASE WHEN q.date BETWEEN ? AND ? THEN COALESCE(q.searched_users, 0) ELSE 0 END) AS d28,
               SUM(CASE WHEN q.date BETWEEN ? AND ? THEN COALESCE(q.searched_users, 0) ELSE 0 END) AS d7,
               SUM(CASE WHEN q.date BETWEEN ? AND ? THEN COALESCE(q.searched_users, 0) ELSE 0 END) AS dprev
        FROM sku_keywords sk JOIN keywords k ON k.id = sk.keyword_id
        LEFT JOIN query_stats q ON q.keyword_id = k.id AND q.sku = sk.sku
        WHERE sk.sku = ? GROUP BY k.id ORDER BY d28 DESC, k.query LIMIT 30
    """, (start, end, *s7, *sp, sku)).fetchall()
    attributes = json.loads(card["attributes"] or "{}") if card else {}
    return {
        "sku": sku, "offer_id": product["offer_id"] if product else None, "name": (card["name"] if card else None) or (product["name"] if product else ""),
        "card": {"description_length": len(card["description"] or "") if card else 0, "attributes": len(attributes),
                 "images": len(json.loads(card["images"] or "[]")) if card else 0, "date": card["date"] if card else None},
        "daily": daily(conn, w, sku), "coverage": rows, "title_coverage": title_coverage(rows), "gaps": gaps(rows),
        "queries": [{"id": q["id"], "query": q["query"], "target": bool(q["is_target"]), "demand_28": float(q["d28"] or 0),
                     "trend": change(q["d7"] or 0, q["dprev"] or 0) if q["id"] in stable else None,
                     "position": positions.get((sku, q["id"]), {}).get("position"),
                     "measured": (sku, q["id"]) in positions} for q in queries],
        "changes": card_changes(conn, sku),
    }


# --- подсказки «что сделать» ------------------------------------------------------------------

def store_diagnosis(k: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """
    Главный вывод недели по магазину: что упало — спрос или конверсия. Если искать стали почти столько же,
    а заказов заметно меньше, проблема не в поиске, а в карточке: цена, остатки, сроки доставки, рейтинг.
    """
    searches, orders = k["searches"]["change"], k["orders"]["change"]
    if searches is None or orders is None or k["orders"]["previous"] < DIAGNOSIS_MIN_ORDERS:
        return None
    if orders <= -0.3 and searches > -0.15:
        return {"kind": "conversion", "weight": float("inf"),
                "text": f"Заказов из поиска меньше на {abs(orders) * 100:.0f}%, а спрос почти не изменился ({searches * 100:+.0f}%). "
                        "Товары находят, но не покупают: проверьте цену и акции, остатки на складах рядом с покупателем, "
                        "сроки доставки и свежие отзывы."}
    if orders <= -0.3 and searches <= -0.15:
        return {"kind": "demand", "weight": float("inf"),
                "text": f"Упали и спрос ({searches * 100:+.0f}%), и заказы ({orders * 100:+.0f}%). Если это не сезон, товары стали "
                        "реже показывать: проверьте остатки, индекс цен и позиции по главным запросам."}
    return None


def _relevant(query: str, title_words: set) -> bool:
    words = set(semantics.lemmas(query))
    return bool(words) and len(words & title_words) / len(words) >= GAP_MIN_TITLE_SHARE


def actions(conn, w: Dict[str, Tuple[str, str]], product_rows: List[Dict[str, Any]], k: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
    """
    Короткий список дел по важности (вес — сколько поисков затронуто): главный вывод недели, пробелы в названиях
    (один пункт на кластер со списком артикулов), растущие запросы про ваши товары, товары, которые стали реже находить.
    """
    items: List[Dict[str, Any]] = []
    diagnosis = store_diagnosis(k) if k else None
    if diagnosis:
        items.append({**diagnosis, "offers": []})
    grouped: Dict[Tuple[int, str], Dict[str, Any]] = {}
    drops: List[Dict[str, Any]] = []
    for p in product_rows:
        for gap in gaps(coverage(conn, p["sku"], w))[:3]:
            entry = grouped.setdefault((gap["cluster_id"], gap["kind"]), {**gap, "offers": [], "weight": 0.0})
            entry["offers"].append(str(p["offer_id"] or p["sku"]))
            # спрос за 28 дней → примерно за неделю; «нет нигде» — втрое ниже: слова могут быть не про этот товар
            entry["weight"] += gap["demand"] / 4 * (MISSING_WEIGHT if gap["kind"] == "missing" else 1.0)
        if p["trend"] is not None and p["trend"] <= DROP_MAX_CHANGE and p["demand_prev7"] >= DROP_MIN_PREVIOUS:
            drops.append(p)
    silent = [p for p in product_rows if p["no_sales"]]
    if silent:
        silent.sort(key=lambda p: -p["demand_7"])
        items.append({"kind": "no_sales", "offers": [str(p["offer_id"] or p["sku"]) for p in silent],
                      "weight": sum(p["demand_7"] for p in silent) * 2,
                      "text": f"{goods(len(silent)).capitalize()} ищут, но не заказывают: за 28 дней ни одного заказа, хотя по запросам "
                              f"с ними искали десятки тысяч раз (например, {silent[0]['offer_id']}: {num(silent[0]['demand_7'])} за неделю). "
                              "Либо товар далеко в выдаче (покажет парсер позиций), либо карточка не убеждает: цена, главное фото, "
                              "отзывы, срок доставки."})
    if drops:
        # Один пункт на все товары, которые стали реже находить: причины у них обычно общие (остатки, цена, акции).
        drops.sort(key=lambda p: p["demand_7"] - p["demand_prev7"])
        lost = sum(p["demand_prev7"] - p["demand_7"] for p in drops)
        worst = drops[0]
        text = (f"Товар {worst['offer_id']} стали реже находить: {num(worst['demand_7'])} искавших за неделю против {num(worst['demand_prev7'])}."
                if len(drops) == 1 else
                f"{goods(len(drops))} стали реже находить (на {abs(DROP_MAX_CHANGE) * 100:.0f}% и больше): минус {num(lost)} искавших за неделю. "
                f"Сильнее всего — {worst['offer_id']}: {num(worst['demand_7'])} против {num(worst['demand_prev7'])}.")
        items.append({"kind": "drop", "offers": [str(p["offer_id"] or p["sku"]) for p in drops], "weight": lost,
                      "text": text + " Проверьте остатки по кластерам, цену и участие в акциях."})
    for gap in grouped.values():
        many = len(gap["offers"]) > 1
        if gap["kind"] == "missing":
            text = (f"По кластеру «{gap['cluster']}» Ozon показывает " + (goods(len(gap["offers"])) if many else "товар")
                    + f", но слов «{', '.join(gap['missing'])}» в карточке нет. Если товар под это подходит — добавьте их; "
                    "если нет (другой материал, размер, назначение) — пропустите: чужие слова снижают конверсию.")
        else:
            text = (f"«{gap['cluster']}» есть в описании или характеристиках, но не в названии"
                    + (f" у {goods(len(gap['offers']))}" if many else "") + ". Ozon сильнее всего учитывает название.")
        items.append({"kind": "gap", "offers": gap["offers"], "weight": gap["weight"], "text": text, "cluster_id": gap["cluster_id"]})
    d7, dprev = _keyword_demand(conn, *w["7"]), _keyword_demand(conn, *w["prev7"])
    names = dict(conn.execute("SELECT id, query FROM keywords").fetchall())
    title_words: set = set()
    for card in db.latest_cards(conn).values():
        title_words |= set(semantics.lemmas(card["name"] or ""))
    stable = stable_keywords(conn, w)
    for keyword_id, value in d7.items():
        if keyword_id not in stable:
            continue
        growth = change(value, dprev.get(keyword_id, 0))
        if value >= RISING_MIN_SEARCHES and growth is not None and growth >= RISING_MIN_GROWTH and _relevant(names.get(keyword_id, ""), title_words):
            items.append({"kind": "rising", "offers": [], "weight": value - dprev.get(keyword_id, 0),
                          "text": f"Спрос растёт: «{names.get(keyword_id)}» — {num(value)} поисков за неделю ({growth * 100:+.0f}%). "
                                  "Проверьте, что товар под этот запрос в наличии и слова запроса есть в названии."})
    items.sort(key=lambda i: -i["weight"])
    return items[:ACTIONS_LIMIT]


# --- позиции ----------------------------------------------------------------------------------

def serp_status(conn, settings: Dict[str, Any]) -> Dict[str, Any]:
    from dotenv import load_dotenv
    load_dotenv(db.ROOT / ".env")
    try:
        import playwright  # noqa: F401
        installed = True
    except ImportError:
        installed = False
    regions = []
    for r in conn.execute("SELECT * FROM regions ORDER BY id").fetchall():
        regions.append({"code": r["code"], "name": r["name"], "fias": r["fias"], "active": bool(r["active"]),
                        "profile": bool(r["profile_path"] and (db.ROOT / r["profile_path"]).exists())})
    last = conn.execute("SELECT * FROM collect_runs WHERE source = 'serp' ORDER BY id DESC LIMIT 1").fetchone()
    blocked = conn.execute("SELECT started_at FROM collect_runs WHERE source = 'serp' AND status = 'blocked' ORDER BY id DESC LIMIT 1").fetchone()
    blocked_until = None
    if blocked:
        until = datetime.fromisoformat(blocked["started_at"]) + timedelta(hours=float(settings["serp"]["cooldown_hours"]))
        blocked_until = until.isoformat() if until > datetime.now(timezone.utc) else None
    return {"enabled": (os.getenv("SERP_PARSER_ENABLED") or "").strip().lower() in ("1", "true", "yes", "да"),
            "installed": installed, "regions": regions, "blocked_until": blocked_until,
            "last_run": dict(last) if last else None,
            "measured_days": conn.execute("SELECT COUNT(DISTINCT date) FROM positions WHERE source = 'serp'").fetchone()[0]}


def positions_table(conn, median_days: int = 3) -> List[Dict[str, Any]]:
    """Запрос × товар: медиана позиции по каждому региону, изменение к замеру неделю назад и разброс между регионами."""
    regions = conn.execute("SELECT id, name FROM regions ORDER BY id").fetchall()
    last = conn.execute("SELECT MAX(date) FROM positions WHERE source = 'serp'").fetchone()[0]
    if not last:
        return []
    week_ago = (date.fromisoformat(last) - timedelta(days=7)).isoformat()
    old = {(r["sku"], r["keyword_id"], r["region_id"]): r["position"] for r in conn.execute(
        "SELECT sku, keyword_id, region_id, position FROM positions WHERE source = 'serp' AND date = ?", (week_ago,))}
    rows: Dict[Tuple[int, int], Dict[str, Any]] = {}
    for region in regions:
        for (sku, keyword_id), value in semantics.serp_positions(conn, median_days, region_id=region["id"]).items():
            entry = rows.setdefault((sku, keyword_id), {"sku": sku, "keyword_id": keyword_id, "regions": {}})
            before = old.get((sku, keyword_id, region["id"]))
            entry["regions"][region["name"]] = {"position": value["position"], "measurements": value["measurements"],
                                                "change": (before - value["position"]) if before and value["position"] else None}
    names = dict(conn.execute("SELECT id, query FROM keywords").fetchall())
    offers = dict(conn.execute("SELECT sku, offer_id FROM products").fetchall())
    result = []
    for entry in rows.values():
        found = [v["position"] for v in entry["regions"].values() if v["position"] is not None]
        spread = (max(found) - min(found)) if len(found) > 1 else (None if not found or len(found) == len(entry["regions"]) else float("inf"))
        result.append({**entry, "query": names.get(entry["keyword_id"]), "offer_id": offers.get(entry["sku"]),
                       "best": min(found) if found else None, "spread": spread})
    result.sort(key=lambda r: (r["best"] is None, r["best"] or 0, r["query"] or ""))
    return result


# --- всё для экрана ---------------------------------------------------------------------------

def load(path: Optional[Path] = None) -> Dict[str, Any]:
    path = Path(path or db.DB_PATH)
    if not path.exists():
        return {"empty": True}
    settings = semantics.load_settings()
    with closing(db.connect(path)) as conn:
        last = _last_date(conn)
        if last is None:
            return {"empty": True}
        w = windows(last)
        product_rows = products(conn, w)
        store_kpis = kpis(conn, w)
        collected = conn.execute("SELECT finished_at FROM collect_runs WHERE source = 'products' AND status = 'ok' ORDER BY id DESC LIMIT 1").fetchone()
        return {
            "empty": False, "last_date": last.isoformat(), "windows": w,
            "collected_at": collected[0] if collected else None,
            "kpis": store_kpis, "daily": daily(conn, w), "actions": actions(conn, w, product_rows, store_kpis),
            "clusters": clusters(conn, w, int(settings["serp"]["median_days"])), "products": product_rows,
            "serp": serp_status(conn, settings), "positions": positions_table(conn, int(settings["serp"]["median_days"])),
        }


def load_product(sku: int, path: Optional[Path] = None) -> Dict[str, Any]:
    settings = semantics.load_settings()
    with closing(db.connect(Path(path or db.DB_PATH))) as conn:
        return product_detail(conn, sku, windows(_last_date(conn)), int(settings["serp"]["median_days"]))


def load_cluster(cluster_id: int, path: Optional[Path] = None) -> List[Dict[str, Any]]:
    with closing(db.connect(Path(path or db.DB_PATH))) as conn:
        return cluster_queries(conn, cluster_id, windows(_last_date(conn)))


def toggle_target(keyword_id: int, target: bool, path: Optional[Path] = None) -> int:
    with closing(db.connect(Path(path or db.DB_PATH))) as conn:
        changed = set_target(conn, keyword_id, target)
        semantics.rebuild(conn)
        return changed
