# -*- coding: utf-8 -*-
"""
Инструменты AI-ассистента поверх аналитической базы (`ai_warehouse`) и живого Ozon (только чтение).

Принцип: всё, что можно посчитать, считает код (SQL), модель получает готовые таблицы и объясняет их.
Каждый инструмент возвращает словарь: `title`, таблицы `tables` ({name: {"columns", "rows"}}),
`notes` (пояснения и предупреждения) или `error`. `render` превращает результат в компактный текст для модели.

Набор:
- describe_data — каталог таблиц базы (что есть, за какие периоды), колонки одной таблицы;
- run_sql — произвольный SELECT по базе (только чтение, лимит строк, таймаут);
- sales_metrics — готовые показатели продаж с разбивкой (месяц, артикул, схема, статус);
- month_metrics — официальные итоги месяцев из отчётов (чистая прибыль, реклама, комиссии…);
- store_overview — обзор магазина за месяц: итоги, лидеры, изменения, проблемы, остатки;
- product_profile — всё об одном артикуле из всех источников;
- product_table — все товары одной таблицей: класс, остаток, цены, продажи, прибыль, проблемы;
- supply_needs — что пополнить: поставка на FBO, нехватка на FBS, деньги по себестоимости;
- find_products — поиск артикула по фрагменту названия или артикула;
- live_prices, live_stocks, live_promotions, live_discount_requests — текущие данные Ozon, только чтение.
"""

import ast
import math
import operator
import re
import sqlite3
import time
from contextlib import closing
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

try:
    from scripts import ai_warehouse as warehouse
except ModuleNotFoundError:
    import ai_warehouse as warehouse

MONTHS_RU = warehouse.MONTHS_RU
SQL_ROW_LIMIT = 200
SQL_ROW_LIMIT_MAX = 1000
SQL_TIMEOUT_SECONDS = 10
# Не чаще одной пересборки базы за этот интервал: один вопрос — несколько вызовов инструментов.
REFRESH_INTERVAL_SECONDS = 30
# Живые данные Ozon кешируются: повтор того же вопроса не тратит лимиты API.
LIVE_CACHE_SECONDS = 300

MARKETPLACES = ("ozon", "wb")
SOLD_STATUSES = {"ozon": ("delivered", "returned"), "wb": ("продажа", "возврат")}
DELIVERED_STATUS = {"ozon": "delivered", "wb": "продажа"}
CANCELLED_STATUS = {"ozon": "cancelled", "wb": "отмена"}
RETURNED_STATUS = {"ozon": "returned", "wb": "возврат"}
ORDERS_TABLE = {"ozon": "ozon_orders", "wb": "wb_orders"}
METRICS_TABLE = {"ozon": "ozon_month_metrics", "wb": "wb_month_metrics"}

# Ключевые итоги месяца для обзора (названия — как в блоке P/Q отчёта).
KEY_MONTH_METRICS = {
    "ozon": [
        "Общая выручка", "Чистая прибыль", "Рентабельность по чистой прибыли (Net Profit Margin) %", "Валовая прибыль",
        "Общее количество заказов", "Количество доставленных заказов", "Количество отменённых заказов", "Средний чек",
        "Продвижение Ozon", "Звёздные товары", "Комиссии Ozon сумма", "Логистика сумма", "Расход хранения FBO",
        "Расходы Ozon вне заказов",
    ],
    "wb": ["Общая выручка", "Итог WB к выплате", "Средний чек продажи", "Продажи (уник. srid)", "Возвраты (уник. srid)",
           "Логистика", "Хранение", "Штрафы"],
}


# ---------- общие помощники ----------

def normalize_period(value: Any) -> Optional[str]:
    """«2026-09», «2026-9», «Сентябрь 2026», «сентябрь 2026» → «2026-09»; иначе None."""
    if value is None:
        return None
    text = str(value).strip()
    match = re.fullmatch(r"(\d{4})-(\d{1,2})", text)
    if match and 1 <= int(match.group(2)) <= 12:
        return f"{match.group(1)}-{int(match.group(2)):02d}"
    return warehouse.period_from_name(text)


def period_title(period: str) -> str:
    year, month = period.split("-")
    return f"{MONTHS_RU[int(month) - 1]} {year}"


def shift_period(period: str, months: int) -> str:
    year, month = (int(x) for x in period.split("-"))
    index = year * 12 + (month - 1) + months
    return f"{index // 12}-{index % 12 + 1:02d}"


def normalize_artikul(value: Any) -> str:
    text = str(value or "").strip()
    return text[:-2] if re.fullmatch(r"\d+\.0", text) else text


def _artikuls(value: Any) -> List[str]:
    if value is None or value == "":
        return []
    items = value if isinstance(value, (list, tuple)) else re.split(r"[,\s]+", str(value))
    return [a for a in (normalize_artikul(item) for item in items) if a]


def _round(value: Any) -> Any:
    if isinstance(value, float):
        if value.is_integer():
            return int(value)
        return round(value, 2)
    return value


def table(columns: Sequence[str], rows: Sequence[Sequence[Any]]) -> Dict[str, Any]:
    return {"columns": list(columns), "rows": [[_round(v) for v in row] for row in rows]}


def query_table(conn: sqlite3.Connection, sql: str, params: Sequence[Any] = ()) -> Dict[str, Any]:
    cursor = conn.execute(sql, params)
    columns = [d[0] for d in cursor.description or []]
    return table(columns, [tuple(row) for row in cursor.fetchall()])


def _change(current: Any, previous: Any, percent_metric: bool = False) -> Optional[str]:
    """Изменение к прошлому периоду: «-30.0%», для показателей в процентах — «+0.47 п.п.»."""
    if not isinstance(current, (int, float)) or not isinstance(previous, (int, float)):
        return None
    if percent_metric:
        return f"{current - previous:+.2f} п.п."
    if not previous:
        return None
    return f"{100.0 * (current - previous) / abs(previous):+.1f}%"


def _cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        value = _round(value)
    return str(value).replace("|", "/").replace("\n", " ↵ ")


def render(result: Dict[str, Any]) -> str:
    """Результат инструмента → компактный текст: заголовок, таблицы (| a | b |), примечания."""
    if result.get("error"):
        text = f"ОШИБКА: {result['error']}"
        if result.get("hint"):
            text += f"\nПодсказка: {result['hint']}"
        return text
    parts = []
    if result.get("title"):
        parts.append(f"## {result['title']}")
    for name, data in (result.get("tables") or {}).items():
        rows = data.get("rows") or []
        if name:
            parts.append(f"### {name}")
        if not rows:
            parts.append("(нет строк)")
            continue
        parts.append("| " + " | ".join(data["columns"]) + " |")
        parts.extend("| " + " | ".join(_cell(v) for v in row) + " |" for row in rows)
    for key, value in (result.get("facts") or {}).items():
        parts.append(f"- {key}: {_cell(value)}")
    for note in result.get("notes") or []:
        parts.append(f"Примечание: {note}")
    return "\n".join(parts)


# ---------- инструменты ----------

class AnalyticsTools:
    """Набор инструментов; одно соединение с базой на экземпляр, пересборка не чаще REFRESH_INTERVAL_SECONDS."""

    def __init__(self, root: Path, today: Optional[date] = None, clock: Callable[[], float] = time.monotonic):
        self.root = Path(root)
        self.today = today
        self.clock = clock
        self._conn: Optional[sqlite3.Connection] = None
        self._refreshed_at: Optional[float] = None
        self._live_cache: Dict[Tuple[str, str], Tuple[float, Dict[str, Any]]] = {}

    # --- соединение ---

    def _today(self) -> date:
        return self.today or date.today()

    def conn(self) -> sqlite3.Connection:
        now = self.clock()
        if self._refreshed_at is None or now - self._refreshed_at >= REFRESH_INTERVAL_SECONDS:
            warehouse.refresh(self.root)
            self._refreshed_at = now
            self.close()
        if self._conn is None:
            self._conn = warehouse.connect(self.root, refresh_first=False)
        return self._conn

    def close(self) -> None:
        if self._conn is not None:
            self._conn.close()
            self._conn = None

    def call(self, name: str, args: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        if name not in TOOL_NAMES:
            return {"error": f"Инструмент «{name}» не найден", "hint": "Доступны: " + ", ".join(TOOL_NAMES)}
        try:
            return getattr(self, name)(dict(args or {}))
        except ToolError as exc:
            return {"error": str(exc), **({"hint": exc.hint} if exc.hint else {})}
        except Exception as exc:  # ошибка инструмента — ответ модели, а не падение ассистента
            return {"error": f"{type(exc).__name__}: {exc}"}

    def _latest_period(self, marketplace: str = "ozon") -> Optional[str]:
        row = self.conn().execute(f"SELECT MAX(period) FROM {ORDERS_TABLE[marketplace]}").fetchone()
        return row[0] if row else None

    def _period_arg(self, args: Dict[str, Any], key: str, marketplace: str = "ozon", default_latest: bool = False) -> Optional[str]:
        raw = args.get(key)
        if raw in (None, ""):
            return self._latest_period(marketplace) if default_latest else None
        period = normalize_period(raw)
        if not period:
            raise ToolError(f"Не понял период «{raw}»", "Пишите месяц как '2026-09' или 'Сентябрь 2026'")
        return period

    def _incomplete_note(self, period: Optional[str]) -> List[str]:
        today = self._today()
        if period == f"{today.year}-{today.month:02d}":
            return [f"{period_title(period)} ещё не закончился: данные по {today.isoformat()}, сравнивать с полными месяцами нельзя напрямую."]
        return []

    # --- каталог и SQL ---

    def describe_data(self, args: Dict[str, Any]) -> Dict[str, Any]:
        catalog = warehouse.describe(self.conn())
        name = str(args.get("table") or "").strip()
        if name:
            item = next((i for i in catalog if i["table"] == name), None)
            if not item:
                raise ToolError(f"Таблицы «{name}» нет", "Вызовите describe_data без аргументов, чтобы увидеть список")
            result = {
                "title": f"Таблица {name}",
                "facts": {"Описание": item["description"], "Источник": item["source"], "Строк": item["rows"]},
                "tables": {"Колонки": table(["колонка", "описание"], list(item["columns"].items()))},
            }
            if item.get("range"):
                result["facts"]["Диапазон"] = " … ".join(str(x) for x in item["range"])
            if item.get("error"):
                result["notes"] = [f"Источник не прочитан: {item['error']}"]
            return result
        rows = []
        notes = []
        for item in catalog:
            span = " … ".join(str(x) for x in item["range"]) if item.get("range") else ""
            rows.append([item["table"], item["rows"], span, item["description"]])
            if item.get("error"):
                notes.append(f"{item['table']}: источник не прочитан ({item['error']})")
        notes.append("Колонки таблицы: describe_data(table=...). Свой запрос: run_sql.")
        return {"title": "Данные программы", "tables": {"": table(["таблица", "строк", "диапазон", "что внутри"], rows)}, "notes": notes}

    def run_sql(self, args: Dict[str, Any]) -> Dict[str, Any]:
        sql = str(args.get("sql") or "").strip().rstrip(";").strip()
        if not sql:
            raise ToolError("Пустой запрос")
        if not re.match(r"(?is)^(select|with)\b", sql):
            raise ToolError("Разрешены только запросы SELECT / WITH", "База доступна только для чтения")
        try:
            limit = max(1, min(int(args.get("limit") or SQL_ROW_LIMIT), SQL_ROW_LIMIT_MAX))
        except (TypeError, ValueError):
            limit = SQL_ROW_LIMIT
        conn = self.conn()
        deadline = time.monotonic() + SQL_TIMEOUT_SECONDS
        conn.set_progress_handler(lambda: 1 if time.monotonic() > deadline else 0, 10000)
        conn.set_authorizer(_read_only_authorizer)
        try:
            cursor = conn.execute(sql)
            columns = [d[0] for d in cursor.description or []]
            rows = cursor.fetchmany(limit + 1)
        except sqlite3.Error as exc:
            message = str(exc)
            if "interrupted" in message:
                raise ToolError(f"Запрос дольше {SQL_TIMEOUT_SECONDS} с — прерван", "Сузьте выборку или добавьте условия")
            raise ToolError(f"SQL: {message}", "Проверьте имена таблиц и колонок: describe_data(table=...)")
        finally:
            conn.set_progress_handler(None, 0)
            conn.set_authorizer(None)
        result = {"title": "Результат запроса", "tables": {"": table(columns, [tuple(r) for r in rows[:limit]])}}
        if len(rows) > limit:
            result["notes"] = [f"Показаны первые {limit} строк, есть ещё. Сгруппируйте или отфильтруйте запрос."]
        return result

    # --- готовые показатели ---

    def sales_metrics(self, args: Dict[str, Any]) -> Dict[str, Any]:
        marketplace = str(args.get("marketplace") or "ozon").lower()
        if marketplace not in MARKETPLACES:
            raise ToolError(f"Неизвестная площадка «{marketplace}»", "ozon или wb")
        group_by = args.get("group_by") or []
        if isinstance(group_by, str):
            group_by = [g for g in re.split(r"[,\s]+", group_by) if g]
        allowed = ["period", "artikul", "status"] + (["scheme"] if marketplace == "ozon" else [])
        unknown = [g for g in group_by if g not in allowed]
        if unknown:
            raise ToolError(f"Нельзя группировать по {', '.join(unknown)}", "Можно: " + ", ".join(allowed))
        period_from = self._period_arg(args, "period_from", marketplace)
        period_to = self._period_arg(args, "period_to", marketplace)
        if not period_from and not period_to:
            period_from = period_to = self._latest_period(marketplace)
        artikuls = _artikuls(args.get("artikuls"))

        sold = ",".join(f"'{s}'" for s in SOLD_STATUSES[marketplace])
        delivered = DELIVERED_STATUS[marketplace]
        expressions = {
            "orders": "COUNT(*)",
            "delivered": f"SUM(status = '{delivered}')",
            "cancelled": f"SUM(status = '{CANCELLED_STATUS[marketplace]}')",
            "returned": f"SUM(status = '{RETURNED_STATUS[marketplace]}')",
            "units": f"SUM(CASE WHEN status = '{delivered}' THEN quantity ELSE 0 END)",
            "revenue": f"SUM(CASE WHEN status IN ({sold}) THEN sale_price ELSE 0 END)",
            "profit": "SUM(profit)",
            "cost": "SUM(cost)",
            "avg_price": f"AVG(CASE WHEN status = '{delivered}' AND sale_price <> 0 THEN sale_price END)",
        }
        if marketplace == "ozon":
            expressions["commission"] = "SUM(ABS(commission))"
            expressions["logistics"] = "SUM(ABS(logistics))"
        expressions["margin_pct"] = f"100.0 * SUM(profit) / NULLIF({expressions['revenue']}, 0)"
        expressions["cancel_rate_pct"] = f"100.0 * {expressions['cancelled']} / COUNT(*)"

        wanted = args.get("metrics") or list(expressions)
        if isinstance(wanted, str):
            wanted = [m for m in re.split(r"[,\s]+", wanted) if m]
        unknown = [m for m in wanted if m not in expressions]
        if unknown:
            raise ToolError(f"Неизвестные показатели: {', '.join(unknown)}", "Доступны: " + ", ".join(expressions))

        where, params = [], []
        if period_from:
            where.append("period >= ?")
            params.append(period_from)
        if period_to:
            where.append("period <= ?")
            params.append(period_to)
        if artikuls:
            where.append(f"artikul IN ({','.join('?' for _ in artikuls)})")
            params.extend(artikuls)
        select = list(group_by) + [f"{expressions[m]} AS {m}" for m in wanted]
        if "artikul" in group_by:
            select.insert(group_by.index("artikul") + 1, "MAX(product_name) AS product_name")
        sql = f"SELECT {', '.join(select)} FROM {ORDERS_TABLE[marketplace]}"
        if where:
            sql += " WHERE " + " AND ".join(where)
        if group_by:
            sql += " GROUP BY " + ", ".join(group_by)
        # Разбивка по месяцам — хронология; иначе рейтинг по прибыли.
        sort_by = args.get("sort_by") or ("period" if "period" in group_by else ("profit" if "profit" in wanted else wanted[0]))
        if sort_by not in wanted and sort_by not in group_by:
            raise ToolError(f"Сортировать можно по выбранным колонкам, а не по «{sort_by}»")
        order = "ASC" if str(args.get("order") or ("asc" if sort_by in group_by else "desc")).lower() == "asc" else "DESC"
        sql += f" ORDER BY {sort_by} {order}"
        if sort_by == "period":
            sql += "".join(f", {g}" for g in group_by if g != "period")
        try:
            limit = max(1, min(int(args.get("limit") or 50), 500))
        except (TypeError, ValueError):
            limit = 50
        conn = self.conn()
        total = None
        if group_by:
            count_sql = f"SELECT COUNT(*) FROM (SELECT 1 FROM {ORDERS_TABLE[marketplace]}"
            count_sql += (" WHERE " + " AND ".join(where)) if where else ""
            count_sql += " GROUP BY " + ", ".join(group_by) + ")"
            total = conn.execute(count_sql, params).fetchone()[0]
        data = query_table(conn, sql + f" LIMIT {limit}", params)
        span = period_title(period_from) if period_from == period_to and period_from else \
            f"{period_title(period_from) if period_from else 'начало'} — {period_title(period_to) if period_to else 'конец'}"
        notes = [
            "revenue — выручка по доставленным и возвращённым заказам (как «Общая выручка» отчёта); "
            "profit — прибыль по заказам до рекламы, хранения и расходов вне заказов. "
            "Чистая прибыль месяца — в month_metrics.",
        ]
        if total is not None and total > limit:
            notes.append(f"Показаны {limit} из {total} строк (сортировка по {sort_by}).")
        notes += self._incomplete_note(period_to)
        return {"title": f"Продажи {marketplace.upper()}: {span}", "tables": {"": data}, "notes": notes}

    def month_metrics(self, args: Dict[str, Any]) -> Dict[str, Any]:
        marketplace = str(args.get("marketplace") or "ozon").lower()
        if marketplace not in MARKETPLACES:
            raise ToolError(f"Неизвестная площадка «{marketplace}»", "ozon или wb")
        conn = self.conn()
        period_to = self._period_arg(args, "period_to", marketplace, default_latest=True)
        period_from = self._period_arg(args, "period_from", marketplace) or (shift_period(period_to, -5) if period_to else None)
        if not period_to:
            raise ToolError("Отчётов за месяцы нет")
        metrics = args.get("metrics")
        if isinstance(metrics, str):
            metrics = [m.strip() for m in metrics.split(";") if m.strip()]
        metrics = metrics or KEY_MONTH_METRICS[marketplace]
        periods = [r[0] for r in conn.execute(
            f"SELECT DISTINCT period FROM {METRICS_TABLE[marketplace]} WHERE period BETWEEN ? AND ? ORDER BY period",
            (period_from, period_to))]
        values = {(p, m): (v, t) for p, m, v, t in conn.execute(
            f"SELECT period, metric, value, text FROM {METRICS_TABLE[marketplace]} WHERE period BETWEEN ? AND ?",
            (period_from, period_to))}
        known = {m for _, m in values}
        missing = [m for m in metrics if m not in known]
        rows = []
        for metric in metrics:
            if metric not in known:
                continue
            row = [metric]
            for period in periods:
                value, text = values.get((period, metric), (None, None))
                row.append(value if value is not None else text)
            rows.append(row)
        notes = []
        if missing:
            available = sorted(known)
            notes.append(f"Нет показателей: {'; '.join(missing)}. Есть: {'; '.join(available)}")
        notes += self._incomplete_note(period_to)
        return {
            "title": f"Итоги месяцев {marketplace.upper()} по отчётам",
            "tables": {"": table(["показатель"] + periods, rows)},
            "notes": notes,
        }

    def store_overview(self, args: Dict[str, Any]) -> Dict[str, Any]:
        conn = self.conn()
        period = self._period_arg(args, "period", "ozon", default_latest=True)
        if not period:
            raise ToolError("Отчётов Ozon нет: сформируйте месячный отчёт")
        current_month = f"{self._today().year}-{self._today().month:02d}"
        partial = None
        if not args.get("period") and period == current_month:
            # Неполный текущий месяц нельзя сравнивать с полными: обзор — по последнему полному, текущий — справкой.
            partial = period
            period = shift_period(period, -1)
        previous, year_ago = shift_period(period, -1), shift_period(period, -12)
        tables: Dict[str, Any] = {}

        compare = [p for p in (year_ago, previous, period)
                   if conn.execute("SELECT 1 FROM ozon_month_metrics WHERE period = ? LIMIT 1", (p,)).fetchone()]
        rows = []
        for metric in KEY_MONTH_METRICS["ozon"]:
            row, found = [metric], False
            for p in compare:
                hit = conn.execute("SELECT value FROM ozon_month_metrics WHERE period = ? AND metric = ?", (p, metric)).fetchone()
                row.append(hit[0] if hit else None)
                found = found or bool(hit)
            if found:
                if previous in compare and period in compare:
                    row.append(_change(row[-1], row[1 + compare.index(previous)], percent_metric="%" in metric))
                rows.append(row)
        columns = ["показатель"] + compare + (["к прошлому месяцу"] if previous in compare and period in compare else [])
        tables["Итоги месяца (из отчётов)"] = table(columns, rows)

        by_artikul = """
            SELECT artikul, MAX(product_name) AS product_name,
                   SUM(status = 'delivered') AS delivered,
                   SUM(CASE WHEN status IN ('delivered','returned') THEN sale_price ELSE 0 END) AS revenue,
                   SUM(profit) AS profit
            FROM ozon_orders WHERE period = ? GROUP BY artikul
        """
        tables["Лидеры по прибыли"] = query_table(conn, f"SELECT * FROM ({by_artikul}) ORDER BY profit DESC LIMIT 5", (period,))
        tables["Убыточные или самые слабые"] = query_table(conn, f"SELECT * FROM ({by_artikul}) ORDER BY profit ASC LIMIT 5", (period,))
        change = f"""
            WITH cur AS ({by_artikul}), prev AS ({by_artikul})
            SELECT COALESCE(cur.artikul, prev.artikul) AS artikul,
                   COALESCE(cur.product_name, prev.product_name) AS product_name,
                   COALESCE(prev.profit, 0) AS profit_prev, COALESCE(cur.profit, 0) AS profit_now,
                   COALESCE(cur.profit, 0) - COALESCE(prev.profit, 0) AS change
            FROM cur LEFT JOIN prev ON prev.artikul = cur.artikul
            UNION
            SELECT prev.artikul, prev.product_name, prev.profit, 0, -prev.profit
            FROM prev LEFT JOIN cur ON cur.artikul = prev.artikul WHERE cur.artikul IS NULL
        """
        tables["Сильнее всего выросла прибыль"] = query_table(
            conn, f"SELECT * FROM ({change}) WHERE change > 0 ORDER BY change DESC LIMIT 5", (period, previous))
        tables["Сильнее всего упала прибыль"] = query_table(
            conn, f"SELECT * FROM ({change}) WHERE change < 0 ORDER BY change ASC LIMIT 5", (period, previous))

        tables["Проблемы товаров (анализ органики)"] = query_table(conn, """
            SELECT artikul, main_problem, organic, organic_prev, stock,
                   substr(advice, 1, CASE WHEN instr(advice, char(10)) > 0 THEN instr(advice, char(10)) - 1 ELSE length(advice) END) AS advice
            FROM organic_products WHERE main_status = 'bad' ORDER BY priority DESC LIMIT 10
        """)
        tables["Нет остатка (последний снимок), но были продажи за 2 месяца"] = query_table(conn, """
            SELECT s.artikul, s.date, s.stock, COUNT(o.order_number) AS orders_2m
            FROM product_snapshots s
            JOIN ozon_orders o ON o.artikul = s.artikul AND o.period >= ?
            WHERE s.date = (SELECT MAX(date) FROM product_snapshots) AND COALESCE(s.stock, 0) <= 0
            GROUP BY s.artikul ORDER BY orders_2m DESC LIMIT 10
        """, (shift_period(period, -1),))
        tables["Не хватает на своём складе (FBS)"] = query_table(conn, """
            SELECT artikul, need, stock, shortage, cost FROM fbs_needs
            WHERE plan_date = (SELECT MAX(plan_date) FROM fbs_needs) AND shortage > 0 ORDER BY shortage DESC LIMIT 10
        """)
        wb = conn.execute("SELECT value FROM wb_month_metrics WHERE period = ? AND metric = 'Общая выручка'", (period,)).fetchone()
        facts = {}
        if partial:
            so_far = {m: v for m, v in conn.execute(
                "SELECT metric, value FROM ozon_month_metrics WHERE period = ? AND metric IN ('Общая выручка', 'Чистая прибыль', 'Общее количество заказов')",
                (partial,))}
            facts[f"{period_title(partial)} (идёт, по {self._today().isoformat()})"] = (
                f"выручка {_round(so_far.get('Общая выручка'))}, чистая прибыль {_round(so_far.get('Чистая прибыль'))}, "
                f"заказов {_round(so_far.get('Общее количество заказов'))} — неполный месяц, с полными не сравнивать")
        if wb and wb[0] is not None:
            facts["Выручка Wildberries за месяц"] = wb[0]
        notes = self._incomplete_note(period)
        if partial:
            notes.append(f"Обзор — за последний полный месяц ({period_title(period)}).")
        snapshot = conn.execute("SELECT MAX(date) FROM product_snapshots").fetchone()[0]
        organic = conn.execute("SELECT MAX(created) FROM organic_products").fetchone()[0]
        notes.append(f"Анализ органики от {organic or '—'}, снимок остатков от {snapshot or '—'}; "
                     "свежие остатки и цены — live_stocks / live_prices.")
        return {"title": f"Обзор магазина Ozon: {period_title(period)}", "tables": tables, "facts": facts, "notes": notes}

    def product_profile(self, args: Dict[str, Any]) -> Dict[str, Any]:
        artikul = normalize_artikul(args.get("artikul"))
        if not artikul:
            raise ToolError("Нужен артикул", "Не знаете артикул — find_products(query=...)")
        conn = self.conn()
        tables: Dict[str, Any] = {}
        facts: Dict[str, Any] = {}
        name = conn.execute("SELECT product_name FROM ozon_orders WHERE artikul = ? ORDER BY period DESC LIMIT 1", (artikul,)).fetchone()
        search = conn.execute("SELECT sku, name, category FROM search_products WHERE artikul = ?", (artikul,)).fetchone()
        product = conn.execute("SELECT * FROM products WHERE artikul = ?", (artikul,)).fetchone()
        if not (name or search or product):
            raise ToolError(f"Артикул {artikul} не найден ни в отчётах, ни в costs.xlsx", "Найдите его: find_products(query=...)")
        facts["Название"] = (name and name[0]) or (search and search["name"]) or "—"
        if search:
            facts["SKU Ozon"] = search["sku"]
            if search["category"]:
                facts["Категория"] = search["category"]
        if product:
            labels = {c: d for c, (_, d) in warehouse.CATALOG["products"]["columns"].items()}
            for key in product.keys():
                if key != "artikul" and product[key] is not None:
                    facts[labels[key]] = _round(product[key])
        abc = conn.execute("SELECT abc, xyz, abcxyz, report_period FROM abc_xyz_latest WHERE artikul = ?", (artikul,)).fetchone()
        if abc:
            facts["ABC/XYZ"] = f"{abc['abcxyz']} (ABC {abc['abc']}, XYZ {abc['xyz']}; отчёт {abc['report_period']})"

        last = self._latest_period("ozon") or f"{self._today().year}-{self._today().month:02d}"
        tables["Продажи Ozon по месяцам (12 мес.)"] = query_table(conn, """
            SELECT period, COUNT(*) AS orders, SUM(status = 'delivered') AS delivered, SUM(status = 'cancelled') AS cancelled,
                   SUM(CASE WHEN status IN ('delivered','returned') THEN sale_price ELSE 0 END) AS revenue,
                   SUM(profit) AS profit,
                   AVG(CASE WHEN status = 'delivered' AND sale_price <> 0 THEN sale_price END) AS avg_price,
                   SUM(scheme = 'FBO') AS fbo_orders
            FROM ozon_orders WHERE artikul = ? AND period > ? GROUP BY period ORDER BY period
        """, (artikul, shift_period(last, -12)))
        tables["Акции"] = query_table(conn, "SELECT promotion, promo_price, boost_pct FROM product_promotions WHERE artikul = ?", (artikul,))
        organic = conn.execute("SELECT * FROM organic_products WHERE artikul = ?", (artikul,)).fetchone()
        if organic:
            tables["Анализ органики"] = table(["показатель", "значение"], [
                ["окно", f"{organic['period_start']} … {organic['period_end']}"],
                ["заказы / было", f"{_round(organic['orders'])} / {_round(organic['orders_prev'])}"],
                ["органика / было", f"{_round(organic['organic'])} / {_round(organic['organic_prev'])}"],
                ["с рекламы", organic["ad_orders"]], ["остаток", organic["stock"]],
                ["цена сейчас / мин.", f"{_round(organic['price_now'])} / {_round(organic['min_price'])}"],
                ["индекс цен", organic["price_index_color"]], ["контент-рейтинг", organic["content_rating"]],
                ["доставка из кластера покупателя, %", organic["local_share"]],
                ["главная проблема", f"{organic['main_problem']} [{organic['main_status']}]"],
                ["диагностика", organic["diagnosis"]], ["советы", organic["advice"]],
            ])
            tables["Анализ органики"]["rows"] = [r for r in tables["Анализ органики"]["rows"] if r[1] not in (None, "")]
        tables["Снимки цены и остатка"] = query_table(conn, """
            SELECT date, price, stock, price_index, promotions_count FROM product_snapshots
            WHERE artikul = ? ORDER BY date DESC LIMIT 7
        """, (artikul,))
        tables["План поставки FBO (последний)"] = query_table(conn, """
            SELECT plan_date, arrival, zone, quantity, per_day, own_stock FROM fbo_supply_plan
            WHERE artikul = ? AND plan_date = (SELECT MAX(plan_date) FROM fbo_supply_plan)
        """, (artikul,))
        need = conn.execute("""
            SELECT need, stock, shortage FROM fbs_needs WHERE artikul = ? AND plan_date = (SELECT MAX(plan_date) FROM fbs_needs)
        """, (artikul,)).fetchone()
        if need:
            facts["FBS: нужно / есть / не хватает"] = f"{_round(need['need'])} / {_round(need['stock'])} / {_round(need['shortage'])}"
        if search:
            sku = search["sku"]
            last_day = conn.execute("SELECT MAX(date) FROM search_sku_daily WHERE sku = ?", (sku,)).fetchone()[0]
            if last_day:
                tables["Поиск за 28 дней"] = query_table(conn, """
                    SELECT MIN(date) AS since, MAX(date) AS till, SUM(searched_users) AS searched_users,
                           SUM(viewed_users) AS viewed_users, AVG(position) AS avg_position,
                           SUM(ordered_units) AS ordered_units, SUM(revenue) AS revenue
                    FROM search_sku_daily WHERE sku = ? AND date > date(?, '-28 days')
                """, (sku, last_day))
                tables["Главные поисковые запросы (28 дней)"] = query_table(conn, """
                    SELECT k.query, k.cluster, SUM(q.searched_users) AS searched_users, SUM(q.orders) AS orders,
                           MAX(sk.is_target) AS is_target
                    FROM search_query_stats q JOIN search_keywords k ON k.keyword_id = q.keyword_id
                    LEFT JOIN search_sku_keywords sk ON sk.sku = q.sku AND sk.keyword_id = q.keyword_id
                    WHERE q.sku = ? AND q.date > date(?, '-28 days')
                    GROUP BY k.keyword_id ORDER BY searched_users DESC LIMIT 10
                """, (sku, last_day))
        wb = query_table(conn, """
            SELECT period, COUNT(*) AS orders, SUM(CASE WHEN status IN ('продажа','возврат') THEN sale_price ELSE 0 END) AS revenue,
                   SUM(profit) AS profit FROM wb_orders WHERE artikul = ? GROUP BY period ORDER BY period
        """, (artikul,))
        if wb["rows"]:
            tables["Продажи Wildberries по месяцам"] = wb
        tables = {k: v for k, v in tables.items() if v["rows"]}
        return {"title": f"Товар {artikul}", "facts": facts, "tables": tables, "notes": self._incomplete_note(last)}

    def supply_needs(self, args: Dict[str, Any]) -> Dict[str, Any]:
        conn = self.conn()
        plan_date = conn.execute("SELECT MAX(plan_date) FROM fbo_supply_plan").fetchone()[0]
        fbs_date = conn.execute("SELECT MAX(plan_date) FROM fbs_needs").fetchone()[0]
        if not plan_date and not fbs_date:
            raise ToolError("Плана поставки нет", "Рассчитайте его в разделе «Поставки» программы")
        artikuls = _artikuls(args.get("artikuls"))
        filter_sql = f" AND artikul IN ({','.join('?' for _ in artikuls)})" if artikuls else ""
        # Строка плана — артикул × зона; total, per_day и cost у артикула общие, их не суммируем по зонам.
        data = query_table(conn, f"""
            WITH fbo AS (
                SELECT artikul, MAX(total) AS fbo_ship, GROUP_CONCAT(zone || ': ' || CAST(quantity AS INTEGER), '; ') AS by_zone,
                       MAX(cost) AS fbo_cost, MAX(per_day) AS per_day
                FROM fbo_supply_plan WHERE plan_date = ? {filter_sql} GROUP BY artikul),
            fbs AS (SELECT artikul, need, stock, shortage, cost FROM fbs_needs WHERE plan_date = ? {filter_sql}),
            keys AS (SELECT artikul FROM fbo UNION SELECT artikul FROM fbs)
            SELECT k.artikul, fbo.fbo_ship, fbo.by_zone, fbo.fbo_cost, fbs.need AS fbs_need, fbs.stock AS fbs_stock,
                   fbs.shortage AS fbs_shortage, fbs.cost AS fbs_cost,
                   COALESCE(fbo.fbo_cost, 0) + COALESCE(fbs.cost, 0) AS total_cost, fbo.per_day
            FROM keys k LEFT JOIN fbo ON fbo.artikul = k.artikul LEFT JOIN fbs ON fbs.artikul = k.artikul
            WHERE COALESCE(fbo.fbo_ship, 0) > 0 OR COALESCE(fbs.shortage, 0) > 0
            ORDER BY total_cost DESC
        """, [plan_date, *artikuls, fbs_date, *artikuls])
        fbo_units = sum(r[1] or 0 for r in data["rows"])
        fbo_cost = sum(r[3] or 0 for r in data["rows"])
        fbs_units = sum(r[6] or 0 for r in data["rows"])
        fbs_cost = sum(r[7] or 0 for r in data["rows"])
        arrival = conn.execute("SELECT MAX(arrival) FROM fbo_supply_plan WHERE plan_date = ?", (plan_date,)).fetchone()[0]
        facts = {
            "План рассчитан": plan_date or fbs_date,
            "Ожидаемый приход поставки FBO": arrival or "—",
            "Везти на FBO, шт. / себестоимость, ₽": f"{_round(fbo_units)} / {_round(fbo_cost)}",
            "Не хватает на своём складе (FBS), шт. / себестоимость, ₽": f"{_round(fbs_units)} / {_round(fbs_cost)}",
            "Всего денег на пополнение (по себестоимости), ₽": _round(fbo_cost + fbs_cost),
        }
        data["columns"] = ["artikul", "везти на FBO, шт.", "по зонам", "на FBO, ₽", "нужно на FBS, шт.", "есть на FBS, шт.",
                           "не хватает на FBS, шт.", "на FBS, ₽", "всего, ₽", "продаж в день, шт."]
        notes = ["Деньги — по себестоимости из costs.xlsx. «Не хватает на FBS» — сколько докупить или произвести для своего склада."]
        try:
            age = (self._today() - date.fromisoformat(plan_date or fbs_date)).days
        except (TypeError, ValueError):
            age = None
        if age is not None and age > 7:
            notes.append(f"План рассчитан {age} дн. назад — пересчитайте его в разделе «Поставки».")
        if len(data["rows"]) > 60:
            data["rows"] = data["rows"][:60]
            notes.append("Показаны 60 артикулов с наибольшей суммой.")
        return {"title": "Что пополнить: план поставки", "facts": facts, "tables": {"": data}, "notes": notes}

    def product_table(self, args: Dict[str, Any]) -> Dict[str, Any]:
        conn = self.conn()
        latest = self._latest_period("ozon")
        if not latest:
            raise ToolError("Отчётов Ozon нет: сформируйте месячный отчёт")
        try:
            months = max(1, min(int(args.get("months") or 3), 24))
        except (TypeError, ValueError):
            months = 3
        since = shift_period(latest, -(months - 1))
        before = shift_period(since, -months)
        data = query_table(conn, """
            WITH sales AS (
                SELECT artikul,
                       SUM(CASE WHEN period >= ? AND status = 'delivered' THEN quantity ELSE 0 END) AS units,
                       SUM(CASE WHEN period >= ? THEN profit ELSE 0 END) AS profit,
                       SUM(CASE WHEN period >= ? AND period < ? THEN profit ELSE 0 END) AS profit_prev,
                       MAX(product_name) AS product_name
                FROM ozon_orders WHERE period >= ? GROUP BY artikul),
            keys AS (SELECT artikul FROM products UNION SELECT artikul FROM sales WHERE artikul NOT LIKE '%,%')
            SELECT k.artikul, COALESCE(sales.product_name, sp.name) AS product_name, abc.abcxyz,
                   snap.stock, p.current_price, p.min_price, p.cost, snap.price_index, snap.promotions_count,
                   sales.units, sales.profit, sales.profit_prev, org.main_problem, fbs.shortage AS fbs_shortage
            FROM keys k
            LEFT JOIN sales ON sales.artikul = k.artikul
            LEFT JOIN products p ON p.artikul = k.artikul
            LEFT JOIN search_products sp ON sp.artikul = k.artikul
            LEFT JOIN abc_xyz_latest abc ON abc.artikul = k.artikul
            LEFT JOIN product_snapshot_latest snap ON snap.artikul = k.artikul
            LEFT JOIN organic_products org ON org.artikul = k.artikul AND org.main_status = 'bad'
            LEFT JOIN fbs_needs_latest fbs ON fbs.artikul = k.artikul
        """, (since, since, before, since, before))
        columns = ["artikul", "товар", "ABC/XYZ", "остаток (снимок)", "цена", "мин. цена", "себестоимость", "индекс цен",
                   "акций", f"продано за {months} мес., шт.", f"прибыль за {months} мес., ₽", f"прибыль за прошлые {months} мес., ₽",
                   "главная проблема", "не хватает на FBS"]
        keyed = [dict(zip(columns, row)) for row in data["rows"]]

        abc = [a.upper() for a in (args.get("abc") or [])] if isinstance(args.get("abc"), list) else \
            [a for a in str(args.get("abc") or "").upper().replace(",", " ").split() if a]
        if abc:
            keyed = [r for r in keyed if r["ABC/XYZ"] and r["ABC/XYZ"][0] in abc]
        if args.get("out_of_stock"):
            keyed = [r for r in keyed if (r["остаток (снимок)"] or 0) <= 0]
        if args.get("below_min_price"):
            keyed = [r for r in keyed if r["цена"] and r["мин. цена"] and r["цена"] < r["мин. цена"]]
        artikuls = set(_artikuls(args.get("artikuls")))
        if artikuls:
            keyed = [r for r in keyed if r["artikul"] in artikuls]
        sort_key = {"profit": columns[10], "units": columns[9], "stock": columns[3], "artikul": "artikul",
                    "profit_change": None}.get(str(args.get("sort_by") or "profit"), columns[10])
        if sort_key is None:
            keyed.sort(key=lambda r: (r[columns[10]] or 0) - (r[columns[11]] or 0))
        else:
            keyed.sort(key=lambda r: (r[sort_key] is None, r[sort_key] if sort_key == "artikul" else -(r[sort_key] or 0)))
        snapshot = conn.execute("SELECT MAX(date) FROM product_snapshots").fetchone()[0]
        abc_report = conn.execute("SELECT MAX(report_period) FROM abc_xyz_latest").fetchone()[0]
        notes = [f"Продажи и прибыль — {period_title(since)} … {period_title(latest)} (прошлые: {period_title(before)} … {period_title(shift_period(since, -1))}); "
                 f"остаток и индекс цен — снимок {snapshot or '—'}; ABC/XYZ — отчёт {abc_report or '—'}; "
                 "главная проблема — из анализа органики (только критичные). Свежие остатки — live_stocks."]
        notes += self._incomplete_note(latest)
        return {"title": f"Товары: {len(keyed)}", "tables": {"": table(columns, [[r[c] for c in columns] for r in keyed])}, "notes": notes}

    def calculate(self, args: Dict[str, Any]) -> Dict[str, Any]:
        expressions = args.get("expressions") or args.get("expression") or []
        if isinstance(expressions, str):
            expressions = [expressions]
        if not expressions:
            raise ToolError("Нужно выражение", "Например: expressions=['103989 - 22262', '103989 / 22262']")
        rows = []
        for expression in list(expressions)[:30]:
            try:
                value = _safe_eval(str(expression))
                rows.append([str(expression), round(value, 4) if isinstance(value, float) else value])
            except (ValueError, SyntaxError, ZeroDivisionError, OverflowError, TypeError) as exc:
                rows.append([str(expression), f"ошибка: {exc}"])
        return {"title": "Расчёт", "tables": {"": table(["выражение", "результат"], rows)}}

    def find_products(self, args: Dict[str, Any]) -> Dict[str, Any]:
        query = str(args.get("query") or "").strip()
        if not query:
            raise ToolError("Нужен фрагмент названия или артикула")
        needle = query.casefold()
        conn = self.conn()
        found: Dict[str, Dict[str, Any]] = {}

        def add(artikul: Any, name: Any, where: str) -> None:
            # «1101, 1105» — отправление из нескольких товаров: название у него общее, берём только сами артикулы.
            parts = [normalize_artikul(a) for a in str(artikul or "").split(",")]
            for part in filter(None, parts):
                item = found.setdefault(part, {"name": None, "where": set()})
                item["name"] = item["name"] or (name if len(parts) == 1 else None)
                item["where"].add(where)

        for artikul, name in conn.execute("SELECT artikul, MAX(product_name) FROM ozon_orders GROUP BY artikul"):
            add(artikul, name, "Ozon")
        for artikul, name in conn.execute("SELECT artikul, name FROM search_products"):
            add(artikul, name, "Ozon")
        for (artikul,) in conn.execute("SELECT artikul FROM products"):
            add(artikul, None, "costs.xlsx")
        for artikul, name in conn.execute("SELECT artikul, MAX(product_name) FROM wb_orders GROUP BY artikul"):
            add(artikul, name, "WB")
        for artikul, name in conn.execute("SELECT artikul, name FROM wb_products"):
            add(artikul, name, "WB")
        words = needle.split()
        rows = []
        for artikul, item in found.items():
            haystack = f"{artikul} {item['name'] or ''}".casefold()
            if artikul.casefold().startswith(needle) or all(w in haystack for w in words):
                rows.append([artikul, item["name"] or "", ", ".join(sorted(item["where"]))])
        rows.sort(key=lambda r: (not r[0].casefold().startswith(needle), len(r[0]), r[0]))
        notes = [] if rows else ["Ничего не найдено. Попробуйте часть слова или артикула."]
        if len(rows) > 30:
            notes.append(f"Найдено {len(rows)}, показаны 30 — уточните запрос.")
        return {"title": f"Товары по запросу «{query}»", "tables": {"": table(["artikul", "название", "где есть"], rows[:30])}, "notes": notes}

    # --- живые данные Ozon (только чтение) ---

    def _live(self, key: str, args: Dict[str, Any], loader: Callable[[], Dict[str, Any]]) -> Dict[str, Any]:
        cache_key = (key, repr(sorted(args.items())))
        now = self.clock()
        cached = self._live_cache.get(cache_key)
        if cached and now - cached[0] < LIVE_CACHE_SECONDS:
            return cached[1]
        _require_ozon_keys()
        result = loader()
        result.setdefault("notes", []).append(f"Данные Ozon на {datetime.now().strftime('%Y-%m-%d %H:%M')}.")
        self._live_cache[cache_key] = (now, result)
        return result

    def _live_artikuls(self, args: Dict[str, Any]) -> List[str]:
        artikuls = _artikuls(args.get("artikuls"))
        if artikuls:
            return artikuls[:200]
        return [r[0] for r in self.conn().execute("SELECT artikul FROM products ORDER BY artikul")]

    def live_prices(self, args: Dict[str, Any]) -> Dict[str, Any]:
        artikuls = self._live_artikuls(args)
        if not artikuls:
            raise ToolError("Нет артикулов: в costs.xlsx пусто", "Передайте artikuls")

        min_prices = {a: p for a, p in self.conn().execute("SELECT artikul, min_price FROM products")}

        def load() -> Dict[str, Any]:
            organic = _import("organic_analysis")
            prices = organic.load_prices(artikuls)
            rows = []
            for artikul in artikuls:
                item = prices.get(artikul)
                own_min = min_prices.get(artikul)
                if not item:
                    rows.append([artikul, None, own_min, None, None, None, None, None, "нет в Ozon"])
                    continue
                index = item.get("index") or {}
                price = item.get("price")
                target = organic.price_for_good_index(price, index.get("value"))
                if target is None:
                    verdict = "индекс в норме" if index.get("color") in ("GREEN", "YELLOW") else "—"
                elif own_min and target < own_min:
                    verdict = "снижать нельзя: нужная цена ниже вашей минимальной"
                else:
                    verdict = "можно снизить до этой цены"
                rows.append([artikul, price, own_min, index.get("color"), index.get("value"), index.get("min_price"),
                             target, len(item.get("actions") or []), verdict])
            return {"title": "Текущие цены Ozon", "tables": {"": table(
                ["artikul", "цена для покупателя", "ваша мин. цена", "индекс цен", "значение индекса", "мин. цена конкурента",
                 "цена для выгодного индекса", "акций", "вывод"], rows)},
                "notes": ["Индекс цен: GREEN — выгодная, YELLOW — умеренная, RED — невыгодная, WITHOUT_INDEX — не с чем сравнить.",
                          "Цена для выгодного индекса — ориентир Ozon (≈ цена / индекс). Ниже вашей минимальной цены продавать "
                          "нельзя — это ниже минимальной маржи."]}
        return self._live("live_prices", {"artikuls": artikuls}, load)

    def live_stocks(self, args: Dict[str, Any]) -> Dict[str, Any]:
        artikuls = self._live_artikuls(args)
        if not artikuls:
            raise ToolError("Нет артикулов: в costs.xlsx пусто", "Передайте artikuls")

        def load() -> Dict[str, Any]:
            supply = _import("fbo_supply_plan")
            stocks: Dict[str, Dict[str, int]] = {}
            for i in range(0, len(artikuls), 1000):
                cursor = ""
                while True:
                    payload: Dict[str, Any] = {"filter": {"offer_id": artikuls[i:i + 1000], "visibility": "ALL"}, "limit": 1000}
                    if cursor:
                        payload["cursor"] = cursor
                    data = supply.ozon_post("https://api-seller.ozon.ru/v4/product/info/stocks", payload)
                    items = data.get("items") or []
                    for item in items:
                        by_type = stocks.setdefault(normalize_artikul(item.get("offer_id")), {})
                        for stock in item.get("stocks") or []:
                            kind = str(stock.get("type") or "").lower()
                            free = int(stock.get("present") or 0) - int(stock.get("reserved") or 0)
                            by_type[kind] = by_type.get(kind, 0) + max(0, free)
                    cursor = data.get("cursor") or ""
                    if not items or not cursor or len(items) < 1000:
                        break
            rows = [[a, stocks.get(a, {}).get("fbo"), stocks.get(a, {}).get("fbs"),
                     sum(stocks.get(a, {}).values()) if a in stocks else None] for a in artikuls]
            return {"title": "Текущие остатки Ozon", "tables": {"": table(["artikul", "FBO свободно", "FBS свободно", "всего"], rows)},
                    "notes": ["Свободно = в наличии минус резерв. Пусто — товара нет в Ozon."]}
        return self._live("live_stocks", {"artikuls": artikuls}, load)

    def live_promotions(self, args: Dict[str, Any]) -> Dict[str, Any]:
        def load() -> Dict[str, Any]:
            prices = _import("recommended_prices")
            rows = []
            for raw in prices.get_actions_list():
                info = prices.describe_action(raw)
                rows.append([info["id"], info["name"], info["action_type"], str(info["date_start"])[:10], str(info["date_end"])[:10],
                             raw.get("participating_products_count"), raw.get("potential_products_count"),
                             info.get("fixed_boost")])
            rows.sort(key=lambda r: r[3] or "")
            return {"title": "Акции Ozon сейчас", "tables": {"": table(
                ["id", "название", "тип", "начало", "конец", "участвует товаров", "можно добавить", "бустинг, %"], rows)},
                "notes": ["Цены и бустинг товаров в акциях — таблица product_promotions (по costs.xlsx)."]}
        return self._live("live_promotions", {}, load)

    def live_discount_requests(self, args: Dict[str, Any]) -> Dict[str, Any]:
        status = str(args.get("status") or "NEW").upper()
        if status not in ("NEW", "SEEN", "APPROVED", "DECLINED", "ALL"):
            raise ToolError(f"Неизвестный статус «{status}»", "NEW, SEEN, APPROVED, DECLINED или ALL")

        def load() -> Dict[str, Any]:
            prices = _import("recommended_prices")
            tasks = prices.get_discount_requests(status=status)
            # В заявке есть SKU, но нет артикула: берём связь из базы поиска. Имена модераторов и почту не отдаём.
            by_sku = {sku: artikul for sku, artikul in self.conn().execute("SELECT sku, artikul FROM search_products")}
            rows = [[str(t.get("created_at") or "")[:16].replace("T", " "), t.get("status"), by_sku.get(t.get("sku")),
                     t.get("name"), t.get("original_price"), t.get("requested_price"), t.get("requested_discount"),
                     t.get("approved_price"), t.get("requested_quantity_max")] for t in tasks[:100]]
            notes = [f"Всего заявок: {len(tasks)}." + (" Показаны первые 100." if len(tasks) > 100 else "")]
            return {"title": f"Заявки покупателей на скидку ({status})", "tables": {"": table(
                ["создана (UTC)", "статус", "artikul", "товар", "цена", "просит", "скидка, %", "одобрено", "шт."], rows)}, "notes": notes}
        return self._live("live_discount_requests", {"status": status}, load)


class ToolError(Exception):
    def __init__(self, message: str, hint: Optional[str] = None):
        super().__init__(message)
        self.hint = hint


_OPERATORS = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul, ast.Div: operator.truediv,
              ast.Pow: operator.pow, ast.Mod: operator.mod, ast.FloorDiv: operator.floordiv}
_FUNCTIONS = {"round": round, "abs": abs, "min": min, "max": max, "sum": lambda *xs: sum(xs), "sqrt": math.sqrt}


def _safe_eval(expression: str) -> float:
    """Арифметика без eval: числа, + - * / // % **, скобки, round/abs/min/max/sum/sqrt."""
    def walk(node: ast.AST) -> Any:
        if isinstance(node, ast.Expression):
            return walk(node.body)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
            return node.value
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.USub, ast.UAdd)):
            value = walk(node.operand)
            return -value if isinstance(node.op, ast.USub) else value
        if isinstance(node, ast.BinOp) and type(node.op) in _OPERATORS:
            left, right = walk(node.left), walk(node.right)
            if isinstance(node.op, ast.Pow) and abs(right) > 100:
                raise ValueError("слишком большая степень")
            return _OPERATORS[type(node.op)](left, right)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id in _FUNCTIONS and not node.keywords:
            return _FUNCTIONS[node.func.id](*(walk(arg) for arg in node.args))
        raise ValueError("разрешены только числа, + - * / % ** и round, abs, min, max, sum, sqrt")
    text = expression.replace("\u00a0", "").replace(" ", "").replace(",", ".") if re.fullmatch(r"[\d\s,.+\-*/()%]+", expression) else expression
    return walk(ast.parse(text, mode="eval"))


def _read_only_authorizer(action: int, arg1: Any, arg2: Any, db_name: Any, trigger: Any) -> int:
    allowed = {sqlite3.SQLITE_SELECT, sqlite3.SQLITE_READ, sqlite3.SQLITE_FUNCTION}
    recursive = getattr(sqlite3, "SQLITE_RECURSIVE", None)
    if recursive is not None:
        allowed.add(recursive)
    return sqlite3.SQLITE_OK if action in allowed else sqlite3.SQLITE_DENY


def _import(name: str):
    try:
        return __import__(f"scripts.{name}", fromlist=[name])
    except ModuleNotFoundError:
        return __import__(name)


def _require_ozon_keys() -> None:
    prices = _import("recommended_prices")
    prices.refresh_api_credentials()
    if not prices.OZON_CLIENT_ID or not prices.OZON_API_KEY:
        raise ToolError("Ключи Ozon не настроены", "Живые данные недоступны; используйте данные отчётов")


# ---------- описания для модели ----------

def tool_schema(name: str, description: str, properties: Optional[Dict[str, Any]] = None, required: Optional[List[str]] = None) -> Dict[str, Any]:
    return {"type": "function", "function": {"name": name, "description": description, "parameters": {
        "type": "object", "properties": properties or {}, "required": required or []}}}


_PERIOD = {"type": "string", "description": "Месяц: '2026-09' или 'Сентябрь 2026'"}
_ARTIKULS = {"type": "array", "items": {"type": "string"}, "description": "Артикулы продавца; пусто — все товары из costs.xlsx"}
_MARKETPLACE = {"type": "string", "enum": list(MARKETPLACES), "description": "Площадка: ozon (по умолчанию) или wb"}

TOOL_SCHEMAS: List[Dict[str, Any]] = [
    tool_schema("store_overview",
                "Обзор магазина Ozon за месяц: итоги месяца против прошлого месяца и года назад, лидеры и аутсайдеры по прибыли, "
                "товары с самым большим ростом и падением, проблемы из анализа органики, нет остатка, нехватка на FBS. "
                "Начинай с него на общие вопросы «как дела», «что происходит», «что улучшить».",
                {"period": {**_PERIOD, "description": "Месяц; по умолчанию последний с данными"}}),
    tool_schema("month_metrics",
                "Официальные итоги месяцев из отчётов по месяцам (таблица показатель × месяц): выручка, чистая прибыль "
                "(после рекламы, хранения и расходов вне заказов), рентабельность, продвижение, комиссии, логистика, хранение, баланс. "
                "Для динамики магазина и вопросов о чистой прибыли.",
                {"period_from": _PERIOD, "period_to": {**_PERIOD, "description": "По умолчанию последний месяц"},
                 "metrics": {"type": "array", "items": {"type": "string"}, "description": "Названия показателей как в отчёте; пусто — ключевые"},
                 "marketplace": _MARKETPLACE}),
    tool_schema("sales_metrics",
                "Показатели продаж, посчитанные по заказам, с разбивкой: orders, delivered, cancelled, returned, units, revenue, profit, "
                "cost, avg_price, commission, logistics, margin_pct, cancel_rate_pct. Для рейтингов товаров, сравнения месяцев, "
                "схем FBO/FBS, отмен. Можно за любой диапазон месяцев.",
                {"period_from": _PERIOD, "period_to": _PERIOD,
                 "group_by": {"type": "array", "items": {"type": "string", "enum": ["period", "artikul", "scheme", "status"]},
                              "description": "Разбивка; пусто — один итог"},
                 "artikuls": {"type": "array", "items": {"type": "string"}, "description": "Только эти артикулы"},
                 "metrics": {"type": "array", "items": {"type": "string"}, "description": "Нужные показатели; пусто — все"},
                 "sort_by": {"type": "string", "description": "Показатель или колонка разбивки для сортировки"},
                 "order": {"type": "string", "enum": ["desc", "asc"]},
                 "limit": {"type": "integer", "description": "Сколько строк, по умолчанию 50"},
                 "marketplace": _MARKETPLACE}),
    tool_schema("product_profile",
                "Всё об одном артикуле: себестоимость и цены, ABC/XYZ, продажи и прибыль по месяцам за год, акции, анализ органики "
                "с диагнозом и советами, динамика цены и остатка, план поставки, поисковый спрос и главные запросы, продажи на WB.",
                {"artikul": {"type": "string", "description": "Артикул продавца, например '1302'"}}, ["artikul"]),
    tool_schema("product_table",
                "Сводная таблица всех товаров, одна строка на артикул: класс ABC/XYZ, остаток, цена, мин. цена, себестоимость, "
                "индекс цен, число акций, продажи и прибыль за последние N месяцев и за N месяцев до них, главная проблема, "
                "нехватка на FBS. Для вопросов «какие товары…» по нескольким признакам сразу.",
                {"months": {"type": "integer", "description": "За сколько последних месяцев продажи, по умолчанию 3"},
                 "abc": {"type": "array", "items": {"type": "string", "enum": ["A", "B", "C"]}, "description": "Только эти классы ABC"},
                 "out_of_stock": {"type": "boolean", "description": "Только без остатка"},
                 "below_min_price": {"type": "boolean", "description": "Только с ценой ниже минимальной"},
                 "artikuls": {"type": "array", "items": {"type": "string"}, "description": "Только эти артикулы"},
                 "sort_by": {"type": "string", "enum": ["profit", "units", "stock", "profit_change", "artikul"],
                             "description": "Сортировка, по умолчанию profit (по убыванию); profit_change — сначала худшая динамика"}}),
    tool_schema("supply_needs",
                "Что и сколько пополнить по последнему плану поставки: сколько штук везти на склады Ozon (FBO) по зонам, "
                "сколько не хватает на своём складе (FBS), сколько денег нужно по себестоимости — по артикулам и итогом.",
                {"artikuls": {"type": "array", "items": {"type": "string"}, "description": "Только эти артикулы; пусто — все"}}),
    tool_schema("calculate",
                "Калькулятор. Любые вычисления для ответа — разница, доля, процент изменения, «во сколько раз», сумма "
                "нескольких чисел — делай только им, не в уме. Можно несколько выражений сразу.",
                {"expressions": {"type": "array", "items": {"type": "string"},
                                 "description": "Выражения, например ['103989 - 22262', '(9305.14 - 12601.54) / 12601.54 * 100']"}},
                ["expressions"]),
    tool_schema("find_products",
                "Найти артикулы по фрагменту названия или артикула (Ozon и WB).",
                {"query": {"type": "string", "description": "Фрагмент, например 'доска 35' или '13'"}}, ["query"]),
    tool_schema("describe_data",
                "Каталог всех данных программы: таблицы, число строк, периоды, описание. С table — колонки таблицы и их смысл. "
                "Вызывай перед run_sql. Представления *_latest уже содержат только самый свежий отчёт, снимок или план.",
                {"table": {"type": "string", "description": "Имя таблицы; пусто — список таблиц"}}),
    tool_schema("run_sql",
                "Свой SQL-запрос (SQLite, только SELECT/WITH) по базе программы, когда готовых инструментов не хватает. "
                "Сначала узнай колонки через describe_data. Группируй и агрегируй в SQL, не выгружай сырые строки.",
                {"sql": {"type": "string", "description": "Один запрос SELECT или WITH"},
                 "limit": {"type": "integer", "description": f"Максимум строк, по умолчанию {SQL_ROW_LIMIT}"}}, ["sql"]),
    tool_schema("live_prices",
                "Сейчас в Ozon (запрос к API): цена для покупателя, индекс цен, минимальная цена конкурента, акции товара.",
                {"artikuls": _ARTIKULS}),
    tool_schema("live_stocks",
                "Сейчас в Ozon (запрос к API): свободные остатки на складах Ozon (FBO) и на своём складе (FBS).",
                {"artikuls": _ARTIKULS}),
    tool_schema("live_promotions",
                "Сейчас в Ozon (запрос к API): список акций, даты, сколько товаров участвует и сколько можно добавить."),
    tool_schema("live_discount_requests",
                "Сейчас в Ozon (запрос к API): заявки покупателей на скидку. Только просмотр, решения не отправляются.",
                {"status": {"type": "string", "enum": ["NEW", "SEEN", "APPROVED", "DECLINED", "ALL"], "description": "По умолчанию NEW"}}),
]
TOOL_NAMES = [schema["function"]["name"] for schema in TOOL_SCHEMAS]
LIVE_TOOLS = {name for name in TOOL_NAMES if name.startswith("live_")}


if __name__ == "__main__":
    import json
    import sys

    tools = AnalyticsTools(warehouse.ROOT)
    name = sys.argv[1] if len(sys.argv) > 1 else "describe_data"
    arguments = json.loads(sys.argv[2]) if len(sys.argv) > 2 else {}
    print(render(tools.call(name, arguments)))
