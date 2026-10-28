# -*- coding: utf-8 -*-
"""
План запасов Wildberries. Считается только по данным WB: свои настройки (wb_supply_settings.json),
свои заказы, остатки, себестоимость (wb_costs.xlsx) — с планом Ozon не смешивается, даже если склад физически один.

Логика (с цифрами она же выводится на лист «Как считается» в отчёте):
1. Темп продаж — заказы WB без отмен: среднее из темпа за последние 30 дней и за весь период, который WB отдал
   (до 90 дней; на деле метод статистики отдаёт меньше — делим на дни, за которые заказы есть, а не на 90).
   Среднее двух темпов быстрее замечает рост и спад: на истории WB ошибка 67% против 81% у темпа за 90 дней.
2. Сезонность — по выкупам WB из финансовых отчётов (история с первого отчёта, сейчас с 09.2025):
   продажи в день в том же окне год назад / продажи в день в те же 90 дней год назад
   (если год назад истории меньше — по той части, что есть). Мало данных — коэффициент 1. Пока истории меньше
   двух лет, коэффициент ослабляем вдвое: один год молодого магазина — скорее шум, чем сезонность.
3. Доля артикула — заказы (приведённые к 90 дням) ×2 плюс выкупы 3–12 месяцев назад (приведённые к 90 дням).
4. Склады WB (FBW): остатки WB отдаёт только общей суммой «Склад WB», поэтому считается общий объём без разбивки
   по складам. Возим только то, что стабильно продаётся (прогноз от N шт. в месяц).
   Цель = квантиль Пуассона от спроса за цикл; к поставке = цель − (остаток WB + возвраты в пути − спрос до прихода).
5. Свой склад (FBS) — всё, что не покроет FBW: спрос артикулов, которые на WB не возим, и спрос до прихода поставки
   сверх остатка на складах WB. Нужно = квантиль Пуассона; не хватает = нужно − остаток FBS, выставленный на WB.
"""

import json
import math
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import requests

script_dir = Path(__file__).resolve().parent
code_root = script_dir.parent
for path in (str(code_root), str(script_dir)):
    if path not in sys.path:
        sys.path.insert(0, path)
from app_paths import DATA_ROOT  # noqa: E402
repo_root = DATA_ROOT

from fbo_supply_plan import STOCKS_REPORTS_DIR, _article_sort_key, format_rub, poisson_quantile  # noqa: E402

try:
    from scripts.file_io import atomic_output_path
except ModuleNotFoundError:
    from file_io import atomic_output_path

WB_CACHE_DIR = ".wb_sales_cache"
SETTINGS_FILENAME = "wb_supply_settings.json"
WB_DEFAULTS: Dict[str, Any] = {"lead_days": 14, "cover_days": 30, "service_level": 0.9, "wb_fbw_min_monthly": 3.0}
RECENT_DAYS = 90
HISTORY_MONTHS = 24
SEASON_MIN_UNITS = 8           # меньше выкупов в окне или базе год назад — сезонность WB не считаем
SEASON_MIN_BASE_DAYS = 28      # минимум дней истории в базовом периоде год назад
SEASON_LIMITS = (0.5, 3.0)
SEASON_FULL_HISTORY_DAYS = 730  # меньше двух лет истории — сезонность ослабляем
SEASON_SHRINK = 0.5
RATE_RECENT_DAYS = 30
TIMEOUT = 60


class WbSupplyError(RuntimeError):
    pass


def load_settings(root: Path) -> Dict[str, Any]:
    """Настройки плана WB — отдельно от плана Ozon."""
    settings = dict(WB_DEFAULTS)
    path = root / SETTINGS_FILENAME
    if path.exists():
        try:
            settings.update(json.loads(path.read_text(encoding="utf-8")))
        except (OSError, ValueError) as exc:
            print(f"⚠️ Не удалось прочитать {SETTINGS_FILENAME}, использую значения по умолчанию: {exc}")
    return settings


def save_settings(root: Path, settings: Dict[str, Any]) -> None:
    (root / SETTINGS_FILENAME).write_text(
        json.dumps({key: settings[key] for key in WB_DEFAULTS}, ensure_ascii=False, indent=2), encoding="utf-8")


def unit_costs_by_article(costs_by_nm: Dict[str, Any], cards: Dict[str, Dict[str, Any]]) -> Dict[str, float]:
    """Себестоимость из wb_costs.xlsx (по nmId) -> по артикулу продавца."""
    result: Dict[str, float] = {}
    for article, entry in cards.items():
        for nm_id in entry["nm_ids"]:
            value = costs_by_nm.get(str(nm_id))
            if value is not None:
                result[article] = float(value)
                break
    return result


def _normalize(article: Any) -> str:
    text = str(article or "").strip()
    try:
        return str(int(float(text)))
    except ValueError:
        return text


# ---------------------------------------------------------------- загрузка из WB (только чтение)

def _get(url: str, token: str, params=None) -> Any:
    response = requests.get(url, headers={"Authorization": token}, params=params, timeout=TIMEOUT)
    if response.status_code != 200:
        raise WbSupplyError(f"WB вернул HTTP {response.status_code} на {url.split('.ru')[-1]}")
    return response.json()


def _post(url: str, token: str, payload: Dict[str, Any]) -> Any:
    response = requests.post(url, headers={"Authorization": token}, json=payload, timeout=TIMEOUT)
    if response.status_code != 200:
        raise WbSupplyError(f"WB вернул HTTP {response.status_code} на {url.split('.ru')[-1]}")
    return response.json()


def load_recent_orders(token: str, today: date) -> pd.DataFrame:
    """Заказы за 90 дней: date, article, qty, cancelled, own (со своего склада)."""
    since = today - timedelta(days=RECENT_DAYS - 1)
    rows = _get("https://statistics-api.wildberries.ru/api/v1/supplier/orders", token, {"dateFrom": since.isoformat()})
    records = []
    for row in rows or []:
        day = pd.Timestamp(str(row.get("date"))[:10])
        if day < pd.Timestamp(since):
            continue
        records.append({
            "date": day,
            "article": _normalize(row.get("supplierArticle")),
            "qty": 1,
            "cancelled": bool(row.get("isCancel")),
            "own": str(row.get("warehouseType") or "") == "Склад продавца",
        })
    return pd.DataFrame(records, columns=["date", "article", "qty", "cancelled", "own"])


def _sales_from_details(details: List[Dict[str, Any]]) -> List[List[Any]]:
    result = []
    for row in details:
        if row.get("sellerOperName") != "Продажа" or not row.get("quantity"):
            continue
        day = str(row.get("orderDt") or row.get("saleDt") or "")[:10]
        if day:
            result.append([day, _normalize(row.get("vendorCode")), int(row.get("quantity") or 0)])
    return result


def load_sales_history(root: Path, token: str, today: date, progress=print) -> pd.DataFrame:
    """
    Выкупы WB по дате заказа из финансовых отчётов. Закрытые месяцы (старше прошлого) кэшируются,
    остальное догружается одним диапазоном: у финансового API лимит — 1 запрос в минуту.
    """
    from scripts.wb_finance import FinanceClient

    cache = root / STOCKS_REPORTS_DIR / WB_CACHE_DIR
    cache.mkdir(parents=True, exist_ok=True)
    first_month = (today.replace(day=1) - timedelta(days=HISTORY_MONTHS * 31)).replace(day=1)
    months, cursor = [], first_month
    while cursor <= today:
        months.append(cursor)
        cursor = (cursor.replace(day=28) + timedelta(days=4)).replace(day=1)
    previous_month = (today.replace(day=1) - timedelta(days=1)).replace(day=1)
    rows: List[List[Any]] = []
    missing = []
    for month in months:
        path = cache / f"{month:%Y-%m}.json"
        if month < previous_month and path.exists():
            rows += json.loads(path.read_text(encoding="utf-8"))
        else:
            missing.append(month)
    if missing:
        start = missing[0]
        progress(f"· Загружаю выкупы WB с {start:%m.%Y} (финансовые отчёты, 1 запрос в минуту)...")
        client = FinanceClient(token, progress=lambda message: None)
        payload = {"dateFrom": start.isoformat(), "dateTo": today.isoformat() + "T23:59:59", "period": "daily"}
        details, cursor_id = [], 0
        while True:
            page = client.request("/detailed", {**payload, "limit": 100000, "rrdId": cursor_id})
            if not page:
                break
            details += page
            cursor_id = page[-1]["rrdId"]
            if len(page) < 100000:
                break
        fresh = _sales_from_details(details)
        for month in missing:
            key = f"{month:%Y-%m}"
            part = [row for row in fresh if row[0].startswith(key)]
            if month < previous_month:
                (cache / f"{key}.json").write_text(json.dumps(part, ensure_ascii=False), encoding="utf-8")
        rows += [row for row in fresh if row[0][:7] in {f"{m:%Y-%m}" for m in missing}]
    frame = pd.DataFrame(rows, columns=["date", "article", "qty"])
    frame["date"] = pd.to_datetime(frame["date"])
    return frame


def load_cards(token: str) -> Dict[str, Dict[str, Any]]:
    """Артикул продавца -> nmId и chrtId размеров (для остатков)."""
    cards: Dict[str, Dict[str, Any]] = {}
    cursor: Dict[str, Any] = {"limit": 100}
    while True:
        body = _post("https://content-api.wildberries.ru/content/v2/get/cards/list", token,
                     {"settings": {"cursor": cursor, "filter": {"withPhoto": -1}}})
        page = body.get("cards") or []
        for card in page:
            article = _normalize(card.get("vendorCode"))
            entry = cards.setdefault(article, {"nm_ids": [], "chrt_ids": []})
            entry["nm_ids"].append(card.get("nmID"))
            entry["chrt_ids"] += [size["chrtID"] for size in card.get("sizes") or [] if size.get("chrtID")]
        next_cursor = body.get("cursor") or {}
        if len(page) < cursor["limit"] or not next_cursor.get("nmID"):
            return cards
        cursor = {"limit": 100, "updatedAt": next_cursor.get("updatedAt"), "nmID": next_cursor.get("nmID")}


def load_own_stock(token: str, cards: Dict[str, Dict[str, Any]]) -> Dict[str, int]:
    """Остаток FBS на своих складах WB по артикулам."""
    chrt_to_article = {chrt: article for article, entry in cards.items() for chrt in entry["chrt_ids"]}
    result: Dict[str, int] = {}
    warehouses = _get("https://marketplace-api.wildberries.ru/api/v3/warehouses", token)
    chrt_ids = list(chrt_to_article)
    for warehouse in warehouses or []:
        for start in range(0, len(chrt_ids), 1000):
            body = _post(f"https://marketplace-api.wildberries.ru/api/v3/stocks/{warehouse['id']}", token,
                         {"chrtIds": chrt_ids[start:start + 1000]})
            for row in body.get("stocks") or []:
                article = chrt_to_article.get(row.get("chrtId"))
                if article:
                    result[article] = result.get(article, 0) + int(row.get("amount") or 0)
    return result


def load_wb_stock(token: str, cards: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, int]]:
    """Остатки на складах WB (общей суммой) и возвраты в пути по артикулам."""
    nm_to_article = {nm: article for article, entry in cards.items() for nm in entry["nm_ids"]}
    result: Dict[str, Dict[str, int]] = {}
    offset = 0
    while True:
        body = _post("https://seller-analytics-api.wildberries.ru/api/analytics/v1/stocks-report/wb-warehouses", token,
                     {"limit": 1000, "offset": offset})
        items = (body.get("data") or {}).get("items") or []
        for row in items:
            article = nm_to_article.get(row.get("nmId"))
            if not article:
                continue
            entry = result.setdefault(article, {"quantity": 0, "to_client": 0, "from_client": 0})
            entry["quantity"] += int(row.get("quantity") or 0)
            entry["to_client"] += int(row.get("inWayToClient") or 0)
            entry["from_client"] += int(row.get("inWayFromClient") or 0)
        if len(items) < 1000:
            return result
        offset += 1000


# ---------------------------------------------------------------- расчёт

def _units(frame: pd.DataFrame, start: date, end: date) -> float:
    if frame.empty:
        return 0.0
    part = frame[(frame["date"] >= pd.Timestamp(start)) & (frame["date"] < pd.Timestamp(end))]
    return float(part["qty"].sum())


def season_factor(history: pd.DataFrame, today: date, start: date, end: date) -> Dict[str, Any]:
    """
    Сезонный коэффициент окна [start, end): продажи в день в том же окне год назад
    к продажам в день в последние 90 дней, сдвинутые на год (по той части, где есть история).
    """
    shift = timedelta(days=364)  # 52 недели — дни недели совпадают
    first = history["date"].min().date() if not history.empty else today
    base_start = max(today - timedelta(days=RECENT_DAYS) - shift, first)
    base_end = today - shift
    window_start, window_end = start - shift, end - shift
    base_days = (base_end - base_start).days
    window_units = _units(history, window_start, window_end)
    base_units = _units(history, base_start, base_end)
    info = {"window_last_year": window_units, "base_last_year": base_units, "base_days": max(0, base_days),
            "base_from": base_start, "window_from": window_start, "window_to": window_end}
    if window_start < first or base_days < SEASON_MIN_BASE_DAYS or window_units < SEASON_MIN_UNITS or base_units < SEASON_MIN_UNITS:
        return {**info, "factor": 1.0, "method": "none"}
    raw = (window_units / (window_end - window_start).days) / (base_units / base_days)
    factor = min(max(raw, SEASON_LIMITS[0]), SEASON_LIMITS[1])
    weight = 1.0 if (today - first).days >= SEASON_FULL_HISTORY_DAYS else SEASON_SHRINK
    return {**info, "factor": 1 + (factor - 1) * weight, "raw": raw, "weight": weight, "method": "season"}


def order_days(orders: pd.DataFrame, today: date) -> int:
    """За сколько дней WB отдал заказы: от первого заказа до сегодня, не больше 90 (метод статистики отдаёт меньше)."""
    if orders.empty:
        return RECENT_DAYS
    first = orders["date"].min().date()
    return max(1, min(RECENT_DAYS, (today - first).days + 1))


def order_rate(orders: pd.DataFrame, today: date) -> Dict[str, float]:
    """Темп заказов в день: среднее из темпа за 30 дней и за весь отданный период."""
    days = order_days(orders, today)
    live = orders[~orders["cancelled"]] if len(orders) else orders
    if live.empty:
        return {"rate": 0.0, "rate_all": 0.0, "rate_recent": 0.0, "days": days}
    recent_days = min(RATE_RECENT_DAYS, days)
    rate_all = float(live["qty"].sum()) / days
    rate_recent = _units(live, today - timedelta(days=recent_days - 1), today + timedelta(days=1)) / recent_days
    return {"rate": (rate_all + rate_recent) / 2, "rate_all": rate_all, "rate_recent": rate_recent, "days": days}


def article_shares(orders: pd.DataFrame, history: pd.DataFrame, today: date) -> pd.Series:
    """Заказы (приведённые к 90 дням) ×2 + выкупы 3–12 месяцев назад, приведённые к 90 дням."""
    recent = orders[~orders["cancelled"]].groupby("article")["qty"].sum() if len(orders) else pd.Series(dtype=float)
    recent = recent * (RECENT_DAYS / order_days(orders, today))
    old_start, old_end = today - timedelta(days=365), today - timedelta(days=RECENT_DAYS)
    old = history[(history["date"] >= pd.Timestamp(old_start)) & (history["date"] < pd.Timestamp(old_end))] if len(history) else history
    older = old.groupby("article")["qty"].sum() * (RECENT_DAYS / (old_end - old_start).days) if len(old) else pd.Series(dtype=float)
    score = recent.mul(2).add(older, fill_value=0)
    return score / score.sum() if score.sum() > 0 else score


def build_wb_plan(
    orders: pd.DataFrame,
    history: pd.DataFrame,
    own_stock: Dict[str, int],
    wb_stock: Dict[str, Dict[str, int]],
    settings: Dict[str, Any],
    today: date,
    articles: Optional[List[str]] = None,
) -> Dict[str, Any]:
    lead, cover = int(settings["lead_days"]), int(settings["cover_days"])
    level = float(settings["service_level"])
    min_monthly = float(settings.get("wb_fbw_min_monthly", WB_DEFAULTS["wb_fbw_min_monthly"]))
    arrival = today + timedelta(days=lead)
    window = (arrival, arrival + timedelta(days=cover))

    live = orders[~orders["cancelled"]] if len(orders) else orders
    rates = order_rate(orders, today)
    rate = rates["rate"]
    season_window = season_factor(history, today, *window)
    season_lead = season_factor(history, today, today, arrival)
    window_units = rate * cover * season_window["factor"]
    lead_units = rate * lead * season_lead["factor"]
    shares = article_shares(orders, history, today)

    recent_by_article = live.groupby("article")["qty"].sum() if len(live) else pd.Series(dtype=float)
    own_by_article = live[live["own"]].groupby("article")["qty"].sum() if len(live) else pd.Series(dtype=float)
    all_articles = sorted(set(articles or []) | set(shares.index) | set(own_stock) | set(wb_stock), key=_article_sort_key)
    rows = []
    for article in all_articles:
        share = float(shares.get(article, 0.0))
        d_lead, d_window = lead_units * share, window_units * share
        monthly = d_window / cover * 30 if cover else 0.0
        stock = wb_stock.get(article, {})
        fbw_now = int(stock.get("quantity", 0)) + int(stock.get("from_client", 0))
        fbw = monthly >= min_monthly
        if fbw:
            target = poisson_quantile(d_window, level)
            on_arrival = max(0.0, fbw_now - d_lead)
            ship = max(0, math.ceil(target - on_arrival))
            fbs_demand = max(0.0, d_lead - fbw_now)  # до прихода поставки сверх остатка WB
        else:
            target, on_arrival, ship = 0, max(0.0, fbw_now - d_lead), 0
            fbs_demand = max(0.0, d_lead + d_window - fbw_now)  # сначала продаётся то, что уже на WB
        fbs_need = poisson_quantile(fbs_demand, level) if fbs_demand > 0 else 0
        fbs_stock = own_stock.get(article)
        shortage = max(0, fbs_need - (fbs_stock or 0))
        if share == 0:
            note = "Нет продаж на WB"
        elif fbw and ship == 0:
            note = "На складах WB хватит"
        elif not fbw:
            note = f"Меньше {min_monthly:g} шт. в месяц — продаём со своего склада"
        else:
            note = ""
        rows.append({
            "Артикул": article,
            "Заказов за период": int(recent_by_article.get(article, 0)),
            "Из них со своего склада": int(own_by_article.get(article, 0)),
            "Доля артикула, %": round(share * 100, 1),
            "Спрос до прихода, шт.": round(d_lead, 1),
            "Спрос за цикл, шт.": round(d_window, 1),
            "В месяц, шт.": round(monthly, 1),
            "Возить на WB": "да" if fbw else "нет",
            "Остаток на складах WB": int(stock.get("quantity", 0)),
            "Возвраты в пути": int(stock.get("from_client", 0)),
            "Цель FBW": target,
            "К поставке на WB": ship,
            "Спрос на своём складе, шт.": round(fbs_demand, 1),
            "Нужно на своём складе": fbs_need,
            "Остаток FBS WB": fbs_stock,
            "Не хватает FBS WB": shortage,
            "Комментарий": note,
        })
    return {
        "today": today,
        "arrival": arrival,
        "window": window,
        "rate": rate,
        "rates": rates,
        "recent_units": float(live["qty"].sum()) if len(live) else 0.0,
        "cancelled": int(orders["cancelled"].sum()) if len(orders) else 0,
        "own_share": float(live["own"].mean()) if len(live) else 0.0,
        "season_window": season_window,
        "season_lead": season_lead,
        "window_units": window_units,
        "lead_units": lead_units,
        "details": pd.DataFrame(rows),
        "history": history,
        "settings": settings,
    }


def _explain(plan: Dict[str, Any], costs: Dict[str, Any]) -> pd.DataFrame:
    s, sw, r = plan["settings"], plan["season_window"], plan["rates"]
    w0, w1 = plan["window"]
    fmt = lambda d: d.strftime("%d.%m.%Y")
    if sw["method"] == "season":
        season = (f"Год назад в окне {fmt(sw['window_from'])}–{fmt(sw['window_to'] - timedelta(days=1))} выкупили {sw['window_last_year']:.0f} шт., "
                  f"в базовые {sw['base_days']} дн. — {sw['base_last_year']:.0f} шт.: окно в {sw['raw']:.2f} раза активнее обычного "
                  f"(ограничен {SEASON_LIMITS[0]}–{SEASON_LIMITS[1]}"
                  + (f", истории меньше двух лет — ослаблен вдвое" if sw["weight"] < 1 else "") + f"): коэффициент {sw['factor']:.2f}.")
    else:
        season = "Истории WB год назад мало для сезонности — коэффициент 1 (продажи как сейчас)."
    lines = [
        ("Параметры", f"Свои для WB: срок {s['lead_days']} дн., цикл {s['cover_days']} дн., уровень сервиса {s['service_level']:.0%}."),
        ("Окно", f"С {fmt(plan['today'])} по {fmt(w1 - timedelta(days=1))}; поставка на склады WB придёт ~{fmt(plan['arrival'])}."),
        ("Темп продаж", f"WB отдал заказы за {r['days']} дн.: без отмен {plan['recent_units']:.0f} шт. ({r['rate_all']:.2f} в день), "
                        f"за последние {min(RATE_RECENT_DAYS, r['days'])} дн. — {r['rate_recent']:.2f} в день, отменено {plan['cancelled']}. "
                        f"Темп — среднее двух: {plan['rate']:.2f} в день. Со своего склада ушло {plan['own_share']:.0%} заказов."),
        ("Сезонность", season),
        ("Прогноз", f"До прихода поставки ~{plan['lead_units']:.1f} шт., в окне ~{plan['window_units']:.1f} шт."),
        ("Доля артикула", "Заказы (приведённые к 90 дням) с весом 2 плюс выкупы 3–12 месяцев назад, приведённые к 90 дням."),
        ("Склады WB (FBW)", f"WB отдаёт остатки только общей суммой, поэтому считается общий объём на WB без разбивки по складам. "
                            f"Возим артикулы с прогнозом от {s.get('wb_fbw_min_monthly', 3):g} шт. в месяц: цель — запас под уровень сервиса, "
                            "к поставке = цель − (остаток WB + возвраты в пути − спрос до прихода)."),
        ("Свой склад для WB", "Спрос артикулов, которые на WB не возим (сверх остатка WB), и спрос до прихода поставки сверх остатка WB. "
                              "Нужно — запас под уровень сервиса; не хватает = нужно − остаток FBS, выставленный на WB. "
                              "Остатки и поставки Ozon сюда не входят."),
        ("Себестоимость", f"Из wb_costs.xlsx: к поставке на WB {format_rub(costs['fbw'])}, нехватка своего склада {format_rub(costs['fbs'])}."
                          + (f" Без себестоимости: {', '.join(costs['missing'])}." if costs["missing"] else "")),
        ("Ограничения", "WB отдаёт заказы не больше чем за 90 дней (на деле меньше), а продаж мало — прогноз грубый: на истории WB "
                         "ошибка прогноза магазина на 30 дней — около двух третей от факта. Страховка по артикулам при этом "
                         "откалибрована: факт не выше цели примерно с той вероятностью, что задана уровнем сервиса."),
    ]
    return pd.DataFrame(lines, columns=["Шаг", "Как считается"])


def plan_costs(plan: Dict[str, Any], unit_costs: Dict[str, float]) -> Dict[str, Any]:
    details = plan["details"]
    totals, missing = {"fbw": 0.0, "fbs": 0.0}, set()
    for _, row in details.iterrows():
        article = str(row["Артикул"])
        for key, column in (("fbw", "К поставке на WB"), ("fbs", "Не хватает FBS WB")):
            qty = int(row[column])
            if not qty:
                continue
            if article in unit_costs:
                totals[key] += unit_costs[article] * qty
            else:
                missing.add(article)
    return {"fbw": round(totals["fbw"], 2), "fbs": round(totals["fbs"], 2), "missing": sorted(missing, key=_article_sort_key)}


def summarize_plan(plan: Dict[str, Any], unit_costs: Dict[str, float]) -> Dict[str, Any]:
    details = plan["details"]
    sw = plan["season_window"]

    def cost_of(article: str, qty: int) -> Optional[float]:
        unit = unit_costs.get(article)
        return round(unit * qty, 2) if unit is not None and qty else None

    fbw_rows = [
        {"article": str(r["Артикул"]), "monthly": float(r["В месяц, шт."]), "stock": int(r["Остаток на складах WB"]),
         "returns": int(r["Возвраты в пути"]), "target": int(r["Цель FBW"]), "ship": int(r["К поставке на WB"]),
         "cost": cost_of(str(r["Артикул"]), int(r["К поставке на WB"]))}
        for _, r in details[details["Возить на WB"] == "да"].iterrows()
    ]
    fbs_rows = [
        {"article": str(r["Артикул"]), "demand": float(r["Спрос на своём складе, шт."]), "need": int(r["Нужно на своём складе"]),
         "stock": None if r["Остаток FBS WB"] is None or pd.isna(r["Остаток FBS WB"]) else int(r["Остаток FBS WB"]),
         "shortage": int(r["Не хватает FBS WB"]), "orders": int(r["Заказов за период"]),
         "cost": cost_of(str(r["Артикул"]), int(r["Не хватает FBS WB"]))}
        for _, r in details[(details["Нужно на своём складе"] > 0) | (details["Остаток FBS WB"].fillna(0) > 0)].iterrows()
    ]
    fbw_rows.sort(key=lambda r: (-r["ship"], -r["monthly"]))
    fbs_rows.sort(key=lambda r: (-r["shortage"], -r["need"], _article_sort_key(r["article"])))
    history = plan["history"]
    monthly = history.groupby(history["date"].dt.to_period("M"))["qty"].sum() if len(history) else pd.Series(dtype=float)
    return {
        "created": plan["today"].isoformat(),
        "settings": {key: plan["settings"][key] for key in WB_DEFAULTS},
        "arrival": plan["arrival"].isoformat(),
        "window": [plan["window"][0].isoformat(), plan["window"][1].isoformat()],
        "recent_units": plan["recent_units"],
        "cancelled": plan["cancelled"],
        "own_share": plan["own_share"],
        "rate": plan["rate"],
        "order_days": plan["rates"]["days"],
        "season": {"method": sw["method"], "factor": sw["factor"], "weight": sw.get("weight", 1.0), "window_last_year": sw["window_last_year"],
                   "base_last_year": sw["base_last_year"], "base_days": sw["base_days"]},
        "window_units": plan["window_units"],
        "lead_units": plan["lead_units"],
        "fbw": {"ship": sum(r["ship"] for r in fbw_rows), "articles": sum(1 for r in fbw_rows if r["ship"]),
                "stock": int(details["Остаток на складах WB"].sum()), "cost": round(sum(r["cost"] or 0 for r in fbw_rows), 2), "rows": fbw_rows},
        "fbs": {"need": int(details["Нужно на своём складе"].sum()), "stock": int(details["Остаток FBS WB"].fillna(0).sum()),
                "shortage": int(details["Не хватает FBS WB"].sum()), "cost": round(sum(r["cost"] or 0 for r in fbs_rows), 2), "rows": fbs_rows},
        "monthly_sales": [{"month": str(period), "units": int(units)} for period, units in monthly.items()],
        "missing_costs": plan_costs(plan, unit_costs)["missing"],
    }


def write_report(plan: Dict[str, Any], unit_costs: Dict[str, float], out_path: Path) -> None:
    from openpyxl.utils import get_column_letter

    history = plan["history"]
    monthly = history.groupby(history["date"].dt.to_period("M"))["qty"].sum() if len(history) else pd.Series(dtype=float)
    sheets = {
        "Поставка WB": plan["details"][plan["details"]["К поставке на WB"] > 0][
            ["Артикул", "В месяц, шт.", "Остаток на складах WB", "Возвраты в пути", "Цель FBW", "К поставке на WB"]],
        "Свой склад для WB": plan["details"][(plan["details"]["Нужно на своём складе"] > 0) | (plan["details"]["Остаток FBS WB"].fillna(0) > 0)][
            ["Артикул", "Заказов за период", "Спрос на своём складе, шт.", "Нужно на своём складе", "Остаток FBS WB", "Не хватает FBS WB"]],
        "Расчёт": plan["details"],
        "Как считается": _explain(plan, plan_costs(plan, unit_costs)),
        "Выкупы WB по месяцам": pd.DataFrame({"Месяц": [p.strftime("%m.%Y") for p in monthly.index], "Выкупы, шт.": monthly.values.astype(int)}),
    }
    with atomic_output_path(out_path) as temporary:
        with pd.ExcelWriter(temporary, engine="openpyxl") as writer:
            for name, frame in sheets.items():
                frame.to_excel(writer, sheet_name=name, index=False)
                ws = writer.sheets[name]
                ws.freeze_panes = "B2"
                for idx, column in enumerate(frame.columns, start=1):
                    values = [str(column)] + [str(v) for v in frame[column].head(200).tolist()]
                    ws.column_dimensions[get_column_letter(idx)].width = min(90, max(10, max(len(v) for v in values) + 2))


def run(root: Path, overrides: Optional[Dict[str, Any]] = None) -> bool:
    from scripts.marketplace_settings import read_settings
    from scripts.wb_finance import read_costs

    token = read_settings(root).get("WB_API_TOKEN")
    if not token:
        print("❌ Добавьте API-токен WB в настройках.")
        return False
    settings = load_settings(root)
    settings.update({key: value for key, value in (overrides or {}).items() if value is not None})
    save_settings(root, settings)
    today = date.today()

    try:
        print("· Загружаю заказы WB (до 90 дней)...")
        orders = load_recent_orders(token, today)
        history = load_sales_history(root, token, today)
        print("· Загружаю карточки и остатки (свой склад и склады WB)...")
        cards = load_cards(token)
        own_stock = load_own_stock(token, cards)
        wb_stock = load_wb_stock(token, cards)
        unit_costs = unit_costs_by_article(read_costs(root), cards)
    except (WbSupplyError, requests.RequestException) as exc:
        print(f"❌ {exc}")
        return False
    except Exception as exc:  # FinanceError и прочие ошибки загрузки
        print(f"❌ Не удалось загрузить данные WB: {exc}")
        return False

    plan = build_wb_plan(orders, history, own_stock, wb_stock, settings, today, articles=list(cards))
    sw = plan["season_window"]
    print(f"· Заказов WB за {plan['rates']['days']} дн.: {plan['recent_units']:.0f} шт., темп {plan['rate']:.2f} в день, "
          f"со своего склада {plan['own_share']:.0%}.")
    print("· Сезонность окна: " + (f"×{sw['factor']:.2f} (год назад в окне {sw['window_last_year']:.0f} шт.)" if sw["method"] == "season" else "истории мало, ×1"))
    details = plan["details"]
    print(f"· На склады WB: {int(details['К поставке на WB'].sum())} шт.; свой склад для WB: нужно {int(details['Нужно на своём складе'].sum())}, "
          f"не хватает {int(details['Не хватает FBS WB'].sum())} шт.")
    costs = plan_costs(plan, unit_costs)
    print(f"· Себестоимость (wb_costs.xlsx): поставка на WB {format_rub(costs['fbw'])}, нехватка своего склада {format_rub(costs['fbs'])}"
          + (f" (без себестоимости: {', '.join(costs['missing'])})" if costs["missing"] else ""))
    out_dir = root / STOCKS_REPORTS_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"Поставка WB {today:%d.%m.%Y}.xlsx"
    write_report(plan, unit_costs, out_path)
    out_path.with_suffix(".json").write_text(
        json.dumps(summarize_plan(plan, unit_costs), ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    print(f"✅ Отчёт сохранён: {out_path}")
    return True


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="План запасов WB: свой склад и склады WB.")
    parser.add_argument("--lead-days", type=int, help="Дней до прихода поставки на склад WB")
    parser.add_argument("--cover-days", type=int, help="На сколько дней продаж рассчитан запас")
    parser.add_argument("--service-level", type=float, help="Вероятность не уйти в ноль, 0.5–0.99")
    parser.add_argument("--fbw-min-monthly", type=float, help="Возить на склады WB артикулы от N шт. в месяц")
    args = parser.parse_args()
    if args.service_level is not None and not 0.5 <= args.service_level <= 0.99:
        parser.error("Уровень сервиса должен быть от 0.5 до 0.99.")
    overrides = {"lead_days": args.lead_days, "cover_days": args.cover_days, "service_level": args.service_level,
                 "wb_fbw_min_monthly": args.fbw_min_monthly}
    sys.exit(0 if run(repo_root, overrides) else 1)


if __name__ == "__main__":
    main()
