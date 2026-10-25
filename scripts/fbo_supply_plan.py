# -*- coding: utf-8 -*-
"""
План поставки FBO по кластерам на основе спроса и сезонности.

Логика (с цифрами она же выводится на лист «Как считается» в отчёте):
1. Спрос — все заказы FBO и FBS (без отмен) из Ozon, по кластеру покупателя (cluster_to).
2. Прогноз магазина на окно поставки = продажи того же окна год назад × рост
   (последние 90 дней к тем же 90 дням год назад). Так учитывается сезонность.
3. Прогноз артикула = прогноз магазина × доля артикула (последние 180 дней,
   усреднённая с долей в том же окне год назад, если артикул тогда продавался).
4. Склад FBO обслуживает не только свой кластер: кластеры покупателей закреплены
   за складами (зонами). Доля зоны по артикулу сглаживается к доле по магазину,
   потому что по одному артикулу заказов мало.
5. Запас на приходе = (свободно + в пути + в заявках по кластерам зоны) − спрос зоны за срок поставки.
6. Цель = квантиль Пуассона (уровень сервиса) от спроса зоны за цикл; к поставке = цель − запас на приходе.
7. Свой склад (FBS) на те же дни — от сегодня до конца окна продаж: спрос регионов, которые обслуживает FBS,
   плюс то, что FBO не покроет до прихода поставки, плюс мелкий спрос зон FBO, который туда не возим.
   Нужно = квантиль Пуассона от этого спроса; не хватает = нужно − текущий остаток FBS.
"""

import argparse
import json
import math
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

script_dir = Path(__file__).resolve().parent
if str(script_dir) not in sys.path:
    sys.path.insert(0, str(script_dir))
from app_paths import DATA_ROOT  # noqa: E402
repo_root = DATA_ROOT

try:
    from scripts.file_io import atomic_output_path
except ModuleNotFoundError:
    from file_io import atomic_output_path

STOCKS_REPORTS_DIR = "stocks reports"
ORDERS_CACHE_DIR = ".orders_cache"
SETTINGS_FILENAME = "fbo_supply_settings.json"
FBS_ZONE = "FBS (ваш склад)"

MOSCOW = "Москва, МО и Дальние регионы"
SPB = "Санкт-Петербург и СЗО"

DEFAULT_SETTINGS: Dict[str, Any] = {
    "lead_days": 14,
    "cover_days": 30,
    "service_level": 0.9,
    # Меньше этого ожидаемого спроса за цикл в зоне — не возим, покрывает FBS.
    "min_window_demand": 0.5,
    # Склад FBO -> кластеры покупателей, которых он обслуживает. Остальные — FBS.
    "zones": {
        MOSCOW: [
            MOSCOW, "Ярославль", "Тверь", "Воронеж", "Ростов", "Краснодар", "Невинномысск",
            "Махачкала", "Саратов", "Самара", "Казань", "Грузия", "Армения",
        ],
        SPB: [SPB, "Калининград", "Беларусь"],
    },
}

GROWTH_WINDOW_DAYS = 90
GROWTH_MIN_UNITS = 30
GROWTH_LIMITS = (0.5, 2.0)
ARTICLE_SHARE_DAYS = 180
CLUSTER_SHARE_DAYS = 365
ZONE_SHRINK_UNITS = 20


# ---------------------------------------------------------------- настройки

def load_settings(root: Path) -> Dict[str, Any]:
    settings = json.loads(json.dumps(DEFAULT_SETTINGS))
    path = root / SETTINGS_FILENAME
    if path.exists():
        try:
            settings.update(json.loads(path.read_text(encoding="utf-8")))
        except (OSError, ValueError) as exc:
            print(f"⚠️ Не удалось прочитать {SETTINGS_FILENAME}, использую значения по умолчанию: {exc}")
    return settings


def save_settings(root: Path, settings: Dict[str, Any]) -> None:
    (root / SETTINGS_FILENAME).write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")


def cluster_zone_map(settings: Dict[str, Any]) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    for zone, clusters in settings["zones"].items():
        for cluster in clusters:
            mapping[cluster] = zone
    return mapping


# ---------------------------------------------------------------- математика

def poisson_quantile(mean: float, level: float) -> int:
    """Наименьшее k, при котором P(X <= k) >= level для X ~ Poisson(mean)."""
    if mean <= 0:
        return 0
    k, term = 0, math.exp(-mean)
    total = term
    while total < level and k < 10_000:
        k += 1
        term *= mean / k
        total += term
    return k


def units_between(daily: pd.Series, start: date, end: date) -> float:
    """Сумма продаж за [start, end)."""
    return float(daily[(daily.index >= pd.Timestamp(start)) & (daily.index < pd.Timestamp(end))].sum())


def year_ago(day: date) -> date:
    return day - timedelta(days=365)


def store_forecast(daily: pd.Series, history_start: date, today: date, start: date, end: date) -> Dict[str, Any]:
    """
    Прогноз продаж магазина за [start, end): то же окно год назад × рост.
    Если год назад данных нет — средняя за последние 30 дней.
    """
    recent = units_between(daily, today - timedelta(days=GROWTH_WINDOW_DAYS), today)
    previous = units_between(daily, year_ago(today - timedelta(days=GROWTH_WINDOW_DAYS)), year_ago(today))
    last_year = units_between(daily, year_ago(start), year_ago(end))
    if year_ago(start) >= history_start and previous >= GROWTH_MIN_UNITS:
        growth = min(max(recent / previous, GROWTH_LIMITS[0]), GROWTH_LIMITS[1])
        return {
            "method": "season",
            "units": last_year * growth,
            "last_year_units": last_year,
            "growth": growth,
            "recent_units": recent,
            "previous_units": previous,
        }
    rate = units_between(daily, today - timedelta(days=30), today) / 30
    return {
        "method": "recent",
        "units": rate * (end - start).days,
        "last_year_units": last_year,
        "growth": None,
        "recent_units": recent,
        "previous_units": previous,
    }


def article_shares(orders: pd.DataFrame, today: date, window: Tuple[date, date]) -> pd.Series:
    """
    Доля артикула в продажах: последние 180 дней, усреднённая с долей в окне год назад,
    если артикул продавался до конца того окна (иначе — только последние 180 дней).
    """
    def shares(start: date, end: date) -> pd.Series:
        part = orders[(orders["date"] >= pd.Timestamp(start)) & (orders["date"] < pd.Timestamp(end))]
        units = part.groupby("offer_id")["qty"].sum()
        return units / units.sum() if units.sum() > 0 else units

    recent = shares(today - timedelta(days=ARTICLE_SHARE_DAYS), today)
    last_year = shares(year_ago(window[0]), year_ago(window[1]))
    first_sale = orders.groupby("offer_id")["date"].min()
    articles = sorted(set(recent.index) | set(last_year.index))
    raw = {}
    for article in articles:
        r = float(recent.get(article, 0.0))
        existed = first_sale.get(article) is not None and first_sale[article] < pd.Timestamp(year_ago(window[1]))
        raw[article] = (r + float(last_year.get(article, 0.0))) / 2 if existed and len(last_year) else r
    result = pd.Series(raw, dtype=float)
    return result / result.sum() if result.sum() > 0 else result


def zone_shares(orders: pd.DataFrame, today: date, zone_of: Dict[str, str]) -> Tuple[pd.Series, pd.DataFrame]:
    """
    Доли зон в спросе за последний год: по магазину и по артикулам
    (по артикулу сглажено к магазину: (n·доля_артикула + K·доля_магазина) / (n + K)).
    """
    part = orders[orders["date"] >= pd.Timestamp(today - timedelta(days=CLUSTER_SHARE_DAYS))].copy()
    part["zone"] = part["cluster_to"].map(lambda c: zone_of.get(c, FBS_ZONE))
    store = part.groupby("zone")["qty"].sum()
    store = store / store.sum()
    per_article = part.pivot_table(index="offer_id", columns="zone", values="qty", aggfunc="sum", fill_value=0)
    per_article = per_article.reindex(columns=store.index, fill_value=0)
    totals = per_article.sum(axis=1)
    smoothed = per_article.add(store * ZONE_SHRINK_UNITS, axis=1).div(totals + ZONE_SHRINK_UNITS, axis=0)
    return store, smoothed


def build_plan(
    orders: pd.DataFrame,
    stock: pd.DataFrame,
    settings: Dict[str, Any],
    today: date,
    articles: Optional[List[str]] = None,
    ozon_ads: Optional[pd.DataFrame] = None,
    own_stock: Optional[Dict[str, int]] = None,
) -> Dict[str, Any]:
    """
    orders: date (datetime), offer_id, qty, cluster_to — заказы без отмен.
    stock: offer_id, cluster, available, transit, requested — остатки FBO по кластерам.
    ozon_ads: offer_id, cluster, ads — средние продажи в день по оценке Ozon (для сравнения).
    own_stock: свободный остаток FBS по артикулам (для расчёта, сколько не хватает на своём складе).
    """
    lead, cover = int(settings["lead_days"]), int(settings["cover_days"])
    level, min_demand = float(settings["service_level"]), float(settings["min_window_demand"])
    zone_of = cluster_zone_map(settings)
    fbo_zones = list(settings["zones"])

    arrival = today + timedelta(days=lead)
    window = (arrival, arrival + timedelta(days=cover))
    daily = orders.groupby(orders["date"].dt.normalize())["qty"].sum().sort_index()
    history_start = orders["date"].min().date() if len(orders) else today

    window_forecast = store_forecast(daily, history_start, today, *window)
    lead_forecast = store_forecast(daily, history_start, today, today, arrival)
    shares = article_shares(orders, today, window)
    store_zone, article_zone = zone_shares(orders, today, zone_of)

    all_articles = sorted(set(articles or []) | set(shares.index), key=_article_sort_key)
    stock = stock.copy()
    stock["zone"] = stock["cluster"].map(lambda c: zone_of.get(c, FBS_ZONE))
    zone_stock = stock.groupby(["offer_id", "zone"])[["available", "transit", "requested"]].sum()
    ads = None
    if ozon_ads is not None and len(ozon_ads):
        ozon_ads = ozon_ads.copy()
        ozon_ads["zone"] = ozon_ads["cluster"].map(lambda c: zone_of.get(c, FBS_ZONE))
        ads = ozon_ads.groupby(["offer_id", "zone"])["ads"].sum()

    rows = []
    for article in all_articles:
        share = float(shares.get(article, 0.0))
        for zone in fbo_zones:
            if article in article_zone.index and zone in article_zone.columns:
                zshare = float(article_zone.loc[article, zone])
            else:
                zshare = float(store_zone.get(zone, 0.0))
            demand_window = window_forecast["units"] * share * zshare
            demand_lead = lead_forecast["units"] * share * zshare
            available = transit = requested = 0
            if (article, zone) in zone_stock.index:
                available, transit, requested = (int(v) for v in zone_stock.loc[(article, zone)])
            on_arrival = max(0.0, available + transit + requested - demand_lead)
            target = poisson_quantile(demand_window, level) if demand_window >= min_demand else 0
            ship = max(0, math.ceil(target - on_arrival)) if target else 0
            if share == 0:
                note = "Нет продаж — не возим"
            elif demand_window < min_demand:
                note = f"Спрос за цикл {demand_window:.1f} шт. — покрывает FBS"
            elif ship == 0:
                note = "Запаса хватит"
            else:
                note = ""
            rows.append({
                "Артикул": article,
                "Склад FBO": zone,
                "Доля артикула, %": round(share * 100, 1),
                "Доля зоны, %": round(zshare * 100, 1),
                "Прогноз в день, шт.": round(demand_window / cover, 2),
                "Ozon: продажи в день, шт.": round(float(ads.get((article, zone), 0.0)), 2) if ads is not None else None,
                "Спрос за срок поставки, шт.": round(demand_lead, 1),
                "Спрос за цикл, шт.": round(demand_window, 1),
                "Свободно": available,
                "В пути": transit,
                "В заявках": requested,
                "Запас на приходе": round(on_arrival, 1),
                "Цель": target,
                "К поставке": ship,
                "Комментарий": note,
            })
    details = pd.DataFrame(rows)
    fbs = fbs_plan(details, all_articles, shares, store_zone, article_zone, window_forecast, lead_forecast, level, own_stock)
    return {
        "fbs": fbs,
        "sensitivity": service_level_sensitivity(details, fbo_zones, min_demand),
        "today": today,
        "arrival": arrival,
        "window": window,
        "window_forecast": window_forecast,
        "lead_forecast": lead_forecast,
        "store_zone": store_zone,
        "details": details,
        "daily": daily,
        "settings": settings,
    }


def fbs_plan(
    details: pd.DataFrame,
    articles: List[str],
    shares: pd.Series,
    store_zone: pd.Series,
    article_zone: pd.DataFrame,
    window_forecast: Dict[str, Any],
    lead_forecast: Dict[str, Any],
    level: float,
    own_stock: Optional[Dict[str, int]],
) -> pd.DataFrame:
    """
    Сколько держать на своём складе (FBS) на те же дни, что и план FBO: от сегодня до конца окна продаж.
    Спрос FBS = регионы без склада FBO (до прихода + в окне)
              + то, что FBO не покроет до прихода поставки (спрос за срок поставки сверх остатка FBO)
              + мелкий спрос зон FBO, куда не везём (сверх запаса на приходе).
    """
    by_article = {article: part for article, part in details.groupby("Артикул", sort=False)}
    rows = []
    for article in articles:
        share = float(shares.get(article, 0.0))
        if article in article_zone.index and FBS_ZONE in article_zone.columns:
            zshare = float(article_zone.loc[article, FBS_ZONE])
        else:
            zshare = float(store_zone.get(FBS_ZONE, 0.0))
        own_lead = lead_forecast["units"] * share * zshare
        own_window = window_forecast["units"] * share * zshare
        spill_lead = spill_window = 0.0
        part = by_article.get(article)
        if part is not None:
            for _, row in part.iterrows():
                fbo_stock = row["Свободно"] + row["В пути"] + row["В заявках"]
                spill_lead += max(0.0, row["Спрос за срок поставки, шт."] - fbo_stock)
                if row["Цель"] == 0:  # в зону не везём — её спрос сверх запаса уйдёт со своего склада
                    spill_window += max(0.0, row["Спрос за цикл, шт."] - row["Запас на приходе"])
        demand = own_lead + own_window + spill_lead + spill_window
        need = poisson_quantile(demand, level) if demand > 0 else 0
        stock = None if own_stock is None else own_stock.get(article)
        shortage = max(0, need - (stock or 0))
        if share == 0:
            note = "Нет продаж"
        elif need == 0:
            note = "Спроса на FBS почти нет"
        elif shortage == 0:
            note = "Остатка хватит"
        else:
            note = ""
        rows.append({
            "Артикул": article,
            "Регионы FBS до прихода, шт.": round(own_lead, 1),
            "Регионы FBS в окне, шт.": round(own_window, 1),
            "FBO не хватит до прихода, шт.": round(spill_lead, 1),
            "Мелкий спрос зон FBO, шт.": round(spill_window, 1),
            "Спрос на FBS, шт.": round(demand, 1),
            "Нужно на FBS": need,
            "Остаток FBS": stock,
            "Не хватает на FBS": shortage,
            "Комментарий": note,
        })
    return pd.DataFrame(rows)


SENSITIVITY_LEVELS = (0.7, 0.8, 0.9, 0.95)


def service_level_sensitivity(details: pd.DataFrame, zones: List[str], min_demand: float) -> pd.DataFrame:
    """Сколько штук везти в каждую зону при разных уровнях сервиса (при тех же остатках)."""
    rows = []
    for level in SENSITIVITY_LEVELS:
        row: Dict[str, Any] = {"Уровень сервиса": f"{level:.0%}"}
        for zone in zones:
            part = details[details["Склад FBO"] == zone]
            row[zone] = sum(
                max(0, math.ceil(poisson_quantile(r["Спрос за цикл, шт."], level) - r["Запас на приходе"]))
                for _, r in part.iterrows() if r["Спрос за цикл, шт."] >= min_demand
            )
        rows.append(row)
    return pd.DataFrame(rows)


def _article_sort_key(article: str):
    return (0, int(article)) if str(article).isdigit() else (1, str(article))


# ---------------------------------------------------------------- данные Ozon

def _month_starts(start: date, end: date) -> List[date]:
    months, cur = [], start.replace(day=1)
    while cur <= end:
        months.append(cur)
        cur = (cur.replace(day=28) + timedelta(days=4)).replace(day=1)
    return months


def load_orders(root: Path, today: date, months_back: int = 24) -> pd.DataFrame:
    """
    Заказы FBO и FBS за ~2 года с кэшем по месяцам: закрытые месяцы старше прошлого
    берутся из кэша, текущий и прошлый — всегда заново.
    """
    import Monthly_sales_report as monthly
    from recommended_prices import _normalize_offer_id

    cache_dir = root / STOCKS_REPORTS_DIR / ORDERS_CACHE_DIR
    cache_dir.mkdir(parents=True, exist_ok=True)
    start = (today.replace(day=1) - timedelta(days=31 * months_back)).replace(day=1)
    fresh_from = (today.replace(day=1) - timedelta(days=1)).replace(day=1)
    rows: List[Dict[str, Any]] = []
    with monthly.create_session() as session:
        for month in _month_starts(start, today):
            cache = cache_dir / f"{month:%Y-%m}.json"
            if month < fresh_from and cache.exists():
                rows.extend(json.loads(cache.read_text(encoding="utf-8")))
                continue
            month_end = (month.replace(day=28) + timedelta(days=4)).replace(day=1)
            since, to = f"{month}T00:00:00Z", f"{min(month_end, today + timedelta(days=1))}T00:00:00Z"
            month_rows = []
            for fetch in (monthly.get_fbo_orders, monthly.get_orders):
                for posting in fetch(since, to, session=session):
                    if posting.get("status") == "cancelled":
                        continue
                    cluster = (posting.get("financial_data") or {}).get("cluster_to")
                    day = (posting.get("in_process_at") or posting.get("created_at") or "")[:10]
                    for product in posting.get("products") or []:
                        month_rows.append({
                            "date": day,
                            "offer_id": _normalize_offer_id(str(product.get("offer_id") or "")),
                            "qty": int(product.get("quantity") or 1),
                            "cluster_to": cluster,
                        })
            cache.write_text(json.dumps(month_rows, ensure_ascii=False), encoding="utf-8")
            print(f"  {month:%m.%Y}: {sum(r['qty'] for r in month_rows)} шт.", flush=True)
            rows.extend(month_rows)
    orders = pd.DataFrame(rows, columns=["date", "offer_id", "qty", "cluster_to"])
    orders["date"] = pd.to_datetime(orders["date"], errors="coerce")
    orders = orders.dropna(subset=["date"])
    return orders[orders["offer_id"] != ""].reset_index(drop=True)


RETRY_DELAYS = (5, 15, 30)  # секунды между повторами, если Ozon временно не отвечает


class OzonUnavailable(RuntimeError):
    pass


def ozon_post(url: str, payload: Dict[str, Any], sleep=None) -> Dict[str, Any]:
    """
    POST в Ozon с повторами: сбои 5xx, 429 и обрывы связи у Ozon бывают временными (метод остатков иногда
    отвечает 500 и через минуту работает). После последней попытки — OzonUnavailable с понятным текстом.
    """
    import time
    import requests
    from recommended_prices import OZON_HEADERS

    sleep = sleep or time.sleep
    method = url.split("ozon.ru")[-1]
    last_error = ""
    for attempt in range(len(RETRY_DELAYS) + 1):
        try:
            response = requests.post(url, headers=OZON_HEADERS, json=payload, timeout=60)
        except requests.RequestException as exc:
            last_error = f"нет связи ({exc.__class__.__name__})"
        else:
            if response.status_code == 200:
                return response.json()
            if response.status_code < 500 and response.status_code != 429:
                raise OzonUnavailable(f"Ozon отклонил запрос {method}: HTTP {response.status_code} {response.text[:200]}")
            last_error = f"HTTP {response.status_code}"
        if attempt < len(RETRY_DELAYS):
            print(f"  Ozon не ответил на {method} ({last_error}) — повтор через {RETRY_DELAYS[attempt]} с...", flush=True)
            sleep(RETRY_DELAYS[attempt])
    raise OzonUnavailable(f"Ozon временно не отвечает на {method} ({last_error}) после {len(RETRY_DELAYS) + 1} попыток. "
                          "Это сбой на стороне Ozon — повторите расчёт через несколько минут.")


def load_cluster_stock(skus_by_offer: Dict[str, int]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Остатки FBO и оценка продаж Ozon по кластерам (/v1/analytics/stocks)."""
    from recommended_prices import _normalize_offer_id

    offer_by_sku = {sku: offer for offer, sku in skus_by_offer.items()}
    stock_rows, ads_rows, seen_ads = [], [], set()
    skus = sorted(offer_by_sku)
    for i in range(0, len(skus), 100):
        body = ozon_post("https://api-seller.ozon.ru/v1/analytics/stocks", {"skus": [str(sku) for sku in skus[i:i + 100]]})
        for item in body.get("items") or []:
            offer = offer_by_sku.get(int(item["sku"])) or _normalize_offer_id(str(item.get("offer_id") or ""))
            cluster = item.get("cluster_name") or ""
            stock_rows.append({
                "offer_id": offer,
                "cluster": cluster,
                "available": int(item.get("available_stock_count") or 0),
                "transit": int(item.get("transit_stock_count") or 0),
                "requested": int(item.get("requested_stock_count") or 0),
            })
            # ads_cluster повторяется на каждом складе кластера — берём один раз.
            if (offer, cluster) not in seen_ads:
                seen_ads.add((offer, cluster))
                ads_rows.append({"offer_id": offer, "cluster": cluster, "ads": float(item.get("ads_cluster") or 0)})
    stock = pd.DataFrame(stock_rows, columns=["offer_id", "cluster", "available", "transit", "requested"])
    ads = pd.DataFrame(ads_rows, columns=["offer_id", "cluster", "ads"])
    return stock, ads


def load_own_stock(offer_ids: List[str]) -> Dict[str, int]:
    """Свободный остаток FBS (ваш склад) по данным Ozon."""
    from recommended_prices import _normalize_offer_id

    result: Dict[str, int] = {}
    for i in range(0, len(offer_ids), 1000):
        cursor = ""
        while True:
            payload: Dict[str, Any] = {"filter": {"offer_id": offer_ids[i:i + 1000], "visibility": "ALL"}, "limit": 1000}
            if cursor:
                payload["cursor"] = cursor
            data = ozon_post("https://api-seller.ozon.ru/v4/product/info/stocks", payload)
            items = data.get("items") or []
            for item in items:
                fbs = sum(
                    int(s.get("present") or 0) - int(s.get("reserved") or 0)
                    for s in item.get("stocks") or [] if s.get("type") == "fbs"
                )
                result[_normalize_offer_id(str(item.get("offer_id") or ""))] = max(0, fbs)
            cursor = data.get("cursor") or ""
            if not items or not cursor or len(items) < 1000:
                break
    return result


# ---------------------------------------------------------------- отчёт

def parse_unit_cost(value: Any) -> Optional[float]:
    """Себестоимость из costs.xlsx (допускается запятая); пусто или не больше нуля — нет данных."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    try:
        number = float(str(value).strip().replace(",", ".").replace(" ", ""))
    except ValueError:
        return None
    return number if number > 0 else None


def format_rub(value: float) -> str:
    return f"{value:,.0f} ₽".replace(",", " ")


def supply_costs(details: pd.DataFrame, unit_costs: Dict[str, float], zones: List[str]) -> Dict[str, Any]:
    """Себестоимость поставки: всего, по складам и список артикулов без себестоимости."""
    shipped = details[details["К поставке"] > 0]
    by_zone = {zone: 0.0 for zone in zones}
    missing = set()
    for _, row in shipped.iterrows():
        unit = unit_costs.get(str(row["Артикул"]))
        if unit is None:
            missing.add(str(row["Артикул"]))
            continue
        by_zone[row["Склад FBO"]] = by_zone.get(row["Склад FBO"], 0.0) + unit * int(row["К поставке"])
    return {
        "total": round(sum(by_zone.values()), 2),
        "by_zone": {zone: round(value, 2) for zone, value in by_zone.items()},
        "missing_articles": sorted(missing, key=_article_sort_key),
    }


def fbs_costs(fbs: pd.DataFrame, unit_costs: Dict[str, float]) -> Dict[str, Any]:
    """Себестоимость того, чего не хватает на своём складе."""
    total, missing = 0.0, []
    for _, row in fbs[fbs["Не хватает на FBS"] > 0].iterrows():
        unit = unit_costs.get(str(row["Артикул"]))
        if unit is None:
            missing.append(str(row["Артикул"]))
        else:
            total += unit * int(row["Не хватает на FBS"])
    return {"total": round(total, 2), "missing_articles": missing}


def _fbs_sheet(fbs: pd.DataFrame, unit_costs: Dict[str, float]) -> pd.DataFrame:
    frame = fbs[(fbs["Нужно на FBS"] > 0) | (fbs["Остаток FBS"].fillna(0) > 0)].copy()
    frame["Себестоимость нехватки, ₽"] = [
        round(unit_costs[str(a)] * int(s), 2) if str(a) in unit_costs and s else None
        for a, s in zip(frame["Артикул"], frame["Не хватает на FBS"])
    ]
    frame["_key"] = frame["Артикул"].map(_article_sort_key)
    frame = frame.sort_values(["Не хватает на FBS", "Нужно на FBS", "_key"], ascending=[False, False, True]).drop(columns="_key")
    total = {column: frame[column].fillna(0).sum() for column in frame.columns if column not in ("Артикул", "Комментарий")}
    total.update({"Артикул": "ИТОГО", "Комментарий": ""})
    for column in ("Регионы FBS до прихода, шт.", "Регионы FBS в окне, шт.", "FBO не хватит до прихода, шт.", "Мелкий спрос зон FBO, шт.", "Спрос на FBS, шт."):
        total[column] = round(float(total[column]), 1)
    return pd.concat([frame, pd.DataFrame([total])], ignore_index=True)


def _supply_sheet(details: pd.DataFrame, own_stock: Dict[str, int], unit_costs: Dict[str, float], zones: List[str],
                  fbs: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    pivot = details.pivot_table(index="Артикул", columns="Склад FBO", values="К поставке", aggfunc="sum", fill_value=0)
    pivot = pivot.reindex(columns=zones, fill_value=0)
    pivot["Итого"] = pivot.sum(axis=1)
    pivot["Себестоимость за шт., ₽"] = [unit_costs.get(str(article)) for article in pivot.index]
    pivot["Себестоимость поставки, ₽"] = [
        round(unit * total, 2) if unit is not None else None
        for unit, total in zip(pivot["Себестоимость за шт., ₽"], pivot["Итого"])
    ]
    pivot["Остаток FBS в Ozon"] = [own_stock.get(article) for article in pivot.index]
    if fbs is not None and len(fbs):
        fbs_by_article = fbs.set_index("Артикул")
        pivot["Нужно на FBS"] = [int(fbs_by_article["Нужно на FBS"].get(article, 0)) for article in pivot.index]
        pivot["Не хватает на FBS"] = [int(fbs_by_article["Не хватает на FBS"].get(article, 0)) for article in pivot.index]
    pivot = pivot.reset_index()
    pivot["_key"] = pivot["Артикул"].map(_article_sort_key)
    pivot = pivot.sort_values(["Итого", "_key"], ascending=[False, True]).drop(columns="_key")
    total = {
        "Артикул": "ИТОГО",
        **{col: pivot[col].sum() for col in [*zones, "Итого"]},
        "Себестоимость за шт., ₽": "",
        "Себестоимость поставки, ₽": round(pivot["Себестоимость поставки, ₽"].fillna(0).sum(), 2),
        "Остаток FBS в Ozon": pivot["Остаток FBS в Ozon"].fillna(0).sum(),
    }
    for column in ("Нужно на FBS", "Не хватает на FBS"):
        if column in pivot.columns:
            total[column] = pivot[column].sum()
    return pd.concat([pivot, pd.DataFrame([total])], ignore_index=True)


def _season_sheet(daily: pd.Series) -> pd.DataFrame:
    monthly = daily.groupby(daily.index.to_period("M")).sum()
    rows = []
    for period, units in monthly.items():
        prev = monthly.get(period - 12)
        rows.append({
            "Месяц": period.strftime("%m.%Y"),
            "Продажи, шт.": int(units),
            "Год назад, шт.": int(prev) if prev is not None else None,
            "Рост к прошлому году": round(units / prev, 2) if prev else None,
        })
    return pd.DataFrame(rows)


def _zone_sheet(orders: pd.DataFrame, today: date, settings: Dict[str, Any]) -> pd.DataFrame:
    zone_of = cluster_zone_map(settings)
    part = orders[orders["date"] >= pd.Timestamp(today - timedelta(days=CLUSTER_SHARE_DAYS))]
    units = part.groupby("cluster_to")["qty"].sum().sort_values(ascending=False)
    return pd.DataFrame({
        "Кластер покупателя": units.index,
        "Обслуживает": [zone_of.get(c, FBS_ZONE) for c in units.index],
        "Продажи за год, шт.": units.values.astype(int),
        "Доля спроса, %": (units / units.sum() * 100).round(1).values,
    })


def _explain_sheet(plan: Dict[str, Any], costs: Dict[str, Any]) -> pd.DataFrame:
    s, wf, lf = plan["settings"], plan["window_forecast"], plan["lead_forecast"]
    cost_parts = ", ".join(f"{zone}: {format_rub(value)}" for zone, value in costs["by_zone"].items())
    cost_missing = f" Без себестоимости в costs.xlsx, не учтены: {', '.join(costs['missing_articles'])}." if costs["missing_articles"] else ""
    w0, w1 = plan["window"]
    fmt = lambda d: d.strftime("%d.%m.%Y")
    lines = [
        ("Параметры", f"Срок поставки {s['lead_days']} дн., цикл {s['cover_days']} дн., уровень сервиса {s['service_level']:.0%}."),
        ("Окно поставки", f"Товар придёт ~{fmt(plan['arrival'])} и должен продаваться до {fmt(w1 - timedelta(days=1))}."),
    ]
    if wf["method"] == "season":
        lines += [
            ("Сезонность", f"За то же окно год назад ({fmt(year_ago(w0))}–{fmt(year_ago(w1) - timedelta(days=1))}) продано {wf['last_year_units']:.0f} шт."),
            ("Рост", f"Последние 90 дней: {wf['recent_units']:.0f} шт. против {wf['previous_units']:.0f} шт. год назад — коэффициент {wf['growth']:.2f} (ограничен 0,5–2)."),
            ("Прогноз магазина", f"{wf['last_year_units']:.0f} × {wf['growth']:.2f} = {wf['units']:.0f} шт. за цикл."),
        ]
    else:
        lines.append(("Прогноз магазина", f"Год назад данных мало — по средней за 30 дней: {wf['units']:.0f} шт. за цикл."))
    lines += [
        ("За срок поставки", f"До прихода товара магазин продаст ~{lf['units']:.0f} шт. — они спишутся с текущего запаса."),
        ("Доля артикула", "Последние 180 дней, усреднённая с долей в том же окне год назад (для новых артикулов — только 180 дней)."),
        ("Зоны складов", "Склад FBO продаёт не только своему кластеру: по истории ~3/4 отгрузок со складов Москвы и СПб ушли в другие регионы. "
                         "Поэтому каждый кластер покупателя закреплён за складом (лист «Зоны»), остальные покрывает FBS."),
        ("Доля зоны", f"По артикулу сглажена к доле магазина: (n·доля артикула + {ZONE_SHRINK_UNITS}·доля магазина) / (n + {ZONE_SHRINK_UNITS}), n — продажи артикула за год."),
        ("Запас на приходе", "Свободно + в пути + в заявках на складах зоны − спрос зоны за срок поставки."),
        ("Цель", f"Сколько штук нужно, чтобы с вероятностью {s['service_level']:.0%} не уйти в ноль за цикл (распределение Пуассона — подходит для штучных продаж)."),
        ("К поставке", "Цель − запас на приходе. Если спрос зоны за цикл меньше "
                       f"{s['min_window_demand']} шт. — не возим, покрывает FBS."),
        *[
            (f"Итог: {zone}", f"Ожидаемый спрос зоны за цикл {part['Спрос за цикл, шт.'].sum():.0f} шт., "
                              f"запас на приходе {part['Запас на приходе'].sum():.0f} шт., к поставке {int(part['К поставке'].sum())} шт. — "
                              "разница со спросом в основном страховой запас (см. лист «Уровень сервиса»).")
            for zone, part in plan["details"].groupby("Склад FBO", sort=False)
        ],
        ("Себестоимость поставки", f"{format_rub(costs['total'])} — количество к поставке × себестоимость из costs.xlsx ({cost_parts}).{cost_missing}"),
        ("Свой склад (FBS)", f"На те же дни — с {fmt(plan['today'])} по {fmt(w1 - timedelta(days=1))}: спрос регионов без склада FBO (до прихода и в окне) "
                             "+ то, что FBO не покроет до прихода поставки + мелкий спрос зон FBO, куда не везём. "
                             f"Нужно = запас с вероятностью {s['service_level']:.0%} не уйти в ноль; не хватает = нужно − текущий остаток FBS (лист «FBS»)."),
        ("Уровень сервиса", "Запас считается по каждому артикулу отдельно, поэтому при штучных продажах страховка заметна. "
                            "Если на FBO закончится, заказ уйдёт FBS со своего склада, только медленнее — можно выбрать уровень ниже."),
        ("Ozon для сравнения", "Колонка «Ozon: продажи в день» — оценка самого Ozon по кластерам зоны. Наша считается по сезону, Ozon — по недавним продажам."),
        ("Ограничения", "Дни, когда товара не было в наличии, занижают историю продаж: истории остатков у Ozon нет. "
                        "Если артикул долго отсутствовал, его прогноз может быть ниже реального спроса."),
    ]
    return pd.DataFrame(lines, columns=["Шаг", "Как считается"])


def write_report(plan: Dict[str, Any], orders: pd.DataFrame, own_stock: Dict[str, int], unit_costs: Dict[str, float], out_path: Path) -> None:
    from openpyxl.utils import get_column_letter

    zones = list(plan["settings"]["zones"])
    sheets = {
        "Поставка": _supply_sheet(plan["details"], own_stock, unit_costs, zones, plan.get("fbs")),
        "Расчёт": plan["details"],
        "FBS": _fbs_sheet(plan["fbs"], unit_costs) if plan.get("fbs") is not None and len(plan["fbs"]) else pd.DataFrame(),
        "Как считается": _explain_sheet(plan, supply_costs(plan["details"], unit_costs, zones)),
        "Уровень сервиса": plan["sensitivity"],
        "Сезонность": _season_sheet(plan["daily"]),
        "Зоны": _zone_sheet(orders, plan["today"], plan["settings"]),
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


def plan_summary(plan: Dict[str, Any], orders: pd.DataFrame, own_stock: Dict[str, int], unit_costs: Dict[str, float]) -> Dict[str, Any]:
    """Сводка расчёта для вкладки «Поставки» в интерфейсе (рядом с Excel, тот же набор цифр)."""
    settings, details = plan["settings"], plan["details"]
    zones = list(settings["zones"])
    wf = plan["window_forecast"]
    costs = supply_costs(details, unit_costs, zones)
    zone_rows = []
    for zone in zones:
        part = details[details["Склад FBO"] == zone]
        zone_rows.append({
            "name": zone,
            "ship": int(part["К поставке"].sum()),
            "demand": round(float(part["Спрос за цикл, шт."].sum()), 1),
            "on_arrival": round(float(part["Запас на приходе"].sum()), 1),
            "articles": int((part["К поставке"] > 0).sum()),
            "cost": costs["by_zone"].get(zone, 0.0),
        })
    supply = []
    for article, part in details.groupby("Артикул", sort=False):
        by_zone = {row["Склад FBO"]: int(row["К поставке"]) for _, row in part.iterrows()}
        total = sum(by_zone.values())
        if total:
            supply.append({
                "article": str(article),
                "by_zone": by_zone,
                "total": total,
                "per_day": round(float(part["Прогноз в день, шт."].sum()), 2),
                "own_stock": own_stock.get(str(article)),
                "unit_cost": unit_costs.get(str(article)),
                "cost": round(unit_costs[str(article)] * total, 2) if str(article) in unit_costs else None,
            })
    supply.sort(key=lambda row: (-row["total"], _article_sort_key(row["article"])))
    zone_table = _zone_sheet(orders, plan["today"], settings)
    sensitivity = plan["sensitivity"]
    fbs = plan.get("fbs")
    fbs_summary = None
    if fbs is not None and len(fbs):
        fbs_rows = []
        for _, row in fbs[(fbs["Нужно на FBS"] > 0) | (fbs["Остаток FBS"].fillna(0) > 0)].iterrows():
            article = str(row["Артикул"])
            stock_value = row["Остаток FBS"]
            fbs_rows.append({
                "article": article,
                "own_lead": float(row["Регионы FBS до прихода, шт."]),
                "own_window": float(row["Регионы FBS в окне, шт."]),
                "spill": round(float(row["FBO не хватит до прихода, шт."]) + float(row["Мелкий спрос зон FBO, шт."]), 1),
                "demand": float(row["Спрос на FBS, шт."]),
                "need": int(row["Нужно на FBS"]),
                "stock": None if pd.isna(stock_value) else int(stock_value),
                "shortage": int(row["Не хватает на FBS"]),
                "cost": round(unit_costs[article] * int(row["Не хватает на FBS"]), 2) if article in unit_costs and row["Не хватает на FBS"] else None,
            })
        fbs_rows.sort(key=lambda r: (-r["shortage"], -r["need"], _article_sort_key(r["article"])))
        shortage_cost = fbs_costs(fbs, unit_costs)
        fbs_summary = {
            "horizon": [plan["today"].isoformat(), plan["window"][1].isoformat()],
            "need": int(fbs["Нужно на FBS"].sum()),
            "stock": int(fbs["Остаток FBS"].fillna(0).sum()),
            "shortage": int(fbs["Не хватает на FBS"].sum()),
            "demand": round(float(fbs["Спрос на FBS, шт."].sum()), 1),
            "shortage_cost": shortage_cost["total"],
            "missing_cost_articles": shortage_cost["missing_articles"],
            "rows": fbs_rows,
        }
    return {
        "fbs": fbs_summary,
        "created": plan["today"].isoformat(),
        "settings": {key: settings[key] for key in ("lead_days", "cover_days", "service_level")},
        "zones": zone_rows,
        "fbs_zone": FBS_ZONE,
        "arrival": plan["arrival"].isoformat(),
        "window": [plan["window"][0].isoformat(), plan["window"][1].isoformat()],
        "base_window": [year_ago(plan["window"][0]).isoformat(), year_ago(plan["window"][1]).isoformat()],
        "forecast": {key: wf[key] for key in ("method", "units", "last_year_units", "growth")},
        "lead_units": plan["lead_forecast"]["units"],
        "sensitivity": [
            {"level": level, "by_zone": {zone: int(row[zone]) for zone in zones}}
            for level, (_, row) in zip(SENSITIVITY_LEVELS, sensitivity.iterrows())
        ],
        "season": [
            {"month": period.strftime("%Y-%m"), "units": int(units)}
            for period, units in plan["daily"].groupby(plan["daily"].index.to_period("M")).sum().items()
        ],
        "clusters": [
            {"cluster": row["Кластер покупателя"], "zone": row["Обслуживает"], "units": int(row["Продажи за год, шт."]), "share": float(row["Доля спроса, %"])}
            for _, row in zone_table.iterrows()
        ],
        "supply": supply,
        "cost": {"total": costs["total"], "missing_articles": costs["missing_articles"]},
    }


# ---------------------------------------------------------------- запуск

def run(root: Path, overrides: Dict[str, Any]) -> bool:
    from recommended_prices import load_costs_df, COSTS_FILENAME, _artikul_normalize, _normalize_offer_id
    from recommended_prices import get_product_id_mapping, get_product_skus

    settings = load_settings(root)
    settings.update({k: v for k, v in overrides.items() if v is not None})
    save_settings(root, settings)
    today = date.today()

    articles: List[str] = []
    unit_costs: Dict[str, float] = {}
    costs_path = root / COSTS_FILENAME
    if costs_path.exists():
        df, key_col, cost_col = load_costs_df(costs_path)
        for raw_article, raw_cost in zip(df[key_col], df[cost_col]):
            article = _artikul_normalize(raw_article)
            if not article:
                continue
            article = _normalize_offer_id(article)
            articles.append(article)
            cost = parse_unit_cost(raw_cost)
            if cost is not None:
                unit_costs[article] = cost

    print("· Загружаю историю заказов FBO и FBS (закрытые месяцы — из кэша)...")
    orders = load_orders(root, today)
    if orders.empty:
        print("❌ Ozon не вернул заказов — считать не из чего.")
        return False
    articles = sorted(set(articles) | set(orders["offer_id"]), key=_article_sort_key)

    print("· Загружаю остатки FBO по кластерам...")
    pid_to_offer = get_product_id_mapping(articles)
    skus = get_product_skus(sorted(pid_to_offer))
    skus_by_offer = {pid_to_offer[pid]: sku for pid, sku in skus.items() if pid in pid_to_offer}
    try:
        stock, ads = load_cluster_stock(skus_by_offer)
        own_stock = load_own_stock(articles)
    except OzonUnavailable as exc:
        print(f"❌ {exc}")
        return False

    plan = build_plan(orders, stock, settings, today, articles=articles, ozon_ads=ads, own_stock=own_stock)
    wf = plan["window_forecast"]
    w0, w1 = plan["window"]
    print(f"· Окно поставки: {w0:%d.%m}–{(w1 - timedelta(days=1)):%d.%m}. Прогноз магазина за цикл: {wf['units']:.0f} шт.", end="")
    print(f" (год назад {wf['last_year_units']:.0f} × рост {wf['growth']:.2f})" if wf["method"] == "season" else " (по последним 30 дням)")

    details = plan["details"]
    for zone in settings["zones"]:
        part = details[details["Склад FBO"] == zone]
        print(
            f"· {zone}: к поставке {int(part['К поставке'].sum())} шт. по {int((part['К поставке'] > 0).sum())} артикулам "
            f"(ожидаемый спрос за цикл {part['Спрос за цикл, шт.'].sum():.0f} шт., остальное — страховой запас)"
        )
    costs = supply_costs(details, unit_costs, list(settings["zones"]))
    missing_note = f" (без себестоимости в costs.xlsx, не учтены: {', '.join(costs['missing_articles'])})" if costs["missing_articles"] else ""
    print(f"· Себестоимость поставки: {format_rub(costs['total'])}{missing_note}")
    fbs = plan["fbs"]
    if len(fbs):
        print(f"· Свой склад (FBS) по {(w1 - timedelta(days=1)):%d.%m}: нужно {int(fbs['Нужно на FBS'].sum())} шт., "
              f"есть {int(fbs['Остаток FBS'].fillna(0).sum())} шт., не хватает {int(fbs['Не хватает на FBS'].sum())} шт. "
              f"на {format_rub(fbs_costs(fbs, unit_costs)['total'])} по себестоимости")
    sensitivity = plan["sensitivity"].set_index("Уровень сервиса")
    print("· Если выбрать другой уровень сервиса, всего к поставке: " + ", ".join(
        f"{level} — {int(row.sum())} шт." for level, row in sensitivity.iterrows()
    ))

    out_dir = root / STOCKS_REPORTS_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"Поставка FBO {today:%d.%m.%Y}.xlsx"
    write_report(plan, orders, own_stock, unit_costs, out_path)
    out_path.with_suffix(".json").write_text(
        json.dumps(plan_summary(plan, orders, own_stock, unit_costs), ensure_ascii=False, indent=1), encoding="utf-8"
    )
    print(f"✅ Отчёт сохранён: {out_path}")
    return True


def main() -> None:
    parser = argparse.ArgumentParser(description="План поставки FBO по кластерам с учётом сезонности.")
    parser.add_argument("--lead-days", type=int, help="Дней от решения до продажи на складе Ozon")
    parser.add_argument("--cover-days", type=int, help="На сколько дней продаж рассчитана поставка")
    parser.add_argument("--service-level", type=float, help="Вероятность не уйти в ноль за цикл, 0.5–0.99")
    args = parser.parse_args()
    if args.service_level is not None and not 0.5 <= args.service_level <= 0.99:
        parser.error("Уровень сервиса должен быть от 0.5 до 0.99.")
    overrides = {"lead_days": args.lead_days, "cover_days": args.cover_days, "service_level": args.service_level}
    sys.exit(0 if run(repo_root, overrides) else 1)


if __name__ == "__main__":
    main()
