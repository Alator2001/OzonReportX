# -*- coding: utf-8 -*-
"""
Акции WB: календарь с бустом и порогами участия и баланс «буст ↔ скидка ↔ маржа» по каждому артикулу.
Только чтение — участие в акциях программа не меняет.

Буст даётся за акцию, а не за артикул: все участники получают одинаковое поднятие в поиске, а его уровень зависит
от доли вашего ассортимента (с остатком) в акции — поле ranging в /calendar/promotions/details.

Какие товары в автоакции и по какой цене, API не отдаёт (/calendar/promotions/nomenclatures для автоакций — 422).
Цену в акции оцениваем по скидкам: автоакция на старте поднимает скидку продавца, поэтому наибольшая скидка
артикула за последние 30 дней — это скидка акции. Источники: история скидок (wb reports/price_history.json, пишется
при каждой синхронизации цен) и скидки в заказах за 30 дней (статистика WB, discountPercent на момент заказа).

Окупаемость буста: прибыль на минимальной цене / прибыль по цене акции — во сколько раз акция должна поднять продажи,
чтобы заработать столько же, сколько без неё. Прибыль с продажи считается тем же расчётом, что и раздел «Цены»
(scripts/wb_prices.py).
"""

import json
import math
import time
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests

from scripts.file_io import atomic_output_path
from scripts.wb_prices import unit_profit

CALENDAR_URL = "https://dp-calendar-api.wildberries.ru/api/v1/calendar/promotions"
DETAILS_URL = "https://dp-calendar-api.wildberries.ru/api/v1/calendar/promotions/details"
ORDERS_URL = "https://statistics-api.wildberries.ru/api/v1/supplier/orders"
PROMOTIONS_FILE = Path("wb reports") / "promotions.json"
DAILY_FILE = Path("wb reports") / "daily_stats.json"
FUNNEL_HISTORY_URL = "https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products/history"
FUNNEL_DAYS, FUNNEL_CHUNK, FUNNEL_PAUSE = 7, 20, 21  # WB: неделя по дням, до 20 товаров в запросе, 3 запроса в минуту
MIN_DAYS = 3  # меньше дней в акции или без неё — сравнивать рано
HISTORY_FILE = Path("wb reports") / "price_history.json"
LOOKBACK_DAYS = 30
CALENDAR_BACK_DAYS, CALENDAR_AHEAD_DAYS = 3, 45
TIMEOUT = 60
# Вердикт по окупаемости: буст должен поднять продажи не больше чем во столько раз, чтобы акция была выгодна.
KEEP_UP_TO, WATCH_UP_TO = 1.5, 3.0


class WbPromotionsError(RuntimeError):
    pass


def _get(url: str, token: str, params) -> Any:
    response = requests.get(url, headers={"Authorization": token}, params=params, timeout=TIMEOUT)
    if response.status_code != 200:
        raise WbPromotionsError(f"WB вернул HTTP {response.status_code} на {url.split('.ru')[-1]}")
    return response.json()


# ── календарь ──────────────────────────────────────────────────────────────

def load_promotions(token: str, now: Optional[datetime] = None) -> List[Dict[str, Any]]:
    """Акции, которые идут сейчас или начнутся в ближайшие 45 дней, с подробностями."""
    now = now or datetime.now()
    params = {"startDateTime": (now - timedelta(days=CALENDAR_BACK_DAYS)).strftime("%Y-%m-%dT00:00:00Z"),
              "endDateTime": (now + timedelta(days=CALENDAR_AHEAD_DAYS)).strftime("%Y-%m-%dT23:59:59Z"),
              "allPromo": "true", "limit": 1000, "offset": 0}
    listed = (_get(CALENDAR_URL, token, params).get("data") or {}).get("promotions") or []
    ids = [promo["id"] for promo in listed]
    details = {}
    for start in range(0, len(ids), 100):
        chunk = ids[start:start + 100]
        body = _get(DETAILS_URL, token, [("promotionIDs", promo_id) for promo_id in chunk])
        for promo in (body.get("data") or {}).get("promotions") or []:
            details[promo["id"]] = promo
    result = []
    for promo in listed:
        merged = dict(promo, **details.get(promo["id"], {}))
        if _date(merged.get("endDateTime")) and _date(merged["endDateTime"]) < now.date():
            continue  # уже закончилась
        result.append(merged)
    return sorted(result, key=lambda promo: (promo.get("startDateTime") or "", promo.get("name") or ""))


def _date(text) -> Optional[date]:
    try:
        return datetime.fromisoformat(str(text).replace("Z", "+00:00")).date()
    except ValueError:
        return None


def tiers(promo: Dict[str, Any]) -> Dict[str, Any]:
    """Текущий буст и что нужно для следующего уровня.

    Доля участия считается по товарам с остатком (participationPercentage WB). Возвращает boost (текущий, %),
    next_boost и need — сколько товаров с остатком добавить в акцию до следующего уровня (None — уровень максимальный)."""
    inside = int(promo.get("inPromoActionLeftovers") or 0)
    total = inside + int(promo.get("notInPromoActionLeftovers") or 0)
    share = promo.get("participationPercentage")
    share = float(share) if share is not None else (inside / total * 100 if total else 0.0)
    steps = sorted((float(r.get("participationRate") or 0), float(r.get("boost") or 0))
                   for r in promo.get("ranging") or [])
    current = max((boost for rate, boost in steps if share >= rate and inside), default=0.0)
    upcoming = [(rate, boost) for rate, boost in steps if boost > current]
    result = {"share": share, "inside": inside, "total": total, "boost": current, "next_boost": None, "need": None}
    if upcoming:
        rate, boost = upcoming[0]
        # Наименьшее n: (inside + n) / total ≥ rate %. Хотя бы один товар — иначе уровень уже был бы набран.
        need = max(1, math.ceil(rate * total / 100 - 1e-9) - inside)
        result.update(next_boost=boost, next_rate=rate, need=need)
    return result


# ── история скидок ─────────────────────────────────────────────────────────

def record_history(root, prices: Dict[str, Dict[str, Any]], when: Optional[datetime] = None) -> Dict[str, Any]:
    """Дописать в историю скидок те артикулы, у которых цена или скидка изменились с прошлой записи."""
    when = when or datetime.now()
    path = Path(root) / HISTORY_FILE
    try:
        history = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        history = {"version": 1, "items": {}}
    items = history.setdefault("items", {})
    changed = False
    for nm_id, item in prices.items():
        events = items.setdefault(str(nm_id), [])
        point = [when.isoformat(timespec="seconds"), float(item["price"]), int(item["discount"])]
        if not events or events[-1][1:] != point[1:]:
            events.append(point)
            changed = True
    if changed:
        path.parent.mkdir(parents=True, exist_ok=True)
        with atomic_output_path(path) as temporary:
            temporary.write_text(json.dumps(history, ensure_ascii=False), encoding="utf-8")
    return history


def load_history(root) -> Dict[str, Any]:
    try:
        return json.loads((Path(root) / HISTORY_FILE).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"version": 1, "items": {}}


def load_order_discounts(token: str, today: Optional[date] = None, days: int = LOOKBACK_DAYS) -> Dict[str, List[list]]:
    """Скидка продавца на момент заказа: {nmId: [[дата, цена до скидки, скидка], ...]} за последние дни."""
    today = today or date.today()
    rows = _get(ORDERS_URL, token, {"dateFrom": (today - timedelta(days=days)).isoformat()}) or []
    result: Dict[str, List[list]] = {}
    for row in rows:
        if row.get("nmId") is None or row.get("discountPercent") is None:
            continue
        result.setdefault(str(row["nmId"]), []).append(
            [str(row.get("date") or "")[:10], float(row.get("totalPrice") or 0), int(row["discountPercent"])])
    return result


# ── статистика карточек по дням ────────────────────────────────────────────

def load_daily_funnel(token: str, nm_ids: List[int], today: Optional[date] = None, pause=time.sleep) -> Dict[str, Dict[str, list]]:
    """Переходы в карточку, корзины и заказы по дням за последнюю неделю: {nmId: {дата: [переходы, корзина, заказы]}}."""
    today = today or date.today()
    period = {"start": (today - timedelta(days=FUNNEL_DAYS - 1)).isoformat(), "end": today.isoformat()}
    result: Dict[str, Dict[str, list]] = {}
    for index in range(0, len(nm_ids), FUNNEL_CHUNK):
        if index:
            pause(FUNNEL_PAUSE)
        response = requests.post(FUNNEL_HISTORY_URL, headers={"Authorization": token}, timeout=TIMEOUT, json={
            "selectedPeriod": period, "nmIds": nm_ids[index:index + FUNNEL_CHUNK], "skipDeletedNm": True,
            "aggregationLevel": "day"})
        if response.status_code != 200:
            raise WbPromotionsError(f"WB вернул HTTP {response.status_code} на статистику карточек по дням")
        for item in response.json() or []:
            days = result.setdefault(str(item["product"]["nmId"]), {})
            for day in item.get("history") or []:
                days[str(day["date"])[:10]] = [int(day.get("openCount") or 0), int(day.get("cartCount") or 0),
                                               int(day.get("orderCount") or 0)]
    return result


def merge_daily(root, fresh: Dict[str, Dict[str, list]]) -> Dict[str, Any]:
    """Дописать неделю в накопленную статистику wb reports/daily_stats.json (свежие дни заменяют прежние)."""
    path = Path(root) / DAILY_FILE
    try:
        stored = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        stored = {"version": 1, "items": {}}
    for nm_id, days in fresh.items():
        stored["items"].setdefault(nm_id, {}).update(days)
    path.parent.mkdir(parents=True, exist_ok=True)
    with atomic_output_path(path) as temporary:
        temporary.write_text(json.dumps(stored, ensure_ascii=False), encoding="utf-8")
    return stored


def load_daily(root) -> Dict[str, Any]:
    try:
        return json.loads((Path(root) / DAILY_FILE).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"version": 1, "items": {}}


def discount_on(day: str, events: List[list], order_days: Dict[str, int]) -> Optional[Dict[str, Any]]:
    """Скидка артикула в этот день: по истории скидок (последняя запись к концу дня), по заказу в этот день,
    иначе — самая ранняя запись истории, перенесённая назад (оценка: до начала истории скидка не менялась)."""
    known = [event for event in events if event[0][:10] <= day]
    if known:
        return {"discount": known[-1][2], "estimated": False}
    if day in order_days:
        return {"discount": order_days[day], "estimated": False}
    if events:
        return {"discount": events[0][2], "estimated": True}
    return None


def daily_effect(nm_ids: List[str], promo_level: Dict[str, int], history: Dict[str, Any],
                 order_discounts: Dict[str, List[list]], daily: Dict[str, Any]) -> Dict[str, Any]:
    """Переходы и заказы в день: в дни со скидкой акции (≥ скидки акции артикула) и без неё — суммарно по nm_ids."""
    sums = {"promo": [0, 0, 0], "plain": [0, 0, 0]}  # дни, переходы, заказы
    estimated = False
    for nm_id in nm_ids:
        events = (history.get("items") or {}).get(nm_id, [])
        orders = {when: discount for when, _price, discount in order_discounts.get(nm_id, [])}
        for day, (views, _cart, ordered) in sorted((daily.get("items") or {}).get(nm_id, {}).items()):
            found = discount_on(day, events, orders)
            if found is None or promo_level.get(nm_id) is None:
                continue
            estimated |= found["estimated"]
            side = sums["promo" if found["discount"] >= promo_level[nm_id] else "plain"]
            side[0] += 1
            side[1] += views
            side[2] += ordered
    def rate(side):
        days, views, ordered = side
        return {"days": days, "views": views / days if days else None, "orders": ordered / days if days else None,
                "orders_total": ordered}
    promo, plain = rate(sums["promo"]), rate(sums["plain"])
    enough = promo["days"] >= MIN_DAYS and plain["days"] >= MIN_DAYS
    uplift = promo["views"] / plain["views"] if enough and plain["views"] else None
    return {"promo": promo, "plain": plain, "enough": enough, "views_uplift": uplift, "estimated": estimated}


def fact_verdict(effect: Dict[str, Any], multiplier: Optional[float], profit_promo: Optional[float]) -> str:
    """Вердикт по факту: рост переходов в акции против нужного роста продаж (конверсия считается одинаковой)."""
    if not effect["enough"]:
        days_promo, days_plain = effect["promo"]["days"], effect["plain"]["days"]
        return f"мало данных ({days_promo} | {days_plain} дн.)"
    if profit_promo is not None and profit_promo <= 0:
        return "выйти: в акции убыток"
    if effect["views_uplift"] is None or multiplier is None:
        return "нет данных"
    return "держать: буст окупается" if effect["views_uplift"] >= multiplier else "выйти: буст не окупается"


def series_of(article: str) -> str:
    return article[:2] + "xx" if len(article) >= 4 and article[:2].isdigit() else article


# ── синхронизация ──────────────────────────────────────────────────────────

def sync(root, token: str, now: Optional[datetime] = None, nm_ids: Optional[List[int]] = None) -> Dict[str, Any]:
    """Календарь акций с подробностями и скидки в заказах за 30 дней → wb reports/promotions.json;
    статистика карточек по дням за неделю (nm_ids) → wb reports/daily_stats.json."""
    from scripts.wb_finance import identity

    if not token:
        raise WbPromotionsError("Добавьте API-токен WB в настройках.")
    now = now or datetime.now()
    data = {"version": 1, "account": identity(token), "fetched_at": now.isoformat(timespec="seconds"),
            "promotions": load_promotions(token, now)}
    try:
        data["order_discounts"] = load_order_discounts(token, now.date())
    except (WbPromotionsError, requests.RequestException) as exc:
        data["order_discounts"] = (load(root) or {}).get("order_discounts", {})  # статистика — запрос в минуту
        data["order_discounts_error"] = str(exc)
    if nm_ids:
        try:
            merge_daily(root, load_daily_funnel(token, nm_ids, now.date()))
        except (WbPromotionsError, requests.RequestException) as exc:
            data["daily_error"] = str(exc)  # статистика по дням копится дальше со следующей синхронизации
    path = Path(root) / PROMOTIONS_FILE
    path.parent.mkdir(parents=True, exist_ok=True)
    with atomic_output_path(path) as temporary:
        temporary.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return data


def load(root) -> Optional[Dict[str, Any]]:
    from scripts.marketplace_settings import read_settings
    from scripts.wb_finance import identity

    try:
        data = json.loads((Path(root) / PROMOTIONS_FILE).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    token = read_settings(root).get("WB_API_TOKEN") or ""
    return data if data.get("version") == 1 and data.get("account") == identity(token) else None


# ── баланс по артикулам ────────────────────────────────────────────────────

def promo_discount(nm_id: str, history: Dict[str, Any], order_discounts: Dict[str, List[list]],
                   today: Optional[date] = None, days: int = LOOKBACK_DAYS) -> Optional[Dict[str, Any]]:
    """Наибольшая скидка артикула за последние дни — оценка скидки акции: {discount, price_before, source, date}."""
    today = today or date.today()
    since = (today - timedelta(days=days)).isoformat()
    candidates = []
    for when, price, discount in (history.get("items") or {}).get(str(nm_id), []):
        if when[:10] >= since:
            candidates.append((discount, when[:10], price, "история скидок"))
    for when, price, discount in order_discounts.get(str(nm_id), []):
        if when >= since and price:
            candidates.append((discount, when, price, "заказы"))
    if not candidates:
        return None
    discount, when, price, source = max(candidates)
    return {"discount": discount, "price_before": price, "source": source, "date": when}


def verdict(profit_promo: Optional[float], multiplier: Optional[float], above_min: bool = False) -> str:
    if profit_promo is None:
        return "нет данных"
    if above_min:
        return "конфликта нет: цена акции не ниже минимума"
    if profit_promo <= 0:
        return "выйти: в акции убыток"
    if multiplier is None or multiplier <= KEEP_UP_TO:
        return "держать в акции"
    if multiplier <= WATCH_UP_TO:
        return "смотреть по продажам"
    return "выйти: буст не окупится"


def advice(verdict_text: str, left_promo: bool) -> Dict[str, Any]:
    """Вердикт словами для статуса товара. Вердикт говорит, стоит ли товару быть в акции; совет — что делать сейчас.
    tone: «act» — нужно действие, «consider» — решить, «ok» — всё правильно, None — нейтрально."""
    exit_verdict = verdict_text.startswith("выйти")
    keep_verdict = verdict_text.startswith(("держать", "смотреть"))
    if left_promo:
        if exit_verdict:  # уже вышел — и правильно
            return {"text": "не возвращать" + verdict_text[len("выйти"):], "tone": "ok"}
        if keep_verdict:
            return {"text": "вернуть в акцию: буст окупается" if verdict_text.startswith("держать")
                    else "можно вернуть: смотреть по продажам", "tone": "consider"}
    elif exit_verdict:
        return {"text": verdict_text + " — поднять цену", "tone": "act"}
    elif verdict_text.startswith("держать"):
        return {"text": verdict_text, "tone": "ok"}
    elif verdict_text.startswith("смотреть"):
        return {"text": verdict_text, "tone": "consider"}
    return {"text": verdict_text, "tone": None}


def article_rows(price_rows: List[Dict[str, Any]], history: Dict[str, Any], order_discounts: Dict[str, List[list]],
                 today: Optional[date] = None, daily: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
    """Строка на артикул: цена и прибыль сейчас, в акции (по наибольшей скидке за 30 дней) и на минимальной цене."""
    rows = []
    for row in price_rows:
        result = {"nm_id": row["nm_id"], "article": row["article"], "price": row["price"], "discount": row["discount"],
                  "discounted_price": row["discounted_price"], "min_price": row["min_price"], "notes": row["notes"],
                  "promo_discount": None, "promo_price": None, "promo_source": None, "status": "",
                  "profit_now": row["profit"], "profit_promo": None, "profit_min": None, "multiplier": None}
        promo = promo_discount(str(row["nm_id"]), history, order_discounts, today)
        if promo and promo["discount"] < row["discount"]:
            promo = None  # сейчас скидка больше всех прежних — значит, акция и есть сейчас
        if promo is None:
            promo = {"discount": row["discount"], "price_before": row["price"], "source": "сейчас", "date": None}
        # Цена акции — от текущей цены до скидки: на неё WB и наложит скидку акции.
        result["promo_discount"] = promo["discount"]
        result["promo_source"] = promo["source"]
        result["promo_date"] = promo["date"]
        result["promo_price"] = round(row["price"] * (100 - promo["discount"]) / 100, 2)
        # Статус по скидкам: скидка ниже прежней наибольшей — товар вывели из акции поднятием цены. Если наибольшая
        # скидка и есть текущая, по истории не отличить «в акции» от «просто такая скидка».
        result["status"] = "вышел из акции" if row["discount"] < promo["discount"] else "по истории не видно"
        if not row["notes"]:
            cost, logistics, commission, acquiring = row["cost"], row["logistics"], row["commission"], row["acquiring"]
            result["profit_promo"] = unit_profit(result["promo_price"], cost, logistics, commission, acquiring)
            result["profit_min"] = unit_profit(row["min_price"], cost, logistics, commission, acquiring) \
                if row["min_price"] else None
            if result["profit_min"] is not None and result["profit_promo"] > 0:
                result["multiplier"] = max(result["profit_min"], result["profit_promo"]) / result["profit_promo"]
        above_min = row["min_price"] is not None and result["promo_price"] >= row["min_price"]
        result["verdict"] = verdict(result["profit_promo"], result["multiplier"], above_min) if not row["notes"] else "нет данных"
        result["effect"] = daily_effect([str(row["nm_id"])], {str(row["nm_id"]): promo["discount"]}, history,
                                        order_discounts, daily or {})
        result["fact"] = fact_verdict(result["effect"], result["multiplier"], result["profit_promo"])
        left = result["status"] == "вышел из акции"
        result["advice"] = advice(result["verdict"], left)
        result["fact_advice"] = advice(result["fact"], left)
        rows.append(result)
    return rows


def series_rows(rows: List[Dict[str, Any]], history: Dict[str, Any], order_discounts: Dict[str, List[list]],
                daily: Dict[str, Any]) -> List[Dict[str, Any]]:
    """То же сравнение по сериям (11xx, 12xx, …): продаж и переходов у серии больше — оценка устойчивее."""
    groups = defaultdict(list)
    for row in rows:
        groups[series_of(row["article"])].append(row)
    result = []
    for name, members in sorted(groups.items()):
        effect = daily_effect([str(r["nm_id"]) for r in members], {str(r["nm_id"]): r["promo_discount"] for r in members},
                              history, order_discounts, daily)
        result.append({"series": name, "articles": len(members), "effect": effect})
    return result
