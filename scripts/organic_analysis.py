# -*- coding: utf-8 -*-
"""
Анализ органических продаж: почему товар продаётся (или нет) без рекламы.

Органика = все заказы товара − заказы, которые Ozon приписал рекламе.
Период сравнивается с таким же предыдущим отрезком, а сезонность — с теми же неделями год назад
(сдвиг 364 дня, чтобы совпадали дни недели).

По каждому товару — шаги диагностики с цифрами: продажи, наличие и доставка, цена, карточка,
воронка (по рекламному трафику), зависимость от рекламы, поиск (если у продавца есть Premium).
Ozon не отдаёт историю цен, остатков и рейтинга, поэтому программа сохраняет свои снимки
(snapshots) и по ним показывает, что менялось перед падением.
"""

import json
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import pandas as pd
import requests

script_dir = Path(__file__).resolve().parent
repo_root = script_dir.parent
if str(script_dir) not in sys.path:
    sys.path.insert(0, str(script_dir))

SELLER_HOST = "https://api-seller.ozon.ru"
HISTORY_DIR = "analytics history"
SNAPSHOT_DIR = "snapshots"
LAST_RESULT = "organic_last.json"
PERIOD_OPTIONS = (7, 14, 28, 56)
DEFAULT_PERIOD = 28
YEAR_SHIFT = 364          # 52 недели — те же дни недели год назад
WEEKS_IN_CHART = 12
ADS_MAX_DAYS = 62

GOOD_INDEX = {"SUPER", "GOOD", "GREEN"}
MID_INDEX = {"AVG", "AVERAGE", "YELLOW", "MODERATE"}
BAD_INDEX = {"BAD", "RED", "POOR", "NON_PROFIT"}
INDEX_NAMES = {"SUPER": "очень выгодная", "GOOD": "выгодная", "GREEN": "выгодная", "AVG": "умеренная", "AVERAGE": "умеренная",
               "YELLOW": "умеренная", "MODERATE": "умеренная", "BAD": "невыгодная", "RED": "невыгодная", "POOR": "невыгодная",
               "NON_PROFIT": "невыгодная", "WITHOUT_INDEX": "нет данных"}

STATUS_ORDER = {"bad": 0, "warn": 1, "info": 2, "ok": 3}


# ---------------------------------------------------------------- периоды

def analysis_periods(today: date, days: int) -> Dict[str, Tuple[date, date]]:
    """Полуоткрытые отрезки [начало, конец): текущий — по вчерашний день включительно."""
    end = today
    current = (end - timedelta(days=days), end)
    previous = (current[0] - timedelta(days=days), current[0])
    shift = timedelta(days=YEAR_SHIFT)
    return {
        "current": current,
        "previous": previous,
        "current_ly": (current[0] - shift, current[1] - shift),
        "previous_ly": (previous[0] - shift, previous[1] - shift),
    }


def _units(orders: pd.DataFrame, span: Tuple[date, date], offer: Optional[str] = None) -> float:
    part = orders[(orders["date"] >= pd.Timestamp(span[0])) & (orders["date"] < pd.Timestamp(span[1]))]
    if offer is not None:
        part = part[part["offer_id"] == offer]
    return float(part["qty"].sum())


def _change(current: float, previous: float) -> Optional[float]:
    return (current - previous) / previous if previous else None


def _pct(value: Optional[float], digits: int = 0, signed: bool = False) -> str:
    if value is None:
        return "—"
    if round(value * 100, digits) == 0:
        return "0%"
    text = f"{value * 100:+.{digits}f}%" if signed else f"{value * 100:.{digits}f}%"
    return text.replace(".", ",")


def _rub(value: Optional[float]) -> str:
    return "—" if value is None else f"{value:,.0f} ₽".replace(",", " ")


def _step(name: str, status: str, text: str) -> Dict[str, str]:
    return {"name": name, "status": status, "text": text}


# ---------------------------------------------------------------- анализ товара (чистая логика)

def season_expectation(orders: pd.DataFrame, periods: Dict[str, Tuple[date, date]]) -> Optional[float]:
    """Как менялись продажи магазина между теми же отрезками год назад (сезонная норма изменения)."""
    before, after = _units(orders, periods["previous_ly"]), _units(orders, periods["current_ly"])
    return _change(after, before) if before >= 10 else None


def analyze_product(p: Dict[str, Any], store: Dict[str, Any]) -> Dict[str, Any]:
    """
    p — данные товара за текущий и предыдущий период; store — средние по магазину.
    Возвращает шаги диагностики, главную проблему и рекомендации.
    """
    steps: List[Dict[str, str]] = []
    advice: List[Tuple[float, str, str]] = []   # (вес, шаг, текст): чем больше вес, тем выше рекомендация
    total, total_prev = p["orders"], p["orders_prev"]
    organic, organic_prev = p["organic"], p["organic_prev"]
    org_change = _change(organic, organic_prev)
    season = store.get("season_change")
    lost = max(0.0, organic_prev - organic)

    # 1. Продажи
    text = f"Заказов {total:.0f} (было {total_prev:.0f}), из них органических {organic:.0f} (было {organic_prev:.0f}"
    text += f", {_pct(org_change, signed=True)})." if org_change is not None else ")."
    if season is not None:
        text += f" Год назад в эти же недели продажи магазина изменились на {_pct(season, signed=True)} — это сезонная норма."
    status = "ok"
    if total == 0 and total_prev == 0:
        status = "bad"
        text += " Товар не продаётся ни с рекламой, ни без неё."
    elif org_change is not None and organic_prev >= 3:
        gap = org_change - (season or 0.0)
        if gap <= -0.3:
            status = "bad" if gap <= -0.5 else "warn"
            if season is not None:
                text += f" По сезону ждали {_pct(season, signed=True)}, а органика {_pct(org_change, signed=True)} — причина в самом товаре, а не в сезоне."
            else:
                text += f" Органика {_pct(org_change, signed=True)} — заметное падение."
        elif season is not None and org_change < 0 and gap > -0.3:
            text += " Падение в пределах сезона."
    elif organic == 0 and total > 0:
        status = "warn"
        text += " Без рекламы товар не продаётся."
    steps.append(_step("Продажи", status, text))

    # 2. Наличие и доставка
    stock = p.get("stock")
    locality = p.get("locality")
    parts = []
    status = "ok"
    if stock == 0:
        status = "bad"
        parts.append("Свободного остатка нет ни на складах Ozon (FBO), ни на вашем складе (FBS) — купить товар сейчас нельзя, и в поиске его не показывают.")
        advice.append((1000, "Наличие и доставка", "Пополните остаток: без товара органики не будет."))
    elif stock is not None:
        parts.append(f"В наличии {stock} шт.")
    if p.get("days_since_sale") is not None:
        parts.append(f"Последний заказ {p['days_since_sale']} дн. назад.")
    if locality and locality.get("total"):
        idx = locality["local"] / locality["total"]
        parts.append(f"Из кластера покупателя доставлено {locality['local']} из {locality['total']} ({_pct(idx)}).")
        if idx < 0.3 and locality["total"] >= 3:
            status = "warn" if status == "ok" else status
            parts.append("Остальные везли из других регионов: доставка дольше, а срок доставки влияет на выдачу и на решение купить.")
            if locality.get("recommended_supply"):
                advice.append((300 + lost * 10, "Наличие и доставка", f"Поставьте товар на FBO ближе к покупателям (Ozon советует {locality['recommended_supply']} шт.) — "
                                                "быстрая доставка поднимает товар в выдаче. Расчёт количества — в разделе «Поставки»."))
    steps.append(_step("Наличие и доставка", status, " ".join(parts) or "Данных об остатке нет."))

    # 3. Цена
    parts, status = [], "ok"
    price_now, price_before = p.get("avg_price"), p.get("avg_price_prev")
    if price_now and price_before:
        delta = _change(price_now, price_before)
        parts.append(f"Средняя цена продажи {_rub(price_now)} против {_rub(price_before)} в прошлом периоде ({_pct(delta, signed=True)}).")
        if delta is not None and delta >= 0.1 and (org_change or 0) < 0:
            status = "warn"
            advice.append((200 + lost * 10, "Цена", f"Цена выросла на {_pct(delta)}, и органика упала — проверьте, не вышли ли вы из акций "
                                            "и не стал ли товар дороже конкурентов."))
    index = p.get("price_index") or {}
    color = index.get("color") or "WITHOUT_INDEX"
    driver, driver_price, value = index.get("driver"), index.get("driver_price"), index.get("value")
    who = {
        "self": f"ваш же товар на другой площадке за {_rub(driver_price)}",
        "ozon": f"другой продавец на Ozon за {_rub(driver_price)}",
        "external": f"конкурент на другой площадке за {_rub(driver_price)}",
    }.get(driver or "", "")
    value_text = f" ({value:.2f})".replace(".", ",") if value else ""
    if color in BAD_INDEX or color in MID_INDEX:
        status = "bad" if color in BAD_INDEX else ("warn" if status == "ok" else status)
        parts.append(f"Индекс цен — «{INDEX_NAMES.get(color, color)}»{value_text}" + (f": дешевле всех — {who}." if who else "."))
        target = price_for_good_index(p.get("price_now"), value)
        floor = p.get("min_price")
        if target:
            parts.append(f"Для индекса около 1,00 цену на Ozon нужно снизить примерно до {_rub(target)} (сейчас {_rub(p.get('price_now'))}).")
        below_floor = bool(target and floor and target < floor)
        if below_floor:
            parts.append(f"Это ниже вашей минимальной цены {_rub(floor)} — снижать на Ozon невыгодно.")
        weight = (400 if color in BAD_INDEX else 200) + lost * 10
        if driver == "self":
            fix = (f"поднимите цену на другой площадке выше {_rub(driver_price)}" if below_floor
                   else f"поднимите цену на другой площадке или снизьте здесь примерно до {_rub(target)}" if target
                   else "выровняйте цены с другой площадкой")
            advice.append((weight, "Цена", f"Ozon понижает товар, потому что дешевле всего ваш же товар на другой площадке ({_rub(driver_price)}): {fix}. "
                                           "Вы конкурируете сами с собой."))
        elif below_floor:
            advice.append((weight, "Цена", f"Цена невыгодная: {who}. Выгодный индекс требует ≈ {_rub(target)}, а это ниже вашей минимальной цены {_rub(floor)} — "
                                           "конкурировать ценой не выйдет; делайте ставку на карточку, отзывы и доставку."))
        else:
            advice.append((weight, "Цена", f"Индекс цен «{INDEX_NAMES.get(color, color)}» ({who or 'есть дешевле'}) — Ozon понижает товар в выдаче"
                                           + (f": снизьте цену примерно до {_rub(target)} или участвуйте в акциях." if target else ".")))
    elif color in GOOD_INDEX:
        parts.append(f"Индекс цен — «{INDEX_NAMES.get(color, color)}»{value_text}: цена конкурентная.")
    else:
        parts.append("Индекса цен нет — Ozon не нашёл таких же товаров ни у кого на Ozon и на других площадках.")
    actions = p.get("actions") or []
    parts.append(f"Участвует в акциях: {len(actions)}." if actions else "В акциях не участвует — нет акционного бустинга в выдаче.")
    if not actions and stock:
        advice.append((60, "Цена", "Товар не в акциях: акционный бустинг поднимает в поиске. Посмотрите выгодные акции в разделе «Акции»."))
    steps.append(_step("Цена", status, " ".join(parts)))

    # 4. Карточка
    rating = p.get("content_rating")
    if rating is None:
        steps.append(_step("Карточка", "info", "Контент-рейтинг не получен."))
    else:
        missing = p.get("content_missing") or []
        status = "ok" if rating >= 90 else ("warn" if rating >= 70 else "bad")
        text = f"Контент-рейтинг {rating:.0f} из 100."
        if missing:
            text += " Что добавить: " + "; ".join(missing[:4]) + "."
        steps.append(_step("Карточка", status, text))
        if status != "ok" and missing:
            advice.append(((150 if status == "warn" else 250) + lost * 5, "Карточка", f"Доработайте карточку (рейтинг {rating:.0f}): " + "; ".join(missing[:3]) + ". "
                                                                         "Полная карточка лучше ранжируется и чаще покупается."))

    # 5. Воронка по рекламному трафику
    ad = p.get("ads") or {}
    if ad.get("views", 0) >= 100:
        ctr = _ratio(ad["clicks"], ad["views"])
        cart = _ratio(ad["to_cart"], ad["clicks"])
        order = _ratio(ad["orders"], ad["to_cart"])
        weak = []
        if ctr is not None and store.get("ctr") and ctr < store["ctr"] * 0.6:
            weak.append("редко кликают — главное фото, цена в выдаче, рейтинг")
        if cart is not None and ad["clicks"] >= 30 and store.get("cart_rate") and cart < store["cart_rate"] * 0.6:
            weak.append("не кладут в корзину — описание, фото, отзывы")
        if order is not None and ad["to_cart"] >= 8 and store.get("cart_to_order") and order < store["cart_to_order"] * 0.6:
            weak.append("не выкупают из корзины — цена и срок доставки")
        text = (f"По рекламному трафику: CTR {_pct(ctr, 1)} (в среднем {_pct(store.get('ctr'), 1)}), в корзину {_pct(cart, 1)} "
                f"({_pct(store.get('cart_rate'), 1)}), из корзины в заказ {_pct(order)} ({_pct(store.get('cart_to_order'))}).")
        if weak:
            text += " Слабо: " + "; ".join(weak) + ". Те же покупатели приходят и из поиска — это проблема и для органики."
            advice.append((180 + lost * 5, "Воронка", "Воронка карточки слабее средней: " + "; ".join(weak) + "."))
        steps.append(_step("Воронка", "warn" if weak else "ok", text))
    else:
        steps.append(_step("Воронка", "info", "Мало рекламного трафика, чтобы оценить карточку по воронке (нужно от 100 показов)."))

    # 6. Зависимость от рекламы
    ad_orders = p.get("ad_orders", 0.0)
    if total > 0:
        share = ad_orders / total
        text = f"С рекламы {ad_orders:.0f} из {total:.0f} заказов ({_pct(share)})."
        status = "ok"
        if share >= 0.7 and total >= 3:
            status = "warn"
            text += " Товар держится на рекламе: в органической выдаче он, скорее всего, стоит низко."
            advice.append((220 + total * 5, "Реклама", "Почти все продажи — с рекламы. Проверьте название и атрибуты на ключевые слова, "
                                            "категорию и цену: органическая позиция слабая."))
        steps.append(_step("Реклама", status, text))
    else:
        steps.append(_step("Реклама", "info", "Заказов в периоде нет."))

    # 7. Поиск (Premium)
    search = p.get("search")
    if search:
        parts = []
        if search.get("searched") is not None:
            parts.append(f"Искали по запросам с товаром: {search['searched']:.0f} покупателей")
        if search.get("viewed") is not None:
            parts.append(f"увидели товар: {search['viewed']:.0f}")
        if search.get("position") is not None:
            parts.append(f"средняя позиция {search['position']:.0f}")
        status = "info"
        if search.get("position") is not None and search["position"] > 50:
            status = "warn"
            advice.append((260 + lost * 10, "Поиск", f"Средняя позиция в поиске {search['position']:.0f} — товар далеко. Добавьте в название и атрибуты слова из запросов, "
                                            "по которым его ищут, и проверьте категорию."))
        queries = search.get("queries") or []
        if queries:
            top = ", ".join(f"«{q['query']}» (поз. {q['position']:.0f})" if q.get("position") is not None else f"«{q['query']}»" for q in queries[:5])
            parts.append(f"главные запросы: {top}")
        if search.get("position") is None and search.get("viewed") is None:
            parts.append("позиции и просмотры в поиске Ozon отдаёт только с подпиской Premium")
        text = "; ".join(parts) + "." if parts else "Данных по запросам нет."
        steps.append(_step("Поиск", status, text[0].upper() + text[1:]))
    else:
        steps.append(_step("Поиск", "info", "Позиции и запросы Ozon отдаёт только с подпиской Premium — без неё видимость в поиске "
                                            "оцениваем косвенно: по воронке, зависимости от рекламы и цене."))

    if total == 0 and total_prev == 0 and stock:
        advice.append((500, "Продажи", "Товар есть, но не продаётся: проверьте, находится ли он в поиске по своим главным словам "
                            "(название, категория, атрибуты) и не завышена ли цена."))

    advice.sort(key=lambda item: -item[0])
    # Главная проблема — шаг, к которому относится самая важная рекомендация; без рекомендаций — самый тяжёлый шаг.
    statuses = {s["name"]: s["status"] for s in steps}
    if advice:
        main = advice[0][1]
        main_status = statuses.get(main, "warn") if statuses.get(main) in ("bad", "warn") else "warn"
    else:
        worst = min(steps, key=lambda s: STATUS_ORDER[s["status"]])
        main = worst["name"] if worst["status"] in ("bad", "warn") else None
        main_status = worst["status"] if main else "ok"
    return {
        "steps": steps,
        "main_problem": main,
        "main_status": main_status,
        "advice": [text for _, _, text in advice],
        "priority": (advice[0][0] if advice else 0) + lost * 10,
        "organic_change": org_change,
        "lost_orders": lost,
    }


def _ratio(numerator: float, denominator: float) -> Optional[float]:
    return numerator / denominator if denominator else None


def store_funnel_benchmarks(ads_by_sku: Dict[str, Dict[str, float]]) -> Dict[str, Optional[float]]:
    views = sum(a["views"] for a in ads_by_sku.values())
    clicks = sum(a["clicks"] for a in ads_by_sku.values())
    carts = sum(a["to_cart"] for a in ads_by_sku.values())
    orders = sum(a["orders"] for a in ads_by_sku.values())
    return {"ctr": _ratio(clicks, views), "cart_rate": _ratio(carts, clicks), "cart_to_order": _ratio(orders, carts),
            "views": views, "clicks": clicks, "to_cart": carts, "orders": orders}


# ---------------------------------------------------------------- снимки (своя история)

def snapshot_dir(root: Path) -> Path:
    path = root / HISTORY_DIR / SNAPSHOT_DIR
    path.mkdir(parents=True, exist_ok=True)
    return path


def save_snapshot(root: Path, day: date, products: Dict[str, Dict[str, Any]]) -> None:
    """Один снимок в день: цена, остаток, индекс цен, контент-рейтинг, число акций по каждому артикулу."""
    (snapshot_dir(root) / f"{day.isoformat()}.json").write_text(json.dumps(products, ensure_ascii=False), encoding="utf-8")


def load_snapshots(root: Path, since: date) -> List[Tuple[date, Dict[str, Dict[str, Any]]]]:
    result = []
    for path in sorted(snapshot_dir(root).glob("*.json")):
        try:
            day = date.fromisoformat(path.stem)
        except ValueError:
            continue
        if day < since:
            continue
        try:
            result.append((day, json.loads(path.read_text(encoding="utf-8"))))
        except (OSError, ValueError):
            continue
    return result


def snapshot_events(snapshots: List[Tuple[date, Dict[str, Dict[str, Any]]]], offer: str) -> List[str]:
    """Что менялось у товара между снимками: цена, остаток, индекс цен, рейтинг, акции."""
    events, previous = [], None
    for day, products in snapshots:
        current = products.get(offer)
        if current is None:
            continue
        if previous is not None:
            when = day.strftime("%d.%m")
            old_price, new_price = previous.get("price"), current.get("price")
            if old_price and new_price and abs(new_price - old_price) / old_price >= 0.05:
                events.append(f"{when}: цена {_rub(old_price)} → {_rub(new_price)}")
            old_stock, new_stock = previous.get("stock"), current.get("stock")
            if old_stock and new_stock == 0:
                events.append(f"{when}: закончился остаток")
            elif old_stock == 0 and new_stock:
                events.append(f"{when}: товар снова в наличии ({new_stock} шт.)")
            if previous.get("index") != current.get("index") and current.get("index"):
                events.append(f"{when}: индекс цен «{INDEX_NAMES.get(previous.get('index'), previous.get('index'))}» → «{INDEX_NAMES.get(current['index'], current['index'])}»")
            if previous.get("actions") != current.get("actions") and current.get("actions") is not None:
                events.append(f"{when}: акций {previous.get('actions')} → {current.get('actions')}")
            old_rating, new_rating = previous.get("rating"), current.get("rating")
            if old_rating is not None and new_rating is not None and abs(new_rating - old_rating) >= 5:
                events.append(f"{when}: контент-рейтинг {old_rating:.0f} → {new_rating:.0f}")
        previous = current
    return events


# ---------------------------------------------------------------- данные Ozon

def _post(path: str, body: Dict[str, Any]) -> requests.Response:
    from recommended_prices import OZON_HEADERS
    return requests.post(SELLER_HOST + path, headers=OZON_HEADERS, json=body, timeout=60)


def parse_price_index(indexes: Dict[str, Any]) -> Dict[str, Any]:
    """
    Индекс цен Ozon: три сравнения — с продавцами на Ozon, с конкурентами вне Ozon и с вашим же товаром
    на других площадках. Худшее (наибольшее значение индекса) и определяет, кто «дешевле».
    Если минимальная цена конкурентов вне Ozon совпадает с вашей на других площадках — дешевле вы сами.
    """
    def read(key: str) -> Tuple[Optional[float], Optional[float]]:
        data = indexes.get(key) or {}
        price = float(data.get("min_price") or 0) or None
        value = float(data.get("price_index_value") or 0) or None
        return price, value

    ozon_price, ozon_value = read("ozon_index_data")
    external_price, external_value = read("external_index_data")
    self_price, self_value = read("self_marketplaces_index_data")
    candidates = [(value, source, price) for source, price, value in
                  (("ozon", ozon_price, ozon_value), ("external", external_price, external_value), ("self", self_price, self_value)) if value]
    driver, driver_price, value = None, None, None
    if candidates:
        value, driver, driver_price = max(candidates)
        if driver == "external" and self_price and external_price and abs(self_price - external_price) < 1:
            driver = "self"
    return {"color": indexes.get("color_index") or "WITHOUT_INDEX", "value": value, "driver": driver, "driver_price": driver_price,
            "ozon_min": ozon_price, "external_min": external_price, "self_min": self_price,
            "min_price": driver_price}


def price_for_good_index(price_now: Optional[float], index_value: Optional[float]) -> Optional[float]:
    """
    Оценка цены, при которой индекс станет около 1,00 (граница «выгодного»): индекс растёт вместе с вашей
    ценой, поэтому делим текущую цену на текущий индекс. Ozon считает индекс по цене для покупателя,
    поэтому это ориентир, а не точное значение.
    """
    if not price_now or not index_value or index_value <= 1:
        return None
    return round(price_now / index_value)


def load_prices(offers: List[str]) -> Dict[str, Dict[str, Any]]:
    """Текущая цена, индекс цен и акции по артикулам (/v5/product/info/prices)."""
    from recommended_prices import _normalize_offer_id
    result: Dict[str, Dict[str, Any]] = {}
    for i in range(0, len(offers), 100):  # /v5/product/info/prices: не больше 100 товаров за запрос
        cursor = ""
        while True:
            response = _post("/v5/product/info/prices", {"cursor": cursor, "filter": {"offer_id": offers[i:i + 100], "visibility": "ALL"}, "limit": 100})
            response.raise_for_status()
            data = response.json()
            items = data.get("items") or []
            for item in items:
                price = item.get("price") or {}
                actions = (item.get("marketing_actions") or {}).get("actions") or []
                result[_normalize_offer_id(str(item.get("offer_id") or ""))] = {
                    "price": float(price.get("marketing_seller_price") or price.get("price") or 0) or None,
                    "index": parse_price_index(item.get("price_indexes") or {}),
                    "actions": [a.get("title") for a in actions if a.get("title")],
                }
            cursor = data.get("cursor") or ""
            if not items or not cursor or len(items) < 1000:
                break
    return result


def load_content_rating(skus: List[int]) -> Dict[int, Dict[str, Any]]:
    """Контент-рейтинг и невыполненные условия, которые дают больше всего баллов."""
    result: Dict[int, Dict[str, Any]] = {}
    for i in range(0, len(skus), 100):
        response = _post("/v1/product/rating-by-sku", {"skus": [str(s) for s in skus[i:i + 100]]})
        response.raise_for_status()
        for product in response.json().get("products") or []:
            missing = []
            for group in product.get("groups") or []:
                open_conditions = sorted((c for c in group.get("conditions") or [] if not c.get("fulfilled")), key=lambda c: -(c.get("cost") or 0))
                if open_conditions:
                    task = open_conditions[0].get("description", "").lower()
                    for done, todo in (("добавлено ", "добавить "), ("добавлена ", "добавить "), ("добавлены ", "добавить "), ("заполнено ", "заполнить ")):
                        task = task.replace(done, todo)
                    missing.append(f"{group.get('name')} — {task}")
                for attribute in group.get("improve_attributes") or []:
                    name = attribute.get("name")
                    if name and not any(name in m for m in missing):
                        missing.append(f"заполнить «{name}»")
            result[int(product["sku"])] = {"rating": float(product.get("rating") or 0), "missing": missing}
    return result


def load_locality(skus: List[int], span: Tuple[date, date]) -> Dict[int, Dict[str, Any]]:
    """Сколько единиц доставлено из кластера покупателя и сколько Ozon советует поставить."""
    result: Dict[int, Dict[str, Any]] = {}
    offset = 0
    while True:
        response = _post("/v1/analytics/local-sale/items-clusters/info", {
            "filter": {"skus": [str(s) for s in skus], "delivery_schema": "ALL",
                       "period": {"from": span[0].isoformat(), "to": (span[1] - timedelta(days=1)).isoformat()}},
            "limit": 1000, "offset": offset,
        })
        response.raise_for_status()
        data = response.json()
        items = data.get("items") or []
        for item in items:
            metrics = item.get("metrics") or {}
            local = metrics.get("local_data") or {}
            entry = result.setdefault(int(item["sku"]), {"local": 0, "total": 0, "recommended_supply": 0, "overpayment": 0.0})
            entry["local"] += int(local.get("local_quantity") or 0)
            entry["total"] += int(local.get("total_quantity") or 0)
            entry["recommended_supply"] += int(metrics.get("recommended_supply") or 0)
            entry["overpayment"] += float((metrics.get("overpayment") or {}).get("total") or 0)
        offset += len(items)
        if not items or offset >= int(data.get("total") or 0):
            break
    return result


def load_search(skus: List[int], span: Tuple[date, date]) -> Tuple[bool, Dict[int, Dict[str, Any]], str]:
    """
    Поиск по запросам (/v1/analytics/product-queries и /details). Позиции и «увидели»
    Ozon отдаёт только с Premium: без подписки — пусто или 403. Возвращает (premium, данные, пояснение).
    """
    period_end = datetime.combine(span[1] - timedelta(days=1), datetime.min.time())
    body = {"date_from": datetime.combine(span[0], datetime.min.time()).isoformat() + "Z", "date_to": period_end.isoformat() + "Z"}
    result: Dict[int, Dict[str, Any]] = {}
    for i in range(0, len(skus), 1000):
        page = 0
        while True:
            response = _post("/v1/analytics/product-queries", {**body, "skus": [str(s) for s in skus[i:i + 1000]], "page": page, "page_size": 1000})
            if response.status_code == 403:
                return False, {}, "Поисковая аналитика за этот период доступна только с подпиской Premium."
            response.raise_for_status()
            data = response.json()
            for item in data.get("items") or []:
                result[int(item["sku"])] = {"searched": item.get("unique_search_users"), "viewed": item.get("unique_view_users"),
                                            "position": item.get("position"), "conversion": item.get("view_conversion"), "queries": []}
            page += 1
            if page >= int(data.get("page_count") or 0):
                break
    premium = any(v.get("position") is not None or v.get("viewed") is not None for v in result.values())
    if not premium:
        return False, result, "Без подписки Premium Ozon не показывает позиции и просмотры в поиске."
    page = 0
    while True:
        response = _post("/v1/analytics/product-queries/details", {**body, "skus": [str(s) for s in result], "page": page, "page_size": 100, "limit_by_sku": 15})
        if response.status_code == 403:
            break
        response.raise_for_status()
        data = response.json()
        for query in data.get("queries") or []:
            entry = result.get(int(query["sku"]))
            if entry is not None:
                entry["queries"].append({"query": query.get("query"), "position": query.get("position"), "orders": query.get("order_count"),
                                         "searched": query.get("unique_search_users"), "viewed": query.get("unique_view_users")})
        page += 1
        if page >= int(data.get("page_count") or 0):
            break
    for entry in result.values():
        entry["queries"].sort(key=lambda q: -(q.get("searched") or 0))
    return True, result, "Подписка Premium подключена — показываем позиции и запросы."


def load_ads(spans: Dict[str, Tuple[date, date]], today: date) -> Tuple[Dict[str, Dict[str, Dict[str, float]]], List[Tuple[date, float]], str]:
    """Рекламная статистика по SKU за оба периода и рекламные заказы по неделям для графика."""
    import ads_cpc
    client = ads_cpc.PerfClient()
    campaigns = ads_cpc.list_cpc_campaigns(spans["previous"][0], today - timedelta(days=1))["campaigns"]
    ids = [c["id"] for c in campaigns]
    result: Dict[str, Dict[str, Dict[str, float]]] = {"current": {}, "previous": {}}
    for key in ("current", "previous"):
        start, end = spans[key]
        if not ids:
            break
        for campaign_id, rows in client.sku_statistics(ids, start, end - timedelta(days=1)).items():
            for row in rows:
                sku = str(row.get("sku"))
                entry = result[key].setdefault(sku, {"views": 0.0, "clicks": 0.0, "to_cart": 0.0, "orders": 0.0, "models": 0.0, "spend": 0.0, "gmv": 0.0})
                entry["views"] += ads_cpc._num(row.get("views"))
                entry["clicks"] += ads_cpc._num(row.get("clicks"))
                entry["to_cart"] += ads_cpc._num(row.get("toCart"))
                entry["orders"] += ads_cpc._num(row.get("orders"))
                entry["models"] += ads_cpc._num(row.get("models"))
                entry["spend"] += ads_cpc._num(row.get("moneySpent"))
                entry["gmv"] = max(entry["gmv"], ads_cpc._num(row.get("product_gmv")))
    weekly: Dict[date, float] = {}
    chart_start = _week_start(today) - timedelta(weeks=WEEKS_IN_CHART - 1)
    all_ids = [c["id"] for c in ads_cpc.list_cpc_campaigns(chart_start, today - timedelta(days=1))["campaigns"]]
    start = chart_start
    while all_ids and start < today:
        end = min(today - timedelta(days=1), start + timedelta(days=ADS_MAX_DAYS - 1))
        data = client.call("GET", "/api/client/statistics/daily/json", params={"campaignIds": all_ids, "dateFrom": start.isoformat(), "dateTo": end.isoformat()})
        for row in data.get("rows") or []:
            day = date.fromisoformat(str(row.get("date")))
            weekly[_week_start(day)] = weekly.get(_week_start(day), 0.0) + ads_cpc._num(row.get("orders"))
        start = end + timedelta(days=1)
    return result, sorted(weekly.items()), f"Рекламных кампаний за период: {len(ids)}."


def _week_start(day: date) -> date:
    return day - timedelta(days=day.weekday())


def load_price_history(root: Path, spans: Dict[str, Tuple[date, date]]) -> Dict[str, Dict[str, float]]:
    """Средняя фактическая цена продажи по артикулу в текущем и предыдущем периоде — из месячных отчётов."""
    from recommended_prices import MONTHS_RU, _artikul_normalize, _normalize_offer_id
    start, end = spans["previous"][0], spans["current"][1]
    frames = []
    month = date(start.year, start.month, 1)
    while month < end:
        path = root / "reports" / f"{MONTHS_RU[month.month - 1]} {month.year}.xlsx"
        if path.exists():
            try:
                frames.append(pd.read_excel(path, sheet_name="Заказы", usecols=["Статус", "Артикул", "Количество шт.", "Цена продажи", "Дата отгрузки"]))
            except (ValueError, OSError):
                pass
        month = date(month.year + (month.month // 12), month.month % 12 + 1, 1)
    if not frames:
        return {}
    df = pd.concat(frames, ignore_index=True)
    df = df[df["Статус"].astype(str).str.lower() != "cancelled"]
    df["date"] = pd.to_datetime(df["Дата отгрузки"], errors="coerce")
    df["price"] = pd.to_numeric(df["Цена продажи"], errors="coerce")
    df["qty"] = pd.to_numeric(df["Количество шт."], errors="coerce").fillna(1)
    df["offer"] = [_normalize_offer_id(_artikul_normalize(v) or "") for v in df["Артикул"]]
    df = df.dropna(subset=["date", "price"])
    result: Dict[str, Dict[str, float]] = {}
    for key in ("current", "previous"):
        a, b = spans[key]
        part = df[(df["date"] >= pd.Timestamp(a)) & (df["date"] < pd.Timestamp(b))]
        for offer, group in part.groupby("offer"):
            qty = group["qty"].sum()
            if qty:
                result.setdefault(offer, {})[key] = float((group["price"] * group["qty"]).sum() / qty)
    return result


# ---------------------------------------------------------------- сборка

def run_analysis(root: Path, days: int = DEFAULT_PERIOD, today: Optional[date] = None,
                 progress: Optional[Callable[[str], None]] = None) -> Dict[str, Any]:
    import fbo_supply_plan
    import ads_cpc
    from recommended_prices import load_costs_df, COSTS_FILENAME, _artikul_normalize, _normalize_offer_id, get_product_id_mapping, get_product_skus

    say = progress or (lambda _msg: None)
    today = today or date.today()
    spans = analysis_periods(today, days)
    sources: List[Dict[str, str]] = []

    say("Заказы: загружаю историю (закрытые месяцы — из кэша)…")
    orders = fbo_supply_plan.load_orders(root, today)
    df, key_col, _ = load_costs_df(root / COSTS_FILENAME)
    from recommended_prices import COL_MIN_PRICE
    offers = []
    min_prices: Dict[str, float] = {}
    for raw, floor in zip(df[key_col], df[COL_MIN_PRICE] if COL_MIN_PRICE in df.columns else [None] * len(df)):
        art = _artikul_normalize(raw)
        if art:
            offers.append(_normalize_offer_id(art))
            try:
                if floor is not None and float(floor) > 0:
                    min_prices[_normalize_offer_id(art)] = float(floor)
            except (TypeError, ValueError):
                pass
    offers = sorted(set(offers) | set(orders["offer_id"]))
    pid = get_product_id_mapping(offers)
    sku_by_offer = {pid[p]: s for p, s in get_product_skus(sorted(pid)).items() if p in pid}
    skus = sorted(sku_by_offer.values())

    def safe(name: str, loader: Callable[[], Any], fallback: Any) -> Any:
        try:
            value = loader()
            sources.append({"name": name, "status": "ok", "text": "получено"})
            return value
        except Exception as exc:  # источник недоступен — анализируем без него
            sources.append({"name": name, "status": "bad", "text": str(exc)[:200]})
            return fallback

    say("Реклама: статистика по товарам за оба периода…")
    ads, ads_weekly, ads_note = safe("Реклама", lambda: load_ads(spans, today), ({"current": {}, "previous": {}}, [], ""))
    say("Остатки, цены, индексы цен…")
    stock = safe("Остатки", lambda: ads_cpc.load_available_stock(offers), {})
    prices = safe("Цены и индексы", lambda: load_prices(offers), {})
    price_history = safe("История цен (отчёты)", lambda: load_price_history(root, spans), {})
    say("Карточки и доставка…")
    rating = safe("Контент-рейтинг", lambda: load_content_rating(skus), {})
    locality = safe("Локальность доставки", lambda: load_locality(skus, spans["current"]), {})
    say("Поиск…")
    premium, search, search_note = safe("Поиск (Premium)", lambda: load_search(skus, spans["current"]), (False, {}, "не получено"))
    sources.append({"name": "Подписка Premium", "status": "ok" if premium else "info", "text": search_note})

    snapshot = {offer: {"price": prices.get(offer, {}).get("price"), "stock": stock.get(offer),
                        "index": (prices.get(offer, {}).get("index") or {}).get("color"),
                        "actions": len(prices.get(offer, {}).get("actions") or []) if offer in prices else None,
                        "rating": (rating.get(sku_by_offer.get(offer)) or {}).get("rating")} for offer in offers}
    save_snapshot(root, today, snapshot)
    snapshots = load_snapshots(root, spans["previous"][0])

    bench = store_funnel_benchmarks(ads["current"])
    store_ctx = {**bench, "season_change": season_expectation(orders, spans)}
    last_sale = orders.groupby("offer_id")["date"].max()

    products = []
    for offer in offers:
        sku = sku_by_offer.get(offer)
        ad_cur = ads["current"].get(str(sku), {}) if sku else {}
        ad_prev = ads["previous"].get(str(sku), {}) if sku else {}
        total, total_prev = _units(orders, spans["current"], offer), _units(orders, spans["previous"], offer)
        ad_orders, ad_orders_prev = ad_cur.get("orders", 0.0), ad_prev.get("orders", 0.0)
        if total == 0 and total_prev == 0 and not ad_cur and not stock.get(offer):
            continue
        last = last_sale.get(offer)
        p = {
            "offer_id": offer, "sku": sku,
            "orders": total, "orders_prev": total_prev,
            "ad_orders": min(ad_orders, total), "organic": max(0.0, total - ad_orders),
            "organic_prev": max(0.0, total_prev - ad_orders_prev),
            "stock": stock.get(offer),
            "days_since_sale": (today - last.date()).days if last is not None and not pd.isna(last) else None,
            "locality": locality.get(sku) if sku else None,
            "avg_price": (price_history.get(offer) or {}).get("current"),
            "avg_price_prev": (price_history.get(offer) or {}).get("previous"),
            "price_now": (prices.get(offer) or {}).get("price"),
            "min_price": min_prices.get(offer),
            "price_index": (prices.get(offer) or {}).get("index"),
            "actions": (prices.get(offer) or {}).get("actions"),
            "content_rating": (rating.get(sku) or {}).get("rating") if sku else None,
            "content_missing": (rating.get(sku) or {}).get("missing") if sku else None,
            "ads": ad_cur, "search": search.get(sku) if sku else None,
        }
        p.update(analyze_product(p, store_ctx))
        p["events"] = snapshot_events(snapshots, offer)
        products.append(p)
    products.sort(key=lambda p: (-p["priority"], -p["orders"]))

    weekly_total = orders[orders["date"] >= pd.Timestamp(_week_start(today) - timedelta(weeks=WEEKS_IN_CHART - 1))]
    weekly_total = weekly_total.groupby(weekly_total["date"].dt.to_period("W-SUN").dt.start_time.dt.date)["qty"].sum()
    ads_map = dict(ads_weekly)
    weeks = [{"week": week.isoformat(), "total": float(units), "ads": float(ads_map.get(week, 0.0))} for week, units in weekly_total.items()]

    total, total_prev = _units(orders, spans["current"]), _units(orders, spans["previous"])
    ad_total = sum(a["orders"] + a["models"] for a in ads["current"].values())
    ad_total_prev = sum(a["orders"] + a["models"] for a in ads["previous"].values())
    search_funnel = None
    if premium:
        search_funnel = {"searched": sum((v.get("searched") or 0) for v in search.values()),
                         "viewed": sum((v.get("viewed") or 0) for v in search.values()),
                         "ordered": sum((q.get("orders") or 0) for v in search.values() for q in v.get("queries") or [])}
    result = {
        "created": datetime.now().isoformat(timespec="seconds"),
        "days": days,
        "periods": {k: [v[0].isoformat(), v[1].isoformat()] for k, v in spans.items()},
        "store": {
            "orders": total, "orders_prev": total_prev,
            "ad_orders": min(ad_total, total), "ad_orders_prev": min(ad_total_prev, total_prev),
            "organic": max(0.0, total - ad_total), "organic_prev": max(0.0, total_prev - ad_total_prev),
            "season_change": store_ctx["season_change"],
            "ad_funnel": {k: bench[k] for k in ("views", "clicks", "to_cart", "orders")},
            "search_funnel": search_funnel,
            "premium": premium,
        },
        "weeks": weeks,
        "products": products,
        "sources": sources,
        "ads_note": ads_note,
    }
    folder = root / HISTORY_DIR
    folder.mkdir(parents=True, exist_ok=True)
    (folder / LAST_RESULT).write_text(json.dumps(result, ensure_ascii=False, default=str), encoding="utf-8")
    return result


def load_last_result(root: Path) -> Optional[Dict[str, Any]]:
    path = root / HISTORY_DIR / LAST_RESULT
    try:
        return json.loads(path.read_text(encoding="utf-8")) if path.exists() else None
    except (OSError, ValueError):
        return None


def snapshot_after_sync(root: Path) -> None:
    """Снимок после синхронизации цен: цена, индекс, акции, остаток — чтобы копить свою историю."""
    import ads_cpc
    from recommended_prices import load_costs_df, COSTS_FILENAME, _artikul_normalize, _normalize_offer_id
    df, key_col, _ = load_costs_df(root / COSTS_FILENAME)
    offers = sorted({_normalize_offer_id(a) for a in (_artikul_normalize(v) for v in df[key_col]) if a})
    prices = load_prices(offers)
    stock = ads_cpc.load_available_stock(offers)
    today = date.today()
    existing = next((data for day, data in load_snapshots(root, today) if day == today), {})
    snapshot = {}
    for offer in offers:
        previous = existing.get(offer, {})
        snapshot[offer] = {"price": prices.get(offer, {}).get("price"), "stock": stock.get(offer),
                           "index": (prices.get(offer, {}).get("index") or {}).get("color"),
                           "actions": len(prices.get(offer, {}).get("actions") or []) if offer in prices else None,
                           "rating": previous.get("rating")}
    save_snapshot(root, today, snapshot)


if __name__ == "__main__":
    data = run_analysis(repo_root, progress=print)
    s = data["store"]
    print(f"Заказов {s['orders']:.0f} (было {s['orders_prev']:.0f}), органика {s['organic']:.0f} (было {s['organic_prev']:.0f})")
    for p in data["products"][:10]:
        print(p["offer_id"], p["main_problem"], "|", (p["advice"] or ["—"])[0])
