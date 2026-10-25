# -*- coding: utf-8 -*-
"""
Предельная цена товара в акциях Ozon (раздел «Акции», ползунок по каждому артикулу).

С октября 2026 участие в «Эластичном бустинге» и «Скидке на сток» задаёт одна цена товара — предельная:
товар в акции, если она не выше порога акции, и выходит, если выше. Порог:
- «Эластичный бустинг» — цена минимального бустинга (price_min_elastic);
- остальные акции — максимальная цена акции (max_action_price).
Цена одна на все акции, поэтому её изменение меняет участие сразу во всех акциях товара — модуль всегда считает
итоговый набор (`evaluate`): кто войдёт, кто выйдет, какой будет бустинг.

«Участвуют сейчас» и «участвуют с (дата)» — независимые состояния со своими ценами: изменение текущей предельной
цены не меняет участие в акциях, которые ещё не начались. Поэтому будущие акции ползунок не трогает: участие в них
показывается со своей ценой (и предупреждением, если она ниже минимальной), в расчёт и отправку они не входят.

Отправка (`apply_prices`): выбранная цена уходит в каждую идущую акцию, где товар уже участвует или куда должен войти
(/v1/actions/products/update). Если таких акций у товара нет (все пороги ниже — товар вне акций), цена уходит как цена
товара (/v1/product/import/prices, `import_product_prices`) — иначе её нечем передать и она не менялась бы. Так работает и при новой логике Ozon, и у аккаунтов, где старое явное
добавление ещё действует. После отправки участие перечитывается и сверяется с ожидаемым.

Баланс прибыли и продаж («Подобрать», `balance_price`): цена в пределах [минимальная, желательная], при которой
больше всего ожидаемой прибыли за период:
    оценка = прибыль с продажи × (1 + бустинг/100) × (текущая цена / цена) ^ PRICE_ELASTICITY
- прибыль с продажи — как «Ожидаемая чистая прибыль» в «Ценах и скидках»: цена × (1 − комиссия) − логистика − себестоимость
  (комиссия и логистика — из отчёта за прошлый месяц) плюс баллы Ozon за скидку («Супербустинг»: цена × 3%), пока товар
  в такой акции;
- бустинг — общий: наибольший из идущих акций, где товар будет при этой цене (бустинги акций не складываются —
  так показывает и кабинет Ozon: «акционный бустинг» = наибольший; надбавки Ozon за цель и индекс цен сюда не входят);
- распродажа (`is_sale`: «Чёрная пятница», «11.11», «Распродажа» и т. п. в названии) даёт ещё SALE_BONUS % к продажам —
  подборки, баннеры и спрос в дни распродажи; несколько распродаж сразу не складываются;
  1 + (бустинг + распродажа)/100 — во сколько раз больше продаж;
- последний множитель — покупатели охотнее берут дешевле: при эластичности 1 цена −10% ≈ продаж +10%, при 2 — +20%.
Это оценка, а не прогноз: эластичность выбирает пользователь (ELASTICITY_PRESETS), замеров по магазину нет.
На данных магазина 23.10.2026: при 1 треть артикулов уходит к желательной цене без бустинга, при 2 — ни один,
цены тянутся к порогам акций с высоким бустингом.

Чего модуль не делает:
- не восстанавливает цену после конца акции — у Ozon она остаётся, какой её поставили; поэтому подбор
  (`auto_price`) не опускает цену ниже вашей минимальной;
- не меняет «ограничение для акций» (минимальную цену в Ozon). Если поставить цену ниже него, Ozon сам опустит
  ограничение до этой цены — интерфейс об этом предупреждает.
"""

import math
import sys
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

script_dir = Path(__file__).resolve().parent
for path in (str(script_dir.parent), str(script_dir)):
    if path not in sys.path:
        sys.path.insert(0, path)

from price_management import _load_plan_context, _moscow_date, _moscow_today, _plan_entry_prices
from recommended_prices import (
    ACTION_TYPE_ELASTIC,
    COSTS_FILENAME,
    _artikul_normalize,
    _normalize_offer_id,
    OZON_HEADERS,
    activate_products_in_action,
    boost_at_price,
    fetch_action_products,
    get_product_prices_from_ozon,
    get_prev_month_year,
    get_report_path,
    load_cost_structure,
    load_costs_df,
    pick_action_price,
)

PRICE_ELASTICITY = 2.0
SALE_BONUS = 30.0
SALE_KEYWORDS = ("чёрная пятница", "черная пятница", "black friday", "распродаж", "11.11", "12.12", "киберпонедельник")
ELASTICITY_PRESETS = {1.0: "слабая", 2.0: "средняя", 3.0: "сильная"}
SETTINGS_FILENAME = "limit_price_settings.json"
BOOST_KEYS = ("min_boost", "max_boost", "price_min_elastic", "price_max_elastic", "current_boost")


def _number(value: Any) -> Optional[float]:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number if number > 0 else None


def threshold(action_info: Dict[str, Any], product: Dict[str, Any]) -> Optional[float]:
    """Порог участия: не выше него — товар в акции."""
    if action_info.get("action_type") == ACTION_TYPE_ELASTIC:
        return _number(product.get("price_min_elastic")) or _number(product.get("max_action_price"))
    return _number(product.get("max_action_price"))


def is_sale(name: str) -> bool:
    """Акция-распродажа (по названию): кроме бустинга даёт трафик из подборок и спрос дня распродажи."""
    lowered = (name or "").lower()
    return any(keyword in lowered for keyword in SALE_KEYWORDS)


def _entry(action_info: Dict[str, Any], product: Dict[str, Any], participating: bool, today: str) -> Dict[str, Any]:
    start = _moscow_date(action_info.get("date_start"))
    limit = threshold(action_info, product)
    action_price = _number(product.get("action_price")) if participating else None
    # Пока Ozon переводит акции на новую логику, товар бывает в акции по цене выше её порога (22–23.10.2026: 1102 в
    # «Эластичном» по 1 421 ₽ при пороге 1 278 ₽). Не противоречим Ozon: по нынешней цене товар в акции.
    above_limit = bool(limit and action_price and not (start and start > today) and action_price > limit)
    return {
        "action_id": action_info["id"],
        "name": action_info["name"],
        "action_type": action_info["action_type"],
        "fixed_boost": action_info.get("fixed_boost"),
        "sale": is_sale(action_info["name"]),
        "points": action_info.get("points_percent") or 0.0,
        "start": start if start and start > today else None,
        "threshold": action_price if above_limit else limit,
        "ozon_threshold": limit,
        "above_limit": above_limit,
        "participating": participating,
        "action_price": action_price,
        "added_by_ozon": participating and product.get("add_mode") == "AUTO",
        # Ozon помечает товар, если скидка перестала соответствовать условиям: 3 рабочих дня на исправление.
        "alert": bool(product.get("alert_max_action_price_failed")),
        "stock": int(product.get("stock") or 0),
        **{key: product.get(key) for key in BOOST_KEYS},
    }


def build_board(context: Dict[str, Any], loaded: Dict[Any, tuple], today: str,
                economics: Optional[Dict[str, Dict[str, float]]] = None) -> List[Dict[str, Any]]:
    """
    context: limits, actions, pid_to_offer (как у планов акций); loaded: {action_id: (участники, кандидаты)};
    economics: {артикул: {cost, commission, logistics}} — для прибыли с продажи (может не быть).
    Возвращает артикулы из costs.xlsx, которым доступна хоть одна акция, с текущей ценой и списком акций.
    """
    limits, pid_to_offer = context["limits"], context["pid_to_offer"]
    articles: Dict[str, Dict[str, Any]] = {}
    for action_info in context["actions"]:
        participants, candidates = loaded.get(action_info["id"], ([], []))
        for participating, products in ((True, participants), (False, candidates)):
            for product in products:
                offer_id = pid_to_offer.get(product.get("id"))
                if not offer_id or not (limits.get(offer_id) or (None,))[0]:
                    continue  # без минимальной цены в costs.xlsx не на что опираться
                entry = _entry(action_info, product, participating, today)
                if entry["threshold"] is None:
                    continue
                min_price, desired_price = limits[offer_id]
                article = articles.setdefault(offer_id, {
                    "offer_id": offer_id,
                    "product_id": product["id"],
                    "price": _number(product.get("price")),
                    "ozon_min": _number(product.get("min_seller_price")),
                    "min_price": min_price,
                    "desired_price": desired_price,
                    "economics": (economics or {}).get(offer_id),
                    "actions": [],
                })
                article["price"] = article["price"] or _number(product.get("price"))
                article["actions"].append(entry)
    for article in articles.values():
        article["actions"].sort(key=lambda a: -(a["threshold"] or 0))
    return sorted(articles.values(), key=lambda a: (len(a["offer_id"]), a["offer_id"]))


def load_board(repo_root: Path) -> Dict[str, Any]:
    """Все акции с участниками и кандидатами из Ozon (только чтение)."""
    context, error_plan = _load_plan_context(repo_root, need_desired=False)
    if error_plan:
        return {"ok": False, "message": error_plan["message"], "articles": []}

    def load(action_info):
        return action_info["id"], (fetch_action_products(action_info["id"], "products"), fetch_action_products(action_info["id"], "candidates"))

    with ThreadPoolExecutor(max_workers=4) as pool:
        loaded = dict(pool.map(load, context["actions"]))
    economics, economics_note = load_economics(repo_root)
    articles = build_board(context, loaded, _moscow_today(), economics)
    return {"ok": True, "loaded": datetime.now().isoformat(timespec="seconds"), "articles": articles,
            "economics_note": economics_note, "message": f"Артикулов с доступными акциями: {len(articles)}."}


def load_economics(repo_root: Path) -> tuple:
    """
    Себестоимость из costs.xlsx и расходы Ozon (комиссия, логистика за штуку) из отчёта за прошлый месяц —
    как в синхронизации «Цен и скидок». Нет отчёта — прибыль не считаем: ({}, причина).
    """
    year, month = get_prev_month_year()
    report = get_report_path(repo_root, year, month)
    try:
        structure = load_cost_structure(report)
    except Exception as exc:
        return {}, f"Прибыль не считается: нет расходов Ozon за прошлый месяц ({exc}). Сформируйте отчёт в «Бизнес-сводке»."
    df, key_col, cost_col = load_costs_df(repo_root / COSTS_FILENAME)
    economics: Dict[str, Dict[str, float]] = {}
    for raw_article, raw_cost in zip(df[key_col], df[cost_col]):
        article = _artikul_normalize(raw_article)
        cost = _number(raw_cost)
        if article and cost:
            economics[_normalize_offer_id(article)] = {"cost": cost, "commission": structure.commission_rate,
                                                       "logistics": structure.logistics_for(article)}
    return economics, f"Расходы Ozon — из отчёта «{report.name}»."


def boost(entry: Dict[str, Any], price: float) -> Optional[float]:
    low, high = _number(entry.get("price_max_elastic")), _number(entry.get("price_min_elastic"))
    if entry["action_type"] == ACTION_TYPE_ELASTIC and low and high and low == high:
        # Шкала эластичного бустинга схлопнулась (так Ozon присылает с 23.10.2026): рассчитать бустинг по цене нельзя —
        # берём тот, что называет Ozon для участника, иначе минимальный.
        reported = _number(entry.get("current_boost")) if entry.get("participating") else None
        return reported or _number(entry.get("min_boost"))
    return boost_at_price({"action_type": entry["action_type"], "fixed_boost": entry["fixed_boost"]}, entry, price)


def evaluate(article: Dict[str, Any], price: float) -> List[Dict[str, Any]]:
    """
    Участие при этой цене: in, boost и change — enter / exit / stay / none; для акций, которые ещё не начались, — future
    (участие по своей цене, ползунок его не меняет; у кандидата in показывает, войдёт ли он при этой цене).
    """
    result = []
    for entry in article["actions"]:
        if entry["start"] and entry["participating"]:
            own = entry["action_price"] or entry["threshold"]
            result.append({**entry, "in": True, "change": "future", "boost": boost(entry, own)})
            continue
        if entry["start"]:
            inside = price <= entry["threshold"]
            result.append({**entry, "in": inside, "change": "future", "boost": boost(entry, price) if inside else None})
            continue
        inside = price <= entry["threshold"]
        if inside:
            change = "stay" if entry["participating"] else "enter"
        else:
            change = "exit" if entry["participating"] else "none"
        result.append({**entry, "in": inside, "change": change, "boost": boost(entry, price) if inside else None})
    return result


def points_percent(article: Dict[str, Any], price: float) -> float:
    """
    Баллы Ozon за скидку, % от цены продажи: наибольший из идущих акций с баллами, где товар будет при этой цене
    (продажа одна — баллы разных акций не складываем). Баллами оплачиваются комиссия и услуги Ozon — это те же деньги.
    """
    return max((entry.get("points") or 0.0 for entry in article["actions"]
                if not entry["start"] and price <= entry["threshold"]), default=0.0)


def unit_profit(article: Dict[str, Any], price: float) -> Optional[float]:
    """
    Чистая прибыль с продажи: цена × (1 − комиссия) − логистика − себестоимость + баллы Ozon за скидку (цена × %).
    None — нет данных.
    """
    economics = article.get("economics")
    if not economics:
        return None
    base = price * (1 - economics["commission"]) - economics["logistics"] - economics["cost"]
    return base + price * points_percent(article, price) / 100


def total_boost(article: Dict[str, Any], price: float) -> float:
    """Общий бустинг: наибольший из идущих акций, где товар будет при этой цене (будущие акции ещё не действуют)."""
    boosts = [boost(entry, price) or 0.0 for entry in article["actions"] if not entry["start"] and price <= entry["threshold"]]
    return max(boosts, default=0.0)


def sale_bonus(article: Dict[str, Any], price: float) -> float:
    """SALE_BONUS, если при этой цене товар в идущей распродаже; несколько распродаж не складываются."""
    return SALE_BONUS if any(entry.get("sale") and not entry["start"] and price <= entry["threshold"] for entry in article["actions"]) else 0.0


def sales_factor(article: Dict[str, Any], price: float, elasticity: float = PRICE_ELASTICITY) -> float:
    """Во сколько раз продажи при этой цене больше, чем у текущей цены без акций (оценка)."""
    reference = article.get("price") or article.get("min_price") or price
    return (1 + (total_boost(article, price) + sale_bonus(article, price)) / 100) * (reference / price) ** elasticity


def balance_score(article: Dict[str, Any], price: float, elasticity: float = PRICE_ELASTICITY) -> Optional[float]:
    profit = unit_profit(article, price)
    return None if profit is None else profit * sales_factor(article, price, elasticity)


def balance_price(article: Dict[str, Any], elasticity: float = PRICE_ELASTICITY) -> Optional[int]:
    """
    Цена в пределах [минимальная, желательная] с наибольшей оценкой «прибыль × продажи»; при равенстве — выше.
    Без экономики (нет отчёта или себестоимости) — прежнее правило auto_price.
    """
    low, high = article.get("min_price"), article.get("desired_price") or article.get("min_price")
    if not low or unit_profit(article, low) is None:
        return auto_price(article)
    best, best_score = None, None
    for price in range(math.ceil(low), math.floor(high) + 1):
        score = balance_score(article, price, elasticity)
        if score is not None and (best_score is None or score >= best_score):
            best, best_score = price, score
    return best if best is not None else auto_price(article)


def auto_price(article: Dict[str, Any]) -> Optional[int]:
    """
    Подбор в пределах [минимальная, желательная]: как можно ближе к желательной, не выходя из акций, где товар уже
    участвует по цене не ниже минимальной, и не опуская цену заметно ради одной новой акции (_plan_entry_prices).
    None — если минимальная цена в costs.xlsx не задана.
    """
    low, high = article["min_price"], article["desired_price"] or article["min_price"]
    if not low:
        return None
    entries = [(article["product_id"], entry["action_id"], low, high, entry["threshold"],
                entry["participating"] and (entry["action_price"] or 0) >= low)
               for entry in article["actions"] if not entry["start"]]
    price = _plan_entry_prices(entries)["price"].get(article["product_id"])
    if price is None:
        price = pick_action_price(low, high, high)
    return max(math.ceil(low), min(int(price), math.floor(high))) if price is not None else math.ceil(low)


def plan_changes(articles: List[Dict[str, Any]], prices: Dict[str, int]) -> List[Dict[str, Any]]:
    """
    Что отправить: по каждому артикулу с новой ценой — в каждую акцию, где он участвует или куда должен войти.
    Акции, где товар уже участвует ровно по этой цене, не трогаем (Ozon ответил бы «нет изменений»).
    """
    items = []
    for article in articles:
        price = prices.get(article["offer_id"])
        if price is None:
            continue
        for entry in evaluate(article, price):
            if entry["change"] in ("none", "future"):
                continue
            if entry["change"] == "stay" and entry["action_price"] is not None and abs(entry["action_price"] - price) < 0.5:
                continue
            items.append({"offer_id": article["offer_id"], "product_id": article["product_id"], "action_id": entry["action_id"],
                          "action_name": entry["name"], "price": price, "expect_in": entry["in"], "change": entry["change"],
                          "stock": entry["stock"]})
    return items


RECHECK_DELAYS = (3, 7)  # секунд: Ozon применяет выход из акций не мгновенно
NO_CHANGES = "no changes"  # «Product update has no changes»: цена уже такая — её поставила отправка в другую акцию


def ozon_notice(reason: str, item: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """
    Предупреждение Ozon из ответа на отправку → понятное пояснение. kind: limit — Ozon опустил «ограничение для акций»
    (минимальную цену в Ozon) до новой цены; left — товар вышел из акции (ожидаемо при выходе, иначе — неожиданно).
    """
    import re
    text = reason or ""
    match = re.search(r"min price[^=]*=\s*([\d.]+)", text, flags=re.IGNORECASE)
    if match:
        value = float(match.group(1))
        return {"kind": "limit", "offer_id": item["offer_id"], "value": value,
                "text": f"{item['offer_id']}: Ozon опустил ограничение для акций до {value:,.0f} ₽".replace(",", " ")
                        + " — теперь он сам может добавлять товар в акции по цене до этой суммы"}
    if "deactivated" in text.lower() or "no longer meets" in text.lower():
        expected = not item["expect_in"]
        return {"kind": "left", "offer_id": item["offer_id"], "expected": expected,
                "text": f"{item['offer_id']}: вышел из «{item['action_name']}»" + ("" if expected else " — неожиданно, проверьте цену")}
    return {"kind": "other", "offer_id": item["offer_id"], "text": f"{item['offer_id']} · {item['action_name']}: {text}"}


def direct_price_articles(articles: List[Dict[str, Any]], prices: Dict[str, int], items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Артикулы с новой ценой, которую не через что передать в акциях (ни одной идущей акции, где товар есть или войдёт)."""
    via_actions = {item["offer_id"] for item in items}
    return [article for article in articles
            if article["offer_id"] in prices and article["offer_id"] not in via_actions
            and prices[article["offer_id"]] != (round(article["price"]) if article.get("price") else None)]


def import_product_prices(rows: List[Dict[str, Any]]) -> Dict[int, str]:
    """/v1/product/import/prices — только поле price (старую цену и ограничение не трогаем). {product_id: ошибка}."""
    import requests
    errors: Dict[int, str] = {}
    for start in range(0, len(rows), 1000):
        body = requests.post("https://api-seller.ozon.ru/v1/product/import/prices", headers=OZON_HEADERS, timeout=30,
                             json={"prices": [{"offer_id": row["offer_id"], "product_id": row["product_id"], "price": str(row["price"])}
                                              for row in rows[start:start + 1000]]})
        if body.status_code != 200:
            for row in rows[start:start + 1000]:
                errors[row["product_id"]] = f"HTTP {body.status_code}: {body.text[:200]}"
            continue
        for result in body.json().get("result") or []:
            if not result.get("updated"):
                messages = "; ".join(str(e.get("message") or e.get("code")) for e in result.get("errors") or [])
                errors[int(result.get("product_id") or 0)] = messages or "Ozon не применил цену"
    return errors


def apply_prices(articles: List[Dict[str, Any]], prices: Dict[str, int], sleep=None) -> Dict[str, Any]:
    """
    Отправляет цены и сверяет участие по каждой затронутой акции. «Нет изменений» от Ozon — не ошибка, если товар
    уже в нужном состоянии (цена одна на все акции: её мог поставить запрос в предыдущую акцию). Предупреждения Ozon
    (сдвиг ограничения для акций, выход из акции) возвращаются в notices — по одному на артикул и вид.
    """
    import time
    sleep = sleep or time.sleep
    items = plan_changes(articles, prices)
    by_action: Dict[Any, List[Dict[str, Any]]] = {}
    for item in items:
        by_action.setdefault(item["action_id"], []).append(item)
    done, failed, errors, notices = 0, [], [], []
    sent: List[tuple] = []
    seen = set()
    for action_id, action_items in by_action.items():
        result = activate_products_in_action(action_id, [
            {"product_id": item["product_id"], "action_price": item["price"], "stock": item["stock"]} for item in action_items
        ])
        by_product = {item["product_id"]: item for item in action_items}
        rejected = {int(entry.get("product_id") or 0): entry.get("reason") or "причина не указана" for entry in result.get("rejected") or []}
        errors += [f"{action_items[0]['action_name']}: {error}" for error in result.get("errors") or []]
        for warning in result.get("warnings") or []:
            item = by_product.get(int(warning.get("product_id") or 0))
            notice = ozon_notice(str(warning.get("reason") or ""), item) if item else None
            key = notice and (notice["offer_id"], notice["kind"], notice.get("value"), notice["text"] if notice["kind"] != "limit" else None)
            if notice and key not in seen:
                seen.add(key)
                notices.append(notice)
        sent.append((action_id, action_items, rejected))

    # Сверка по факту. Ozon применяет выход из акций с задержкой в несколько секунд (5004, 23.10.2026: в «11.11» сразу
    # после запроса ещё был, через минуту — нет), поэтому то, что не на месте, перепроверяем после пауз.
    def in_place(item: Dict[str, Any], actual: Dict[Any, Dict[str, Any]]) -> bool:
        product = actual.get(item["product_id"])
        if item["expect_in"]:
            return product is not None and abs(float(product.get("action_price") or 0) - item["price"]) < 0.5
        return product is None

    pending = [(action_id, action_items, rejected) for action_id, action_items, rejected in sent]
    settled: Dict[Any, Dict[Any, Dict[str, Any]]] = {}
    for delay in (0, *RECHECK_DELAYS):
        if delay:
            sleep(delay)
        still = []
        for action_id, action_items, rejected in pending:
            actual = {product["id"]: product for product in fetch_action_products(action_id, "products")}
            settled[action_id] = actual
            waiting = [item for item in action_items if not in_place(item, actual)
                       and (item["product_id"] not in rejected or NO_CHANGES in rejected[item["product_id"]].lower())]
            if waiting:
                still.append((action_id, action_items, rejected))
        pending = still
        if not pending:
            break
    for action_id, action_items, rejected in sent:
        actual = settled.get(action_id, {})
        for item in action_items:
            reason = rejected.get(item["product_id"])
            placed = in_place(item, actual)
            if reason and not (NO_CHANGES in reason.lower() and placed):
                failed.append({**item, "reason": f"Ozon отклонил: {reason}"})
            elif not placed:
                failed.append({**item, "reason": "товар не в акции по этой цене после отправки" if item["expect_in"] else "товар остался в акции"})
            else:
                done += 1

    # Товары вне акций: цена товара напрямую, затем проверка по /v5/product/info/prices.
    direct = direct_price_articles(articles, prices, items)
    if direct:
        rows = [{"offer_id": a["offer_id"], "product_id": a["product_id"], "price": prices[a["offer_id"]]} for a in direct]
        try:
            import_errors = import_product_prices(rows)
        except Exception as exc:
            import_errors = {row["product_id"]: str(exc) for row in rows}
        actual_prices, _ = get_product_prices_from_ozon([row["offer_id"] for row in rows])
        for row in rows:
            item = {**row, "action_id": None, "action_name": "цена товара", "expect_in": False, "change": "price", "stock": 0}
            items.append(item)
            actual = actual_prices.get(row["offer_id"])
            if row["product_id"] in import_errors:
                failed.append({**item, "reason": f"Ozon отклонил: {import_errors[row['product_id']]}"})
            elif actual is None or abs(actual - row["price"]) >= 0.5:
                failed.append({**item, "reason": f"цена в Ozon после отправки — {actual}" if actual is not None else "Ozon не вернул цену"})
            else:
                done += 1
    message = f"Применено: {done} из {len(items)} (проверено по данным Ozon)."
    if failed:
        message += f" Не применилось: {len(failed)}."
    if errors:
        message += f" Ошибок запросов: {len(errors)}."
    lowered = [n for n in notices if n["kind"] == "limit"]
    if lowered:
        message += f" Ozon опустил ограничение для акций у {len(lowered)} арт."
    return {"ok": done == len(items) and not errors, "done": done, "expected": len(items), "failed": failed, "errors": errors,
            "notices": notices, "message": message, "items": items}


def load_elasticity(root: Path) -> float:
    """Выбранная реакция покупателей на цену (limit_price_settings.json); по умолчанию — средняя."""
    import json
    try:
        value = float(json.loads((root / SETTINGS_FILENAME).read_text(encoding="utf-8")).get("elasticity"))
    except (OSError, ValueError, TypeError, AttributeError):
        return PRICE_ELASTICITY
    return value if value in ELASTICITY_PRESETS else PRICE_ELASTICITY


def save_elasticity(root: Path, value: float) -> None:
    import json
    (root / SETTINGS_FILENAME).write_text(json.dumps({"elasticity": value}, ensure_ascii=False, indent=2), encoding="utf-8")
