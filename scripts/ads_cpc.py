# -*- coding: utf-8 -*-
"""
Реклама «Оплата за клик»: подбор ставок по экономике товара.

Логика:
1. Прибыль с продажи до рекламы и себестоимость — из costs.xlsx («Ожидаемая чистая прибыль»).
2. Рекламируем только товары с положительной прибылью, которые есть в наличии.
3. Конверсия клика в заказ по товару сглаживается к средней по кампании:
   (заказы + K·конверсия кампании) / (клики + K) — у одного товара заказов мало.
4. Безубыточная ставка = конверсия × прибыль. Целевая = безубыточная × доля прибыли на рекламу.
5. Ставка меняется не больше чем на ±25% за раз и не ниже минимальной ставки Ozon.
Ставки товаров Ozon учитывает только в стратегии «Средняя стоимость клика» (TARGET_BIDS).
"""

import json
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests

script_dir = Path(__file__).resolve().parent
repo_root = script_dir.parent
if str(script_dir) not in sys.path:
    sys.path.insert(0, str(script_dir))

PERF_HOST = "https://api-performance.ozon.ru"
SETTINGS_FILENAME = "ads_settings.json"
ADS_REPORTS_DIR = "ads reports"
CHANGES_LOG = "changes.jsonl"

DEFAULT_SETTINGS: Dict[str, Any] = {
    "profit_share": 0.4,          # какую долю прибыли с заказа можно отдать на рекламу
    "lookback_days": 21,          # период статистики
    "include_model_orders": True,  # учитывать покупки соседних вариантов после клика
}

PRIOR_CLICKS = 60          # вес средней конверсии кампании при сглаживании
MAX_STEP = 0.25            # максимальное изменение ставки за один раз
MIN_CHANGE_RUB = 0.10      # меньшие изменения не предлагаем
NO_ORDERS_EXPECTED = 2.0   # «нет заказов» считаем сигналом, если ожидалось хотя бы столько
FALLBACK_CR = 0.02
MICRO = 1_000_000          # ставки и бюджеты в API — в миллионных долях рубля
RETRY_ATTEMPTS = 9         # повторы при 429: 5, 10, 15 … 30 секунд
STATS_MAX_DAYS = 62       # Ozon отдаёт статистику кампаний не больше чем за 62 дня в запросе
SUMMABLE = ("moneySpent", "views", "clicks", "orders", "ordersMoney", "toCart")
STATE_NAMES = {
    "CAMPAIGN_STATE_RUNNING": "Активна",
    "CAMPAIGN_STATE_STOPPED": "Остановлена",
    "CAMPAIGN_STATE_INACTIVE": "Неактивна",
    "CAMPAIGN_STATE_PLANNED": "Запланирована",
    "CAMPAIGN_STATE_FINISHED": "Завершена",
    "CAMPAIGN_STATE_ARCHIVED": "В архиве",
    "CAMPAIGN_STATE_MODERATION_DRAFT": "Черновик",
    "CAMPAIGN_STATE_MODERATION_IN_PROGRESS": "На модерации",
    "CAMPAIGN_STATE_MODERATION_FAILED": "Не прошла модерацию",
}
STRATEGY_NAMES = {
    "TARGET_BIDS": "Средняя стоимость клика",
    "MAX_CLICKS": "Максимум кликов",
    "TOP_MAX_CLICKS": "Максимум кликов в поиске",
    "TARGET_CIR": "Целевой расход",
    "TOP_PROMOTION": "Вывод в топ",
}


# ---------------------------------------------------------------- настройки

def load_settings(root: Path) -> Dict[str, Any]:
    settings = dict(DEFAULT_SETTINGS)
    path = root / SETTINGS_FILENAME
    if path.exists():
        try:
            settings.update(json.loads(path.read_text(encoding="utf-8")))
        except (OSError, ValueError):
            pass
    return settings


def save_settings(root: Path, settings: Dict[str, Any]) -> None:
    (root / SETTINGS_FILENAME).write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")


# ---------------------------------------------------------------- Performance API

def _num(value: Any) -> float:
    try:
        return float(str(value if value is not None else 0).replace(",", ".").replace(" ", "").replace("\xa0", ""))
    except ValueError:
        return 0.0


class PerfClient:
    """Минимальный клиент Performance API с повтором при 429."""

    def __init__(self) -> None:
        from dotenv import load_dotenv
        import performance_api

        load_dotenv(repo_root / ".env")
        client_id, secret = performance_api.get_performance_api_credentials()
        if not client_id or not secret:
            raise RuntimeError("Не заданы OZON_PERF_CLIENT_ID и OZON_PERF_API_KEY — настройте подключение к рекламному кабинету.")
        self.session = requests.Session()
        token = performance_api.get_performance_token(self.session, client_id, secret)
        if not token:
            raise RuntimeError("Не удалось получить токен Performance API.")
        self.headers = {"Authorization": f"Bearer {token}", "Accept": "application/json", "Content-Type": "application/json"}

    def call(self, method: str, path: str, **kwargs) -> Any:
        # Ozon разрешает только один формирующийся отчёт статистики одновременно и ограничивает частоту
        # запросов — при 429 ждём дольше (до ~3 минут суммарно), пока прошлый отчёт освободит очередь.
        for attempt in range(RETRY_ATTEMPTS):
            response = self.session.request(method, PERF_HOST + path, headers=self.headers, timeout=60, **kwargs)
            if response.status_code == 429 and attempt < RETRY_ATTEMPTS - 1:
                time.sleep(min(5 + attempt * 5, 30))
                continue
            if response.status_code >= 400:
                raise RuntimeError(f"{method} {path}: HTTP {response.status_code} {response.text[:300]}")
            return response.json() if response.text else {}
        raise RuntimeError(f"{method} {path}: слишком много запросов")

    def running_cpc_campaigns(self) -> List[Dict[str, Any]]:
        data = self.call("GET", "/api/client/campaign", params={"advObjectType": "SKU", "state": "CAMPAIGN_STATE_RUNNING"})
        return [c for c in data.get("list") or [] if (c.get("PaymentType") or c.get("paymentType") or "CPC") == "CPC"]

    def all_cpc_campaigns(self) -> List[Dict[str, Any]]:
        """Все кампании «Оплата за клик» в любом статусе."""
        data = self.call("GET", "/api/client/campaign", params={"advObjectType": "SKU"})
        return [c for c in data.get("list") or [] if (c.get("PaymentType") or c.get("paymentType") or "CPC") == "CPC"]

    def campaign_totals(self, campaign_ids: List[str], date_from: date, date_to: date) -> Dict[str, Dict[str, float]]:
        """
        Показатели кампаний за период. Ozon отдаёт не больше 62 дней за запрос —
        длинный период делим на части и складываем суммируемые показатели.
        """
        totals: Dict[str, Dict[str, float]] = {}
        start = date_from
        while start <= date_to:
            end = min(date_to, start + timedelta(days=STATS_MAX_DAYS))
            for i in range(0, len(campaign_ids), 50):
                data = self.call("GET", "/api/client/statistics/campaign/product/json", params={
                    "campaignIds": campaign_ids[i:i + 50], "dateFrom": start.isoformat(), "dateTo": end.isoformat(),
                })
                for row in data.get("rows") or []:
                    entry = totals.setdefault(str(row.get("id")), {key: 0.0 for key in SUMMABLE})
                    for key in SUMMABLE:
                        entry[key] += _num(row.get(key))
            start = end + timedelta(days=1)
        return totals

    def campaign_products(self, campaign_id: str) -> List[Dict[str, Any]]:
        products, page = [], 1
        while True:
            data = self.call("GET", f"/api/client/campaign/{campaign_id}/v2/products", params={"page": page, "pageSize": 100})
            batch = data.get("products") or []
            products.extend(batch)
            if len(batch) < 100:
                return products
            page += 1

    def sku_statistics(self, campaign_ids: List[str], date_from: date, date_to: date) -> Dict[str, List[Dict[str, Any]]]:
        """Статистика по товарам за период (асинхронный отчёт, до 10 кампаний за запрос)."""
        result: Dict[str, List[Dict[str, Any]]] = {}
        for i in range(0, len(campaign_ids), 10):
            chunk = campaign_ids[i:i + 10]
            created = self.call("POST", "/api/client/statistics/json", json={
                "campaigns": chunk, "dateFrom": date_from.isoformat(), "dateTo": date_to.isoformat(), "groupBy": "NO_GROUP_BY",
            })
            uuid = created.get("UUID")
            for _ in range(60):
                state = self.call("GET", f"/api/client/statistics/{uuid}").get("state")
                if state == "OK":
                    break
                if state == "ERROR":
                    raise RuntimeError("Ozon не смог сформировать отчёт по статистике.")
                time.sleep(3)
            else:
                raise RuntimeError("Отчёт по статистике не сформировался за 3 минуты.")
            report = self.call("GET", "/api/client/statistics/report", params={"UUID": uuid})
            if len(chunk) == 1 and "report" in report:
                report = {chunk[0]: report}
            for campaign_id in chunk:
                result[campaign_id] = ((report.get(campaign_id) or {}).get("report") or {}).get("rows") or []
        return result

    def competitive_bids(self, campaign_id: str, skus: List[str]) -> Dict[str, float]:
        bids: Dict[str, float] = {}
        for i in range(0, len(skus), 200):
            data = self.call("GET", f"/api/client/campaign/{campaign_id}/products/bids/competitive", params={"skus": skus[i:i + 200]})
            bids.update({str(b["sku"]): _num(b.get("bid")) / MICRO for b in data.get("bids") or []})
        return bids

    def min_bids(self, skus: List[str]) -> Dict[str, float]:
        if not skus:
            return {}
        data = self.call("POST", "/api/client/min/sku", json={"sku": skus})
        return {str(b["sku"]): _num(b.get("bid")) for b in data.get("minBids") or []}

    def daily_spend(self, campaign_ids: List[str], date_from: date, date_to: date) -> Dict[str, Dict[str, float]]:
        """{кампания: {дата: расход}} — для темпа расхода недельного бюджета."""
        data = self.call("GET", "/api/client/statistics/daily/json", params={
            "campaignIds": campaign_ids, "dateFrom": date_from.isoformat(), "dateTo": date_to.isoformat(),
        })
        result: Dict[str, Dict[str, float]] = {}
        for row in data.get("rows") or []:
            result.setdefault(str(row.get("id")), {})[str(row.get("date"))] = _num(row.get("moneySpent"))
        return result

    def update_bids(self, campaign_id: str, bids: Dict[str, float]) -> None:
        payload = {"bids": [{"sku": sku, "bid": str(to_micro(bid))} for sku, bid in bids.items()]}
        self.call("PUT", f"/api/client/campaign/{campaign_id}/products", json=payload)

    def remove_products(self, campaign_id: str, skus: List[str]) -> None:
        self.call("POST", f"/api/client/campaign/{campaign_id}/products/delete", json={"sku": skus})


def to_micro(rub: float) -> int:
    """Рубли -> миллионные доли рубля с округлением до копейки."""
    return int(round(rub * 100)) * 10_000


def list_cpc_campaigns(date_from: date, date_to: date) -> Dict[str, Any]:
    """
    Кампании «Оплата за клик» с показателями за период: работающие сейчас
    и все, у которых в этом периоде были показы или расход.
    """
    client = PerfClient()
    campaigns = client.all_cpc_campaigns()
    totals = client.campaign_totals([str(c["id"]) for c in campaigns], date_from, date_to) if campaigns else {}
    rows = []
    for campaign in campaigns:
        campaign_id = str(campaign["id"])
        t = totals.get(campaign_id, {key: 0.0 for key in SUMMABLE})
        state = campaign.get("state") or ""
        running = state == "CAMPAIGN_STATE_RUNNING"
        if not running and t["views"] <= 0 and t["moneySpent"] <= 0:
            continue
        rows.append({
            "id": campaign_id,
            "title": campaign.get("title") or f"Кампания {campaign_id}",
            "state": state,
            "state_name": STATE_NAMES.get(state, state.replace("CAMPAIGN_STATE_", "").lower()),
            "running": running,
            "strategy": STRATEGY_NAMES.get(campaign.get("productAutopilotStrategy") or "", campaign.get("productAutopilotStrategy") or "—"),
            "strategy_code": campaign.get("productAutopilotStrategy") or "",
            "weekly_budget": _num(campaign.get("weeklyBudget")) / MICRO,
            "daily_budget": _num(campaign.get("dailyBudget")) / MICRO,
            "from_date": campaign.get("fromDate") or "",
            "spend": t["moneySpent"],
            "views": t["views"],
            "clicks": t["clicks"],
            "to_cart": t["toCart"],
            "orders": t["orders"],
            "revenue": t["ordersMoney"],
            "ctr": _ratio(t["clicks"], t["views"]),
            "cpc": _ratio(t["moneySpent"], t["clicks"]),
            "drr": _ratio(t["moneySpent"], t["ordersMoney"]),
            "cpo": _ratio(t["moneySpent"], t["orders"]),
        })
    rows.sort(key=lambda r: (not r["running"], -r["spend"], r["title"]))
    return {"period": [date_from.isoformat(), date_to.isoformat()], "campaigns": rows, "loaded": datetime.now().isoformat(timespec="seconds")}


# ---------------------------------------------------------------- данные магазина

def load_unit_economics(root: Path) -> Dict[str, Dict[str, Optional[float]]]:
    """{артикул: {"cost", "profit", "price"}} из costs.xlsx."""
    from recommended_prices import load_costs_df, COSTS_FILENAME, COL_EXPECTED_PROFIT, COL_MARKETING_PRICE
    from recommended_prices import _artikul_normalize, _normalize_offer_id

    df, key_col, cost_col = load_costs_df(root / COSTS_FILENAME)
    if COL_EXPECTED_PROFIT not in df.columns:
        raise RuntimeError("В costs.xlsx нет «Ожидаемой чистой прибыли» — откройте вкладку «Цены и скидки», чтобы синхронизировать цены.")
    result = {}
    for _, row in df.iterrows():
        article = _artikul_normalize(row.get(key_col))
        if not article:
            continue
        values = {}
        for key, column in (("cost", cost_col), ("profit", COL_EXPECTED_PROFIT), ("price", COL_MARKETING_PRICE)):
            try:
                value = float(str(row.get(column)).replace(",", "."))
                values[key] = None if value != value else value
            except (TypeError, ValueError):
                values[key] = None
        result[_normalize_offer_id(article)] = values
    return result


def load_available_stock(offer_ids: List[str]) -> Dict[str, int]:
    """Сколько можно купить сейчас: свободный остаток FBO + FBS."""
    from recommended_prices import OZON_HEADERS, _normalize_offer_id

    result: Dict[str, int] = {}
    for i in range(0, len(offer_ids), 1000):
        cursor = ""
        while True:
            payload: Dict[str, Any] = {"filter": {"offer_id": offer_ids[i:i + 1000], "visibility": "ALL"}, "limit": 1000}
            if cursor:
                payload["cursor"] = cursor
            response = requests.post("https://api-seller.ozon.ru/v4/product/info/stocks", headers=OZON_HEADERS, json=payload, timeout=60)
            response.raise_for_status()
            data = response.json()
            items = data.get("items") or []
            for item in items:
                free = sum(int(s.get("present") or 0) - int(s.get("reserved") or 0) for s in item.get("stocks") or [])
                result[_normalize_offer_id(str(item.get("offer_id") or ""))] = max(0, free)
            cursor = data.get("cursor") or ""
            if not items or not cursor or len(items) < 1000:
                break
    return result


# ---------------------------------------------------------------- решение по товару

def smoothed_cr(orders: float, clicks: float, prior_cr: float) -> float:
    return (orders + PRIOR_CLICKS * prior_cr) / (clicks + PRIOR_CLICKS)


TEST_CLICKS = 150          # меньше кликов и меньше TEST_DAYS в кампании — товар в тесте
TEST_DAYS = 28
GENTLE_STEP_DOWN = 0.10    # в тесте и при большой органике снижаем осторожнее
FUNNEL_WEAK = 0.6          # показатель ниже 60% от среднего по кампании — слабое место
SEASON_RISE, SEASON_FALL = 1.3, 0.75
WEEKDAYS = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"]


def _pct(value: Optional[float], digits: int = 1) -> str:
    return "—" if value is None else f"{value * 100:.{digits}f}%".replace(".", ",")


def _rub(value: Optional[float], digits: int = 0) -> str:
    return "—" if value is None else f"{value:,.{digits}f} ₽".replace(",", " ").replace(".", ",")


def _ratio(numerator: float, denominator: float) -> Optional[float]:
    return numerator / denominator if denominator else None


def campaign_benchmarks(items: List[Dict[str, Any]]) -> Dict[str, Optional[float]]:
    """Средние показатели воронки по кампании — с ними сравниваем каждый товар."""
    views = sum(i["views"] for i in items)
    clicks = sum(i["clicks"] for i in items)
    carts = sum(i["to_cart"] for i in items)
    orders = sum(i["orders_total"] for i in items)
    return {"ctr": _ratio(clicks, views), "cart_rate": _ratio(carts, clicks), "cart_to_order": _ratio(orders, carts)}


def _step(name: str, status: str, text: str) -> Dict[str, str]:
    return {"name": name, "status": status, "text": text}


def analyze_item(item: Dict[str, Any], prior_cr: float, bench: Dict[str, Optional[float]], settings: Dict[str, Any]) -> Dict[str, Any]:
    """
    Решение по товару так, как его принимает маркетолог: шаг за шагом,
    с цифрами на каждом шаге и итоговыми рекомендациями.
    """
    share = float(settings["profit_share"])
    profit, cost, price = item.get("profit"), item.get("cost"), item.get("price")
    current = item["current_bid"]
    clicks, orders, spend = item["clicks"], item["orders_total"], item["spend"]
    steps: List[Dict[str, str]] = []
    advice: List[str] = []
    result = {"cr": smoothed_cr(orders, clicks, prior_cr), "breakeven_bid": None, "target_bid": None,
              "recommended_bid": current, "action": "keep", "reason": "", "stage": None,
              "ad_drr": _ratio(spend, item["revenue"]), "general_drr": _ratio(spend, item.get("product_gmv") or 0)}

    def finish(action: str, reason: str) -> Dict[str, Any]:
        result.update(action=action, reason=reason, steps=steps, advice=advice)
        return result

    # 1. Экономика
    if profit is None or cost is None:
        steps.append(_step("Экономика", "bad", "Нет прибыли или себестоимости в costs.xlsx — без них нельзя посчитать, сколько можно платить за клик."))
        advice.append("Заполните себестоимость и синхронизируйте цены на вкладке «Цены и скидки».")
        return finish("info", "Нет данных о прибыли товара.")
    margin = _ratio(profit, price) if price else None
    allowed_cpo = profit * share
    max_drr = margin * share if margin is not None else None
    passed = profit > 0
    steps.append(_step("Экономика", "ok" if passed else "bad",
        f"Прибыль с продажи до рекламы {_rub(profit)} при себестоимости {_rub(cost)}"
        + (f" и цене {_rub(price)} — маржа {_pct(margin)}" if margin is not None else "") + ". "
        + (f"На рекламу можно отдать {share:.0%} прибыли: до {_rub(allowed_cpo)} на заказ"
           + (f", то есть ДРР до {_pct(max_drr)}." if max_drr is not None else ".") if passed
           else "Товар продаётся в минус — любая реклама только увеличит убыток.")))
    if not passed:
        advice.append("Поднимите цену или снизьте себестоимость: пока продажа убыточна, рекламировать товар нельзя.")
        return finish("remove", f"Продажа убыточна: прибыль {_rub(profit)}.")

    # 2. Наличие
    stock = item.get("stock")
    if stock == 0:
        steps.append(_step("Наличие", "bad", "Товара нет в наличии ни на FBO, ни на FBS — Ozon не показывает рекламу, пока товар не пополнят."))
        advice.append("Пополните остаток: реклама без товара не работает, а после пополнения ставка останется прежней.")
        return finish("info", "Нет в наличии — ставку не меняем.")
    steps.append(_step("Наличие", "ok", f"В наличии {stock} шт." if stock is not None else "Остаток не удалось проверить."))

    # 3. Стадия
    days = item.get("days_in_campaign")
    in_test = clicks < TEST_CLICKS and (days is None or days < TEST_DAYS)
    days_text = f", в кампании {days} дн" if days is not None else ""
    result["stage"] = "test" if in_test else "mature"
    if in_test:
        steps.append(_step("Стадия", "info", f"Тест: {clicks:.0f} из {TEST_CLICKS} кликов{days_text}. Данных мало — не делаем резких выводов: "
                                              f"за отсутствие заказов не режем, снижаем не больше чем на {GENTLE_STEP_DOWN:.0%} за раз."))
    elif clicks < TEST_CLICKS:
        steps.append(_step("Стадия", "info", f"Мало трафика: {clicks:.0f} кликов за {days} дн. Тест окончен, но статистика скромная — "
                                              "конверсию сильнее подтягиваем к средней по кампании."))
    else:
        steps.append(_step("Стадия", "ok", f"Зрелый товар: {clicks:.0f} кликов{days_text} — данных достаточно для решений."))

    # 4. Воронка
    # Из корзины считаем только прямые заказы: покупки соседних вариантов в корзину этого товара не попадают.
    ctr, cart_rate, cart_to_order = _ratio(clicks, item["views"]), _ratio(item["to_cart"], clicks), _ratio(item.get("orders", orders), item["to_cart"])
    funnel = [f"CTR {_pct(ctr)} (по кампании {_pct(bench.get('ctr'))})",
              f"в корзину {_pct(cart_rate)} кликов ({_pct(bench.get('cart_rate'))})",
              f"из корзины покупают {_pct(cart_to_order, 0)} ({_pct(bench.get('cart_to_order'), 0)})"]
    weak = []
    if item["views"] >= 300 and ctr is not None and bench.get("ctr") and ctr < bench["ctr"] * FUNNEL_WEAK:
        weak.append("Покупатели видят товар, но редко кликают — проверьте главное фото, цену и рейтинг в выдаче.")
    # Если покупают в основном соседние варианты, корзина этого товара занижена — карточка тут ни при чём.
    variants = item.get("model_orders") or 0
    switch_to_variants = variants >= 1 and variants >= item.get("orders", 0)
    if not switch_to_variants:
        if clicks >= 40 and cart_rate is not None and bench.get("cart_rate") and cart_rate < bench["cart_rate"] * FUNNEL_WEAK:
            weak.append("Заходят в карточку, но не кладут в корзину — проблема в карточке: фото, описание, отзывы, размеры.")
        if item["to_cart"] >= 8 and cart_to_order is not None and bench.get("cart_to_order") and cart_to_order < bench["cart_to_order"] * FUNNEL_WEAK:
            weak.append("Кладут в корзину, но не покупают — сравните цену и сроки доставки с конкурентами.")
    variant_note = (f" Покупатели чаще выбирают соседний вариант ({variants:.0f} заказов), поэтому корзина и покупки именно этого товара выглядят ниже — "
                    "это работа рекламы на всю карточку, а не проблема.") if switch_to_variants else ""
    steps.append(_step("Воронка", "warn" if weak else "ok",
                       "; ".join(funnel) + "." + variant_note + (" " + " ".join(weak) if weak else (" Слабых мест не видно." if not variant_note else ""))))
    advice.extend(weak)

    # 5. Окупаемость
    cpo = _ratio(spend, orders)
    organic = bool(item.get("product_gmv") and item["revenue"] and item["product_gmv"] >= 2 * item["revenue"])
    pays = cpo is not None and cpo <= allowed_cpo
    payback = (f"Реклама на заказ {_rub(cpo)} при допустимых {_rub(allowed_cpo)}" if cpo is not None
               else f"Расход {_rub(spend)} — заказов с рекламы пока нет")
    payback += f"; ДРР рекламный {_pct(result['ad_drr'])}" + (f" (допустимо {_pct(max_drr)})" if max_drr is not None else "")
    payback += f", общий с органикой {_pct(result['general_drr'])}."
    if organic:
        payback += " У товара много органических продаж — реклама, вероятно, поддерживает его позицию в поиске."
    steps.append(_step("Окупаемость", "ok" if pays else ("info" if in_test or cpo is None else "warn"), payback))

    # 6. Ставка
    if item["strategy"] != "TARGET_BIDS":
        name = STRATEGY_NAMES.get(item["strategy"], item["strategy"])
        steps.append(_step("Ставка", "info", f"Ставки задаёт стратегия «{name}» — менять их по товарам нельзя."))
        advice.append("Чтобы управлять ставками по товарам, переключите кампанию на «Среднюю стоимость клика» в кабинете Ozon.")
        return finish("info", f"Стратегия «{name}»: ставками управляет Ozon.")
    breakeven = result["cr"] * profit
    target = breakeven * share
    clicks_per_order = 1 / result["cr"] if result["cr"] else None
    notes = [
        f"Заказ приходит примерно с каждого {clicks_per_order:.0f}-го клика (конверсия {_pct(result['cr'])}, сглажена к средней по кампании {_pct(prior_cr)}), "
        f"значит клик в среднем приносит {_rub(breakeven, 2)} прибыли — это безубыточная ставка: при ней реклама съедает всю прибыль.",
        f"Вы отдаёте на рекламу {share:.0%} прибыли, поэтому клик должен стоить не больше {share:.0%} от {_rub(breakeven, 2)} = {_rub(target, 2)}.",
    ]
    if result["cr"]:
        cost_per_order = current / result["cr"]
        notes.append(f"Сейчас клик стоит {_rub(current, 2)}: заказ обходится примерно в {_rub(cost_per_order)} рекламы — "
                     f"{_pct(cost_per_order / profit, 0)} прибыли при лимите {share:.0%} ({_rub(allowed_cpo)}).")
    expected = clicks * prior_cr
    if not in_test and orders == 0 and expected >= NO_ORDERS_EXPECTED:
        target = min(target, current * 0.5)
        notes.append(f"{clicks:.0f} кликов без заказов при ожидаемых ~{expected:.0f} — режем ставку сильнее.")
        advice.append("Много кликов без заказов: если после снижения ставки ситуация не изменится — уберите товар из рекламы.")
    gentle = (organic or in_test) and target < current
    step_down = GENTLE_STEP_DOWN if gentle else MAX_STEP
    recommended = min(max(target, current * (1 - step_down)), current * (1 + MAX_STEP))
    if abs(recommended - target) > 0.01:
        why = "товар на тесте" if in_test else "бережём органическую позицию"
        notes.append(f"Шаг ограничен {f'−10%: {why}' if gentle else '±25%'}, дальше — по новым данным.")
    min_bid = item.get("min_bid") or 0.0
    if recommended < min_bid:
        recommended = min_bid
        notes.append(f"Не ниже минимальной ставки Ozon {_rub(min_bid, 2)}.")
    recommended = round(recommended, 2)
    competitive = item.get("competitive_bid")
    if competitive and competitive > breakeven * 1.5:
        times = f"{competitive / breakeven:.1f}".replace(".", ",")
        notes.append(f"Ozon советует {_rub(competitive, 2)} — в {times} раза выше безубыточной: на такой ставке каждый заказ в минус.")
        advice.append("Рекламой этот товар выгодно не разогнать до топа. Рост — через цену, карточку и отзывы, а не через ставку.")
    result.update(breakeven_bid=round(breakeven, 2), target_bid=round(target, 2))
    if round(abs(recommended - current), 2) < round(max(MIN_CHANGE_RUB, current * 0.03), 2):
        steps.append(_step("Ставка", "ok", " ".join(notes) + f" Текущая {_rub(current, 2)} уже близка к цели."))
        return finish("keep", f"Ставка {_rub(current, 2)} близка к цели {_rub(target, 2)}.")
    action = "bid_down" if recommended < current else "bid_up"
    steps.append(_step("Ставка", "warn" if action == "bid_down" else "ok", " ".join(notes) + f" Решение: {_rub(current, 2)} → {_rub(recommended, 2)}."))
    if action == "bid_up":
        advice.append(f"Клик приносит больше, чем стоит: поднимаем ставку до {_rub(recommended, 2)}, чтобы получать больше показов.")
        reason = f"Окупается с запасом: цель {_rub(target, 2)}, поднимаем до {_rub(recommended, 2)}."
    else:
        reason = f"Переплачиваем за клик: цель {_rub(target, 2)}, снижаем до {_rub(recommended, 2)}."
    result["recommended_bid"] = recommended
    return finish(action, reason)


def recommend(item: Dict[str, Any], prior_cr: float, settings: Dict[str, Any]) -> Dict[str, Any]:
    """Решение по товару без сравнения с кампанией (для простых вызовов и тестов)."""
    item = {"views": 0.0, "to_cart": 0.0, "revenue": 0.0, "spend": 0.0, "price": None, **item}
    return analyze_item(item, prior_cr, {}, settings)


def season_outlook(root: Path, today: date) -> Optional[Dict[str, Any]]:
    """Сезон по истории продаж магазина (кэш заказов из раздела «Поставки»): следующие 30 дней год назад к предыдущим 30."""
    cache = root / "stocks reports" / ".orders_cache"
    if not cache.exists():
        return None
    daily: Dict[str, int] = {}
    for path in cache.glob("*.json"):
        try:
            for row in json.loads(path.read_text(encoding="utf-8")):
                daily[row["date"]] = daily.get(row["date"], 0) + int(row.get("qty") or 1)
        except (OSError, ValueError, KeyError):
            continue
    base = today - timedelta(days=365)
    units = lambda start, end: sum(q for d, q in daily.items() if start.isoformat() <= d < end.isoformat())
    before, after = units(base - timedelta(days=30), base), units(base, base + timedelta(days=30))
    if before < 10:
        return None
    return {"index": after / before, "before": before, "after": after}


def campaign_advice(campaign: Dict[str, Any], items: List[Dict[str, Any]], spend_week: Optional[float], week_start: Optional[date],
                    today: date, season: Optional[Dict[str, Any]], settings: Dict[str, Any]) -> List[Dict[str, str]]:
    """Рекомендации по кампании целиком: окупаемость, бюджет, сезон, структура."""
    share = float(settings["profit_share"])
    notes: List[Dict[str, str]] = []
    gross = sum(i["orders_total"] * i["profit"] for i in items if i.get("profit") is not None)
    spend = sum(i["spend"] for i in items)
    spent_share = _ratio(spend, gross)
    if spent_share is None:
        notes.append(_step("Окупаемость", "info", f"Расход {_rub(spend)}, заказов с рекламы пока нет — оценивать рано."))
    else:
        status = "ok" if spent_share <= share else ("warn" if spent_share <= 1 else "bad")
        notes.append(_step("Окупаемость", status,
            f"Реклама потратила {_rub(spend)} — {_pct(spent_share, 0)} прибыли с рекламных заказов ({_rub(gross)}), ваш лимит {share:.0%}. "
            + ("В пределах лимита." if status == "ok" else "Выше лимита — рекомендации ниже возвращают расходы в рамки." if status == "warn"
               else "Реклама работает в убыток: тратит больше, чем приносит.")))

    weekly = _num(campaign.get("weeklyBudget")) / MICRO
    if weekly and spend_week is not None and week_start is not None:
        days = (today - week_start).days + 1
        projected = spend_week / days * 7
        text = f"С {week_start:%d.%m} потрачено {_rub(spend_week)} из {_rub(weekly)} недельного бюджета; в таком темпе за неделю — {_rub(projected)}."
        if projected >= weekly * 0.95 and spent_share is not None and spent_share <= share:
            notes.append(_step("Бюджет", "warn", text + " Бюджет закончится раньше конца недели, а реклама окупается — "
                                                        f"можно поднять бюджет на 20% (до {_rub(weekly * 1.2)})."))
        elif projected >= weekly * 0.95:
            notes.append(_step("Бюджет", "warn", text + " Бюджет кончается до конца недели, но реклама не укладывается в лимит — "
                                                        "не увеличивайте его, пока ставки не придут к цели."))
        else:
            notes.append(_step("Бюджет", "ok", text + " Бюджет не ограничивает показы — ограничивают ставки."))

    if season:
        idx = season["index"]
        idx_text = f"{idx:.1f}".replace(".", ",")
        if idx >= SEASON_RISE:
            notes.append(_step("Сезон", "warn", f"Год назад в следующие 30 дней продажи магазина выросли в {idx_text} раза ({season['before']} → {season['after']} шт.). "
                                                 "Маркетологи поднимают бюджет и ставки за 1–2 недели до пика, чтобы к нему товар уже стоял высоко, "
                                                 "и заранее обеспечивают остатки (раздел «Поставки»)."))
        elif idx <= SEASON_FALL:
            notes.append(_step("Сезон", "info", f"Год назад в следующие 30 дней продажи упали до {idx:.0%} ({season['before']} → {season['after']} шт.). "
                                                 "В спад расходы на рекламу обычно снижают."))
        else:
            notes.append(_step("Сезон", "ok", f"По прошлому году спрос в ближайшие 30 дней ровный ({idx_text}×) — особых действий не нужно."))

    tests = [i for i in items if i.get("stage") == "test"]
    mature = [i for i in items if i.get("stage") == "mature"]
    if tests and mature:
        notes.append(_step("Структура", "info", f"В кампании вместе {len(tests)} товаров на тесте и {len(mature)} зрелых. "
                                                 "Маркетологи держат новинки в отдельной кампании с отдельным бюджетом — "
                                                 "иначе зрелые товары забирают бюджет, а новинки не набирают статистику."))
    return notes


def _week_start(today: date, start_weekday: str) -> date:
    weekday = WEEKDAYS.index(start_weekday) if start_weekday in WEEKDAYS else 0
    return today - timedelta(days=(today.weekday() - weekday) % 7)


# ---------------------------------------------------------------- план и применение

def build_ads_plan(root: Path, settings: Dict[str, Any], today: Optional[date] = None,
                   campaign_ids: Optional[List[str]] = None) -> Dict[str, Any]:
    """Рекомендации по работающим кампаниям; campaign_ids — только по выбранным."""
    from recommended_prices import get_sku_to_offer_id_mapping

    today = today or date.today()
    client = PerfClient()
    campaigns = client.running_cpc_campaigns()
    if campaign_ids is not None:
        campaigns = [c for c in campaigns if str(c["id"]) in set(campaign_ids)]
    if not campaigns:
        return {"ok": True, "message": "Кампания сейчас не работает — подбирать ставки не для чего.", "campaigns": [], "items": [],
                "period": [today.isoformat(), today.isoformat()], "settings": settings, "created": datetime.now().isoformat(timespec="seconds")}
    economics = load_unit_economics(root)
    date_from = today - timedelta(days=int(settings["lookback_days"]))
    campaign_ids = [str(c["id"]) for c in campaigns]
    stats = client.sku_statistics(campaign_ids, date_from, today)
    daily = client.daily_spend(campaign_ids, today - timedelta(days=7), today)

    products_by_campaign = {str(c["id"]): client.campaign_products(str(c["id"])) for c in campaigns}
    all_skus = sorted({str(p["sku"]) for products in products_by_campaign.values() for p in products})
    offers = {str(sku): offer for sku, offer in get_sku_to_offer_id_mapping([int(s) for s in all_skus]).items()}
    stock = load_available_stock(sorted(set(offers.values())))
    min_bids = client.min_bids(all_skus)
    season = season_outlook(root, today)

    include_models = bool(settings["include_model_orders"])
    campaign_rows, items = [], []
    for campaign in campaigns:
        campaign_id = str(campaign["id"])
        strategy = campaign.get("productAutopilotStrategy") or ""
        rows = {str(r.get("sku")): r for r in stats.get(campaign_id, [])}
        products = products_by_campaign[campaign_id]
        competitive = client.competitive_bids(campaign_id, [str(p["sku"]) for p in products]) if products else {}

        campaign_items = []
        for product in products:
            sku = str(product["sku"])
            row = rows.get(sku, {})
            offer = offers.get(sku)
            econ = economics.get(offer or "", {})
            direct, models = _num(row.get("orders")), _num(row.get("models"))
            added = row.get("createdAt")
            try:
                days_in = (today - datetime.strptime(added, "%Y-%m-%d").date()).days if added else None
            except ValueError:
                days_in = None
            campaign_items.append({
                "campaign_id": campaign_id,
                "campaign_title": campaign.get("title") or f"Кампания {campaign_id}",
                "strategy": strategy,
                "sku": sku,
                "offer_id": offer,
                "title": product.get("title") or row.get("title") or "",
                "current_bid": _num(product.get("bid")) / MICRO,
                "competitive_bid": competitive.get(sku),
                "min_bid": min_bids.get(sku),
                "views": _num(row.get("views")),
                "clicks": _num(row.get("clicks")),
                "to_cart": _num(row.get("toCart")),
                "spend": _num(row.get("moneySpent")),
                "orders": direct,
                "model_orders": models,
                "orders_total": direct + (models if include_models else 0),
                "revenue": _num(row.get("ordersMoney")) + (_num(row.get("modelsMoney")) if include_models else 0),
                "product_gmv": _num(row.get("product_gmv")),
                "price": _num(row.get("price")) or econ.get("price"),
                "profit": econ.get("profit"),
                "cost": econ.get("cost"),
                "stock": stock.get(offer) if offer else None,
                "days_in_campaign": days_in,
            })
        clicks = sum(i["clicks"] for i in campaign_items)
        orders = sum(i["orders_total"] for i in campaign_items)
        prior_cr = orders / clicks if clicks >= 100 and orders > 0 else None
        campaign_rows.append({"campaign": campaign, "items": campaign_items, "prior_cr": prior_cr, "clicks": clicks, "orders": orders})

    # Если в кампании мало данных — берём среднюю конверсию по всем кампаниям.
    total_clicks = sum(c["clicks"] for c in campaign_rows)
    total_orders = sum(c["orders"] for c in campaign_rows)
    pooled_cr = total_orders / total_clicks if total_clicks and total_orders else FALLBACK_CR

    campaigns_out = []
    for entry in campaign_rows:
        prior = entry["prior_cr"] or pooled_cr
        bench = campaign_benchmarks(entry["items"])
        for item in entry["items"]:
            item.update(analyze_item(item, prior, bench, settings))
            item["selected"] = item["action"] in ("bid_down", "bid_up", "remove")
            items.append(item)
        campaign = entry["campaign"]
        campaign_id = str(campaign["id"])
        week_start = _week_start(today, campaign.get("startWeekDay") or "MONDAY")
        spend_week = sum(spend for day, spend in daily.get(campaign_id, {}).items() if day >= week_start.isoformat())
        spend = sum(i["spend"] for i in entry["items"])
        gross = sum(i["orders_total"] * i["profit"] for i in entry["items"] if i["profit"] is not None)
        campaigns_out.append({
            "id": campaign_id,
            "title": campaign.get("title") or f"Кампания {campaign_id}",
            "strategy": STRATEGY_NAMES.get(campaign.get("productAutopilotStrategy") or "", campaign.get("productAutopilotStrategy") or ""),
            "weekly_budget": _num(campaign.get("weeklyBudget")) / MICRO,
            "daily_budget": _num(campaign.get("dailyBudget")) / MICRO,
            "spend": round(spend, 2),
            "clicks": entry["clicks"],
            "orders": entry["orders"],
            "conversion": prior,
            "benchmarks": bench,
            "profit_before_ads": round(gross, 2),
            "profit_after_ads": round(gross - spend, 2),
            "profit_share_spent": round(spend / gross, 3) if gross > 0 else None,
            "advice": campaign_advice(campaign, entry["items"], spend_week, week_start, today, season, settings),
        })
    changes = sum(1 for i in items if i["action"] in ("bid_down", "bid_up", "remove"))
    message = f"Кампаний: {len(campaigns_out)}, товаров: {len(items)}, предложений: {changes}. Период {date_from:%d.%m}–{today:%d.%m}."
    return {
        "ok": True,
        "message": message,
        "created": datetime.now().isoformat(timespec="seconds"),
        "period": [date_from.isoformat(), today.isoformat()],
        "settings": settings,
        "campaigns": campaigns_out,
        "items": items,
    }


def execute_ads_plan(root: Path, items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Применяет отмеченные изменения и перепроверяет по факту: ставка считается
    применённой, только если Ozon после отправки показывает новое значение.
    """
    client = PerfClient()
    selected = [i for i in items if i.get("selected") and i["action"] in ("bid_down", "bid_up", "remove")]
    by_campaign: Dict[str, List[Dict[str, Any]]] = {}
    for item in selected:
        by_campaign.setdefault(item["campaign_id"], []).append(item)

    confirmed: List[Dict[str, Any]] = []
    failed: List[Dict[str, Any]] = []
    for campaign_id, campaign_items in by_campaign.items():
        bids = {i["sku"]: i["recommended_bid"] for i in campaign_items if i["action"] in ("bid_down", "bid_up")}
        removals = [i["sku"] for i in campaign_items if i["action"] == "remove"]
        errors: Dict[str, str] = {}
        try:
            if bids:
                client.update_bids(campaign_id, bids)
        except Exception as exc:
            errors.update({sku: str(exc) for sku in bids})
        try:
            if removals:
                client.remove_products(campaign_id, removals)
        except Exception as exc:
            errors.update({sku: str(exc) for sku in removals})
        actual = {str(p["sku"]): _num(p.get("bid")) / MICRO for p in client.campaign_products(campaign_id)}
        for item in campaign_items:
            sku = item["sku"]
            if item["action"] == "remove":
                applied = sku not in actual
            else:
                applied = sku in actual and abs(actual[sku] - item["recommended_bid"]) < 0.015
            entry = {**_short(item), "actual_bid": actual.get(sku)}
            if applied:
                confirmed.append(entry)
            else:
                failed.append({**entry, "error": errors.get(sku) or "Ozon принял запрос, но изменение не видно в кампании"})
    _log_changes(root, confirmed)
    message = f"Применено и подтверждено Ozon: {len(confirmed)} из {len(selected)}."
    if failed:
        message += f" Не применилось: {len(failed)}."
    return {"ok": not failed, "message": message, "confirmed": confirmed, "failed": failed}


def _short(item: Dict[str, Any]) -> Dict[str, Any]:
    keys = ("campaign_id", "campaign_title", "sku", "offer_id", "title", "action", "current_bid", "recommended_bid", "reason")
    return {key: item.get(key) for key in keys}


def _log_changes(root: Path, entries: List[Dict[str, Any]]) -> None:
    if not entries:
        return
    folder = root / ADS_REPORTS_DIR
    folder.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().isoformat(timespec="seconds")
    with (folder / CHANGES_LOG).open("a", encoding="utf-8") as log:
        for entry in entries:
            log.write(json.dumps({"time": stamp, **entry}, ensure_ascii=False) + "\n")


def main() -> None:
    """Консольный просмотр рекомендаций (без применения)."""
    settings = load_settings(repo_root)
    plan = build_ads_plan(repo_root, settings)
    print(plan["message"])
    for campaign in plan["campaigns"]:
        print(f"\n{campaign['title']} — {campaign['strategy']}: расход {campaign['spend']:.0f} ₽, заказов {campaign['orders']:.0f}, "
              f"прибыль до рекламы {campaign['profit_before_ads']:.0f} ₽, после {campaign['profit_after_ads']:.0f} ₽")
        for item in [i for i in plan["items"] if i["campaign_id"] == campaign["id"]]:
            change = f"{item['current_bid']:.2f} -> {item['recommended_bid']:.2f}" if item["action"] in ("bid_down", "bid_up") else item["action"]
            print(f"  {item['offer_id'] or item['sku']}: {change} | {item['reason']}")


if __name__ == "__main__":
    main()
