"""
Сбор поисковой аналитики из Seller API для раздела «Поиск и ключевые слова».

| Метод | Что берём | Без подписки (проверено 23.10.2026) |
|---|---|---|
| /v3/product/list + /v3/product/info/list | SKU, артикул, название | всё |
| /v1/analytics/product-queries | по SKU за день: искали, увидели, позиция, конверсия, продажи | только unique_search_users |
| /v1/analytics/product-queries/details | по SKU и запросу за день (до 15 запросов на SKU) | текст запроса, query_index, unique_search_users |
| /v1/analytics/data | по SKU и дню: заказано штук и на сумму | ordered_units, revenue; 1 запрос в минуту, 50 в сутки |

Период Ozon — полуинтервал: день D = [D 00:00 UTC, D+1 00:00 UTC). Поля, которые Ozon отдаёт
только с Premium, сохраняются как есть (NULL без подписки) — при появлении подписки они заполнятся сами.

Ошибка доступа (403) — не авария: источник помечается как недоступный, остальные собираются.
Любая другая ошибка страницы — исключение: неполный день не сохраняется (запись дня — одна транзакция).
"""
import os
import time
from datetime import date, datetime, timedelta, timezone
from typing import Any, Callable, Dict, List, Optional

import requests
from dotenv import load_dotenv
try:
    from scripts.app_paths import ENV_FILE
except ModuleNotFoundError:
    from app_paths import ENV_FILE

try:
    from scripts.report_http import create_session
except ModuleNotFoundError:
    from report_http import create_session

BASE_URL = "https://api-seller.ozon.ru"
TIMEOUT = 60
DETAILS_PAGE_SIZE = 100
DETAILS_LIMIT_BY_SKU = 15
SKU_CHUNK = 1000
ANALYTICS_DATA_LIMIT = 1000
ANALYTICS_DATA_INTERVAL = 61.0      # Ozon: не чаще 1 раза в минуту
ANALYTICS_DATA_DAILY_QUOTA = 45     # Ozon без Premium Plus: 50 в сутки; оставляем запас для других программ


class AccessDenied(Exception):
    """Ozon отказал в доступе к методу (обычно нужна подписка)."""


def _headers() -> Dict[str, str]:
    # Ключи читаются при каждом запросе: модуль импортирует приложение, и без ключей он не должен падать при импорте.
    load_dotenv(ENV_FILE)
    client_id, api_key = os.getenv("OZON_CLIENT_ID"), os.getenv("OZON_API_KEY")
    if not client_id or not api_key:
        raise RuntimeError("Не заданы OZON_CLIENT_ID и OZON_API_KEY. Укажите их в «Настройках».")
    return {"Client-Id": client_id, "Api-Key": api_key, "Content-Type": "application/json"}


def _ozon_message(response: requests.Response) -> str:
    try:
        return str(response.json().get("message") or "")
    except (ValueError, AttributeError):
        return response.text[:300]


def post(session: requests.Session, path: str, body: Dict[str, Any]) -> Dict[str, Any]:
    response = session.post(BASE_URL + path, headers=_headers(), json=body, timeout=TIMEOUT)
    if response.status_code == 403:
        raise AccessDenied(f"{path}: Ozon отказал в доступе — {_ozon_message(response)}")
    if response.status_code >= 400:
        raise RuntimeError(f"{path}: HTTP {response.status_code} — {_ozon_message(response)}")
    return response.json()


def day_bounds(day: date) -> Dict[str, str]:
    start = datetime.combine(day, datetime.min.time())
    return {"date_from": start.isoformat() + "Z", "date_to": (start + timedelta(days=1)).isoformat() + "Z"}


def _int(value: Any) -> Optional[int]:
    return None if value is None else int(value)


def _float(value: Any) -> Optional[float]:
    return None if value is None else float(value)


# --- товары -----------------------------------------------------------------------------------

def load_products(session: requests.Session) -> List[Dict[str, Any]]:
    """Все товары магазина с SKU: [{sku, offer_id, name}]. Пагинация с защитой от зацикливания."""
    product_ids: List[str] = []
    seen, last_id = set(), ""
    while True:
        data = post(session, "/v3/product/list", {"filter": {"visibility": "ALL"}, "last_id": last_id, "limit": 1000})
        result = data.get("result") or {}
        items = result.get("items") or []
        product_ids += [str(i["product_id"]) for i in items]
        last_id = str(result.get("last_id") or "")
        if not items or not last_id:
            break
        if last_id in seen:
            raise RuntimeError("/v3/product/list: Ozon вернул тот же last_id повторно — список товаров не загружен.")
        seen.add(last_id)
    products = []
    for i in range(0, len(product_ids), 1000):
        data = post(session, "/v3/product/info/list", {"product_id": product_ids[i:i + 1000]})
        for item in data.get("items") or []:
            if item.get("sku"):
                products.append({"sku": int(item["sku"]), "product_id": int(item["id"]) if item.get("id") else None,
                                 "offer_id": item.get("offer_id"), "name": item.get("name")})
    return products


# --- карточки ---------------------------------------------------------------------------------

# Атрибуты Ozon с особым смыслом (проверено на /v4/product/info/attributes 23.10.2026).
NAME_ATTRIBUTE = 4180            # «Название» — то же, что name
DESCRIPTION_ATTRIBUTE = 4191     # «Аннотация» — описание карточки
RICH_CONTENT_ATTRIBUTE = 11254   # Rich-контент, JSON
SERVICE_ATTRIBUTES = {9024}      # «Артикул» — не текст для покупателя


def _attribute_values(attribute: Dict[str, Any]) -> List[str]:
    return [str(v.get("value")).strip() for v in attribute.get("values") or [] if str(v.get("value") or "").strip()]


def parse_card(item: Dict[str, Any]) -> Dict[str, Any]:
    """Карточка из /v4/product/info/attributes: название, описание, прочие характеристики {id: [значения]}, фото."""
    description, rich_content, attributes = "", "", {}
    for attribute in item.get("attributes") or []:
        attribute_id, values = int(attribute.get("id") or 0), _attribute_values(attribute)
        if attribute_id == DESCRIPTION_ATTRIBUTE:
            description = "\n".join(values)
        elif attribute_id == RICH_CONTENT_ATTRIBUTE:
            rich_content = "\n".join(values)
        elif attribute_id not in SERVICE_ATTRIBUTES and attribute_id != NAME_ATTRIBUTE and values:
            attributes[str(attribute_id)] = values
    images = [item["primary_image"]] if item.get("primary_image") else []
    images += [url for url in item.get("images") or [] if isinstance(url, str) and url not in images]
    return {"sku": int(item["sku"]), "product_id": _int(item.get("id")), "name": item.get("name") or "",
            "description": description, "attributes": attributes, "rich_content": rich_content, "images": images}


def fetch_cards(session: requests.Session, skus: List[int]) -> List[Dict[str, Any]]:
    cards = []
    for i in range(0, len(skus), SKU_CHUNK):
        seen, last_id = set(), ""
        while True:
            data = post(session, "/v4/product/info/attributes", {
                "filter": {"sku": [str(s) for s in skus[i:i + SKU_CHUNK]], "visibility": "ALL"},
                "limit": 1000, "last_id": last_id,
            })
            items = data.get("result") or []
            cards += [parse_card(item) for item in items if item.get("sku")]
            last_id = str(data.get("last_id") or "")
            if not items or not last_id:
                break
            if last_id in seen:
                raise RuntimeError("/v4/product/info/attributes: Ozon вернул тот же last_id повторно — карточки не загружены.")
            seen.add(last_id)
    return cards


# --- product-queries --------------------------------------------------------------------------

def _paged(session: requests.Session, path: str, body: Dict[str, Any], key: str) -> List[Dict[str, Any]]:
    rows, page = [], 0
    while True:
        data = post(session, path, {**body, "page": page})
        rows += data.get(key) or []
        page += 1
        if page >= int(data.get("page_count") or 0):
            return rows


def parse_product_queries(items: List[Dict[str, Any]], day: str) -> List[Dict[str, Any]]:
    return [{
        "date": day, "sku": int(i["sku"]),
        "searched_users": _int(i.get("unique_search_users")), "viewed_users": _int(i.get("unique_view_users")),
        "view_conversion": _float(i.get("view_conversion")), "position": _float(i.get("position")),
        "gmv": _float(i.get("gmv")),
    } for i in items if i.get("sku")]


def parse_query_details(queries: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    rows = []
    for q in queries:
        text = " ".join(str(q.get("query") or "").split())
        if not q.get("sku") or not text:
            continue
        rows.append({
            "sku": int(q["sku"]), "query": text, "query_index": _int(q.get("query_index")),
            "searched_users": _int(q.get("unique_search_users")), "viewed_users": _int(q.get("unique_view_users")),
            "view_conversion": _float(q.get("view_conversion")), "position": _float(q.get("position")),
            "orders": _int(q.get("order_count")), "gmv": _float(q.get("gmv")),
        })
    return rows


def fetch_product_queries_day(session: requests.Session, skus: List[int], day: date) -> Dict[str, List[Dict[str, Any]]]:
    """Сводка по SKU и детализация по запросам за один день."""
    bounds, iso = day_bounds(day), day.isoformat()
    sku_rows, query_rows = [], []
    for i in range(0, len(skus), SKU_CHUNK):
        chunk = [str(s) for s in skus[i:i + SKU_CHUNK]]
        items = _paged(session, "/v1/analytics/product-queries", {**bounds, "skus": chunk, "page_size": 1000}, "items")
        parsed = parse_product_queries(items, iso)
        sku_rows += parsed
        found = [str(r["sku"]) for r in parsed if r["searched_users"]]
        if not found:
            continue
        queries = _paged(session, "/v1/analytics/product-queries/details",
                         {**bounds, "skus": found, "page_size": DETAILS_PAGE_SIZE, "limit_by_sku": DETAILS_LIMIT_BY_SKU}, "queries")
        query_rows += parse_query_details(queries)
    return {"sku_daily": sku_rows, "query_stats": query_rows}


# --- analytics/data ---------------------------------------------------------------------------

def parse_analytics_data(data: Dict[str, Any], metrics: List[str]) -> List[Dict[str, Any]]:
    """Строки ответа с dimension [sku, day] → [{date, sku, <метрика>: число}]."""
    rows = []
    for item in (data.get("result") or {}).get("data") or []:
        dims = item.get("dimensions") or []
        values = item.get("metrics") or []
        if len(dims) < 2 or not dims[0].get("id") or not dims[1].get("id"):
            continue
        row: Dict[str, Any] = {"sku": int(dims[0]["id"]), "date": str(dims[1]["id"])[:10]}
        # Без подписки Ozon молча выкидывает недоступные метрики: значений может быть меньше, чем запрошено.
        if len(values) != len(metrics):
            raise RuntimeError(f"/v1/analytics/data: ожидали {len(metrics)} метрик, Ozon вернул {len(values)}.")
        row.update({m: _float(v) for m, v in zip(metrics, values)})
        rows.append(row)
    return rows


def fetch_sales_by_day(session: requests.Session, date_from: date, date_to: date,
                       before_request: Callable[[], None] = lambda: None) -> List[Dict[str, Any]]:
    """Заказано штук и на сумму по SKU и дням за [date_from, date_to] включительно."""
    metrics = ["ordered_units", "revenue"]
    rows, offset = [], 0
    while True:
        before_request()
        data = post(session, "/v1/analytics/data", {
            "date_from": date_from.isoformat(), "date_to": date_to.isoformat(),
            "dimension": ["sku", "day"], "metrics": metrics, "limit": ANALYTICS_DATA_LIMIT, "offset": offset,
        })
        rows += parse_analytics_data(data, metrics)
        if len((data.get("result") or {}).get("data") or []) < ANALYTICS_DATA_LIMIT:
            return rows
        offset += ANALYTICS_DATA_LIMIT


def utc_today() -> date:
    return datetime.now(timezone.utc).date()


def sleep_until(moment: float) -> None:
    pause = moment - time.time()
    if pause > 0:
        print(f"⏳ Пауза {pause:.0f} с: /v1/analytics/data можно вызывать не чаще раза в минуту.", flush=True)
        time.sleep(pause)

