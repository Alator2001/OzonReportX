"""
Разбор ответов выдачи Ozon: чистые функции без сети и браузера (их проверяют тесты на сохранённых JSON).
Пути к данным — в serp_paths.py.
"""
import json
from typing import Any, Dict, Iterable, List, Optional

try:
    from scripts import serp_paths as paths
except ModuleNotFoundError:
    import serp_paths as paths


class SerpBlocked(Exception):
    """Ozon показал капчу или заблокировал запросы — прогон нужно остановить, а не повторять."""


class SerpFormatError(Exception):
    """Ответ не похож на выдачу: Ozon поменял формат. Замер не пишем."""


def check_response(status: int, content_type: str, text: str) -> Dict[str, Any]:
    """Ответ JSON-эндпоинта → словарь. Капча/блокировка → SerpBlocked, не-JSON → SerpFormatError."""
    content_type = (content_type or "").lower()
    if status in (403, 429) or paths.CAPTCHA_TITLE_RE.search(text[:5000] if "html" in content_type else ""):
        raise SerpBlocked(f"Ozon ответил HTTP {status} ({'капча' if 'html' in content_type else content_type or 'без типа'}).")
    if status >= 400:
        raise SerpFormatError(f"HTTP {status} от выдачи Ozon.")
    if "json" not in content_type:
        raise SerpFormatError(f"Ожидали JSON, пришёл {content_type or 'ответ без типа'}.")
    try:
        return json.loads(text)
    except ValueError as e:
        raise SerpFormatError(f"Ответ выдачи не разобран как JSON: {e}") from None


def _widget(data: Dict[str, Any], prefixes: Iterable[str]) -> Optional[Dict[str, Any]]:
    states = data.get("widgetStates") or {}
    for key in sorted(states):
        if key.startswith(tuple(prefixes)):
            value = states[key]
            return json.loads(value) if isinstance(value, str) else value
    return None


def _strings(value: Any) -> Iterable[str]:
    if isinstance(value, str):
        yield value
    elif isinstance(value, dict):
        for v in value.values():
            yield from _strings(v)
    elif isinstance(value, list):
        for v in value:
            yield from _strings(v)


def is_ad(item: Dict[str, Any]) -> bool:
    return any(text.strip() in paths.AD_LABELS for text in _strings(item))


def location(data: Dict[str, Any]) -> Dict[str, Optional[str]]:
    current: Any = data
    for key in paths.LOCATION_PATH:
        current = (current or {}).get(key) if isinstance(current, dict) else None
    current = current if isinstance(current, dict) else {}
    return {"fias": current.get(paths.LOCATION_FIAS_KEY), "city": current.get(paths.LOCATION_CITY_KEY)}


def parse_page(data: Dict[str, Any]) -> Dict[str, Any]:
    """Порция выдачи → {items: [{sku, is_ad}], next_page, location}. Нет виджета товаров — SerpFormatError."""
    tiles = _widget(data, paths.TILE_WIDGET_PREFIXES)
    if tiles is None:
        raise SerpFormatError("В ответе нет виджета с товарами выдачи — возможно, Ozon поменял формат страницы.")
    items = []
    for item in tiles.get(paths.ITEMS_KEY) or []:
        sku = item.get(paths.SKU_KEY)
        if sku in (None, ""):
            continue
        items.append({"sku": int(sku), "is_ad": is_ad(item)})
    paginator = _widget(data, paths.PAGINATOR_WIDGET_PREFIXES) or {}
    return {"items": items, "next_page": paginator.get(paths.NEXT_PAGE_KEY) or None, "location": location(data)}


def rank(pages: List[List[Dict[str, Any]]]) -> Dict[int, Dict[str, Any]]:
    """
    Порции выдачи по порядку → {sku: позиция}. Позиция — номер в общей выдаче (с рекламой), органическая — без
    рекламных товаров выше и самого товара, если он рекламный (тогда organic_position = None). Учитывается первое появление.
    """
    result: Dict[int, Dict[str, Any]] = {}
    position = organic = 0
    for page_no, items in enumerate(pages, start=1):
        for item in items:
            position += 1
            if not item["is_ad"]:
                organic += 1
            if item["sku"] in result:
                continue
            result[item["sku"]] = {"position": position, "organic_position": None if item["is_ad"] else organic,
                                   "page": page_no, "is_ad": int(item["is_ad"])}
    return result


def positions_for(skus: Iterable[int], ranked: Dict[int, Dict[str, Any]], depth: int) -> List[Dict[str, Any]]:
    """Строки для записи по нашим SKU: найден — позиция; не найден в просмотренной глубине — position NULL, found 0."""
    rows = []
    for sku in skus:
        hit = ranked.get(int(sku))
        if hit:
            rows.append({"sku": int(sku), **hit, "found": 1, "depth": depth})
        else:
            rows.append({"sku": int(sku), "position": None, "organic_position": None, "page": None, "is_ad": None,
                         "found": 0, "depth": depth})
    return rows
