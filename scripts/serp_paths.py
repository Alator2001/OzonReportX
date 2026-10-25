"""
Всё, что парсер выдачи знает об устройстве ozon.ru, — в одном месте. Когда Ozon меняет сайт, правится только этот файл.

Проверено во встроенном браузере 23.10.2026 (анонимная сессия, запрос «доска разделочная деревянная»):
- Выдача отдаётся JSON-ом: GET /api/entrypoint-api.bx/page/json/v2?url=<путь страницы>.
  Поиск /search/?text=… Ozon часто переадресует в категорию (/category/…/?text=…) — это та же выдача.
- Товары: widgetStates["tileGridDesktop-…"] — строка JSON, в ней items[] с полем sku. Первая порция — 8 товаров,
  дальше порции подгружаются по ссылке nextPage из widgetStates["infiniteVirtualPaginator-…"].
- Регион: верхний ключ location.current — {city, uid (код ФИАС), name, …}. Москва: 0c5b2444-70a0-4932-980c-b4dc0d3f02b5.
- Капча: HTTP 403, text/html, <title>Antibot Captcha</title>. Появилась на 5-м запросе за минуту.
- Реклама: явной метки не нашли — ни текста «Реклама», ни erid; поля advert/advertLite есть у ВСЕХ товаров (это
  токены трекинга, а не признак рекламы). Пока считаем рекламой товар с текстовой меткой из AD_LABELS; сырые ответы
  сохраняются (serp_collect.RAW_DIR), чтобы уточнить правило, когда метка встретится.
"""
import re
from urllib.parse import quote

SITE = "https://www.ozon.ru"
JSON_ENDPOINT = "/api/entrypoint-api.bx/page/json/v2?url="

# Префиксы ключей widgetStates. Ключ вида «tileGridDesktop-3669724-default-1»: id и номер порции меняются.
TILE_WIDGET_PREFIXES = ("tileGridDesktop-", "searchResultsV2-")
PAGINATOR_WIDGET_PREFIXES = ("infiniteVirtualPaginator-", "megaPaginator-")
ADDRESS_WIDGET_PREFIX = "addressBookBarWeb-"

ITEMS_KEY = "items"
SKU_KEY = "sku"
NEXT_PAGE_KEY = "nextPage"
LOCATION_PATH = ("location", "current")
LOCATION_FIAS_KEY = "uid"
LOCATION_CITY_KEY = "city"

# Текстовые метки рекламы в карточке товара (ищутся в любом строковом поле товара целиком).
AD_LABELS = ("Реклама",)

CAPTCHA_TITLE_RE = re.compile(r"<title>\s*Antibot", re.IGNORECASE)


def search_path(query: str) -> str:
    return f"/search/?text={quote(query)}&from_global=true"
