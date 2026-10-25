# 03. Работа с API Ozon

Программа работает с двумя API Ozon:

| API | Хост | Авторизация | Ключи в `.env` |
|---|---|---|---|
| **Seller API** — товары, цены, акции, заказы, финансы, аналитика | `https://api-seller.ozon.ru` | заголовки `Client-Id` и `Api-Key` | `OZON_CLIENT_ID`, `OZON_API_KEY` |
| **Performance API** — реклама | `https://api-performance.ozon.ru` | `Authorization: Bearer <token>`; токен по `client_credentials` (`/api/client/token`) | `OZON_PERF_CLIENT_ID`, `OZON_PERF_API_KEY` (client secret) |

Документация Ozon: [docs.ozon.ru/api/seller](https://docs.ozon.ru/api/seller),
новости об изменениях — Telegram-канал [@OzonSellerAPI](https://t.me/s/OzonSellerAPI)
и раздел новостей [dev.ozon.ru](https://dev.ozon.ru/news/).

## Где берутся заголовки

Единого клиента нет; модули исторически делают запросы по-разному.

| Модуль | Как отправляет | Заголовки |
|---|---|---|
| `Monthly_sales_report.py` | `create_session()` → `ReportSession` + `HTTPAdapter` с повторами | свой `HEADERS`, ключи читаются при импорте (без ключей — `RuntimeError`) |
| `balance_report.py` | `requests.post` | `_headers()` — ключи читаются при каждом запросе; без ключей — `RuntimeError` при запросе |
| `recommended_prices.py`, `price_management.py`, `update_prices.py` | `requests.post/get` напрямую | `recommended_prices.OZON_HEADERS` |
| `auto_actions.py` | `_post` — 3 попытки | `OZON_HEADERS` |
| `fbo_supply_plan.py` | `ozon_post` — 4 попытки (паузы 5/15/30 с) | `OZON_HEADERS` |
| `organic_analysis.py` | `_post` — без повторов | `OZON_HEADERS` |
| `ads_cpc.py` | `PerfClient` (Performance API) — до 9 попыток при 429 | Bearer-токен |
| `performance_api.py` | сессия месячного отчёта | Bearer-токен |

`OZON_HEADERS` — словарь, общий для всех модулей в процессе. После сохранения ключей в
«Настройках» приложение вызывает `recommended_prices.refresh_api_credentials()`, которая
обновляет этот словарь на месте — поэтому новые ключи работают без перезапуска для функций,
вызываемых в потоках. Дочерние процессы и так читают `.env` заново.

## Лимиты

Что известно о лимитах Seller API:

- **50 запросов в секунду на Client-Id** суммарно по всем методам (с мая 2025).
- **Отдельные лимиты у некоторых методов** поверх общего. Из объявлений Ozon: черновики поставок
  (2 запроса в минуту), операции с товарами (30 000 операций в минуту с 24 февраля 2026),
  остатки — не чаще раза в 30 секунд на пару товар–склад. Точные значения смотрите в описании метода.
- С 25 сентября 2026 Ozon возвращает заголовки `ratelimit-remaining` и `retry-after`.
  При превышении — HTTP 429 с телом `{"code": 8, "message": "You have reached request rate limit per second"}`.
- **Методы, которые выводят из работы, Ozon сначала «душит»:** `v3/posting/fbs/list` в октябре 2026
  отвечал 429 с `ratelimit-remaining: 0` на любой, даже одиночный запрос, а соседний
  `v3/posting/fbs/unfulfilled/list` уже отвечал `400 {"code": 9, "message": "obsolete method cannot be used"}`.
  **Постоянный 429 на один метод при свободном лимите остальных — признак, что метод выводится и нужна новая версия.**
  Проверить можно одним запросом и сравнить `ratelimit-remaining` с другим методом.

### `ReportSession` — темп и повторы для отчётов

`scripts/report_http.py`. Используется месячным отчётом (и всем, что берёт `report_http.create_session()`,
например загрузкой заказов для плана поставки и анализа).

- Касается только хоста `api-seller.ozon.ru`; прочие хосты идут как обычная `requests.Session`.
- Все запросы сессии проходят через один `threading.Lock` — строго по одному.
- Между запросами пауза не меньше `min_interval` (0,5 с).
- **При 429** ждёт `max(5·2^попытка, Retry-After)` секунд: 5, 10, 20, 40, 80. Максимум 5 повторов
  и не больше 300 секунд ожидания суммарно, затем `RuntimeError` с методом и текстом Ozon:
  «Ozon продолжает ограничивать запросы к /v4/posting/fbs/list (HTTP 429). Ответ Ozon: «…». Если ошибка
  повторяется, лимит этого метода, скорее всего, расходует другая программа…».
- После каждого 429 `min_interval` удваивается (от 1 до 5 с) — сессия замедляется до конца работы.
- Повторяется **тот же запрос с тем же телом** (важно для пагинации: страница не теряется).

`create_session()` (`scripts/report_http.py`; `Monthly_sales_report` импортирует её оттуда) монтирует два адаптера:
- для всех хостов — повторы при 429/500/502/503/504 (5 попыток, `backoff_factor=0.5`, учитывает `Retry-After`);
- для `https://api-seller.ozon.ru/` — повторы **только при 5xx**: 429 обрабатывает `ReportSession`,
  иначе адаптер и сессия умножали бы ожидания друг друга (тест `test_adapter_does_not_hide_429_from_pacing_layer`).

Для Performance API отдельно действует правило: одновременно может формироваться только один отчёт
статистики, поэтому `PerfClient.call` при 429 ждёт 5, 10, … 30 секунд (до ~3 минут суммарно).

## Полный список методов

Статус проверен запросами 19–22 октября 2026 (для методов записи — только что метод не заблокирован).
«—» — не проверялся.

### Заказы и финансы

| Метод | Модуль | Назначение | Статус |
|---|---|---|---|
| `POST /v4/posting/fbs/list` | `Monthly_sales_report.py` | Заказы FBS за период; курсор, `statuses` списком, `limit ≤ 100` | ✅ |
| `POST /v3/posting/fbo/list` | `Monthly_sales_report.py` | Заказы FBO за период; курсор | ✅ |
| `POST /v1/finance/accrual/by-day` | `Monthly_sales_report.py` | Начисления за день, пагинация `last_id`; заменил `v3/finance/transaction/list` | ✅ |
| `POST /v1/finance/balance` | `balance_report.py` | Баланс: выплаты, начисления, услуги; период ≤ 30 дней | ✅ |

### Товары, цены, остатки

| Метод | Модуль | Назначение | Статус |
|---|---|---|---|
| `POST /v3/product/list` | `recommended_prices.py` | Список товаров | ✅ |
| `POST /v3/product/info/list` | `recommended_prices.py` (в т. ч. `get_product_info_by_id`) | Карточки: SKU, `offer_id` по `product_id`/`sku` | ✅ |
| `POST /v5/product/info/prices` | `recommended_prices.py`, `auto_actions.py`, `organic_analysis.py`, `update_prices.py` | Цены, индекс цен, флаги автодобавления, `product_id → offer_id`. **С 23.10.2026 — не больше 100 товаров за запрос** (`limit` и фильтр `offer_id`): на большее Ozon отвечает HTTP 429 «Уменьшите батч до 100», похоже на ограничение частоты, но повтор не помогает (`PRICES_BATCH`) | ✅ |
| `POST /v1/product/import/prices` | `update_prices.py`, `auto_actions.py` | Запись минимальных цен и флагов автодобавления / учёта минимальной цены | — (запись) |
| `POST /v4/product/info/stocks` | `recommended_prices.py`, `fbo_supply_plan.py`, `ads_cpc.py` | Остатки по складам | ✅ |
| `POST /v1/analytics/stocks` | `fbo_supply_plan.py`, `recommended_prices.py` | Остатки и продажи FBO по кластерам | ✅ |
| `POST /v1/product/rating-by-sku` | `organic_analysis.py` | Контент-рейтинг карточки | ✅ |
| `POST /v1/product/action/timer/status` | `auto_actions.py` | Учёт минимальной цены при автодобавлении и срок его действия | ✅ |
| `POST /v1/product/action/timer/update` | `auto_actions.py` | Продлить учёт минимальной цены | — (запись) |
| `POST /v1/warehouse/fbo/list` | `recommended_prices.get_fbo_warehouse_ids` | Склады FBO | ❌ **400: обязательно `filter_by_supply_type`** |

### Акции

С 19 октября 2026 действует «единая цена» — подробно в разделе об акциях (будет добавлен).

| Метод | Модуль | Назначение | Статус |
|---|---|---|---|
| `GET /v1/actions` | `recommended_prices.get_actions_list` | Список доступных акций | ✅ |
| `POST /v2/actions/products` | `recommended_prices.fetch_action_page` | Участники акции; деньги — объекты `{amount, currency}` | ✅ |
| `POST /v2/actions/candidates` | `recommended_prices.fetch_action_page` | Кандидаты в акцию | ✅ |
| `POST /v1/actions/products/update` | `recommended_prices.activate_products_in_action` | Добавить товар / изменить цену в акции / вывести из бустинга ценой выше лимита | ✅ |
| `POST /v2/actions/products/deactivate` | `recommended_prices.deactivate_products_in_action` | Удалить из акции «Промокоды»; для бустингов отвечает code 9 | ✅ (для бустингов заблокирован) |
| `POST /v2/actions/discounts-task/list` | `recommended_prices.py` | Заявки покупателей на скидку; `limit` ∈ {5,10,15,20,30,50} | ✅ |
| `POST /v1/actions/discounts-task/approve` / `decline` | `recommended_prices.py` | Одобрить / отклонить заявки | — (запись) |
| `v1/actions/products`, `v1/actions/candidates`, `v1/actions/products/activate`, `v1/actions/products/deactivate` | — | Старые методы; в коде больше не используются | ❌ заблокированы («unified price logic») |

### Аналитика

| Метод | Модуль | Назначение | Статус |
|---|---|---|---|
| `POST /v1/analytics/product-queries` | `recommended_prices.get_search_visibility`, `organic_analysis.load_search`, `search_api.fetch_product_queries_day` | Видимость в поиске. Без Premium позиции и просмотры `null`, а период должен заканчиваться **не позже полуночи вчерашнего дня (UTC)** и быть не длиннее ~30 дней, иначе 403 «available starting from the premium subscription». Период — полуинтервал `[date_from, date_to)`: `date_from = date_to` даёт пустой ответ, один день — `D 00:00Z … D+1 00:00Z` | ✅ |
| `POST /v1/analytics/product-queries/details` | `organic_analysis.py`, `search_api.fetch_product_queries_day` | Запросы, по которым находят товар, до 15 на SKU. Без подписки отвечает 200: текст запроса, `query_index`, `unique_search_users`; позиция и просмотры `null`, `order_count` и `gmv` — 0 | ✅ |
| `POST /v1/analytics/data` | `search_api.fetch_sales_by_day` | Заказано штук и на сумму по SKU и дням. Не чаще 1 раза в минуту; без Premium Plus — 50 запросов в сутки, данные за 3 месяца, и метрики поиска (`hits_view_search`, `position_category` …) **молча выбрасываются** из ответа — значений меньше, чем запрошено | ✅ |
| `POST /v1/search-queries/text`, `/v1/search-queries/top` | — | Частотность запросов по всему Ozon. Только Premium Pro, иначе 403 `code 7` | ❌ без подписки |
| `POST /v1/analytics/local-sale/items-clusters/info` | `organic_analysis.load_locality` | Доля доставок из кластера покупателя, рекомендованная поставка | — |

### Performance API (реклама)

| Метод | Модуль | Назначение |
|---|---|---|
| `POST /api/client/token` | `performance_api.get_performance_token` | Токен доступа |
| `GET /api/client/campaign` | `performance_api.py`, `ads_cpc.py` | Кампании (фильтры `advObjectType`, `state`) |
| `GET /api/client/campaign/{id}/v2/products`, `POST .../products`, `.../products/delete`, `.../products/bids/competitive` | `ads_cpc.py` | Товары кампании, ставки, конкурентные ставки |
| `GET /api/client/min/sku` | `ads_cpc.py` | Минимальные ставки |
| `POST /api/client/statistics`, `GET /api/client/statistics/{uuid}`, `GET /api/client/statistics/report` | `performance_api.py`, `ads_cpc.py` | Асинхронный отчёт статистики: запрос → опрос статуса → скачивание |
| `GET /api/client/statistics/daily/json`, `.../campaign/product/json`, `.../json` | `ads_cpc.py`, `organic_analysis.py`, `performance_api.py` | Статистика без асинхронного отчёта |

## Переходы на новые версии (октябрь 2026)

| Было | Стало | Что поменялось в данных | Где учтено |
|---|---|---|---|
| `v3/posting/fbs/list` (offset, один `status` на запрос) | `v4/posting/fbs/list` | Курсор `cursor`/`has_next`, `filter.statuses` списком, цена товара `{amount, currency}` | `Monthly_sales_report._fetch_fbs_page`, общий `_fetch_posting_page` нормализует цены как у FBO |
| `v3/finance/transaction/list` | `v1/finance/accrual/by-day` | Начисления по дням, суммы — объекты денег | `get_transactions_by_posting` |
| `v1/actions/products`, `v1/actions/candidates` | `v2/actions/products`, `v2/actions/candidates` | Все деньги — `{amount: "960", currency: "RUB"}`, ответ без обёртки `result` | `normalize_action_product` превращает деньги в `float` и запоминает `currency` |
| `v1/actions/products/activate` | `v1/actions/products/update` | `action_price` — объект денег; `stock` только для «Промокодов»; ответ `active_product_ids` / `deactivated_product_ids` (строки), `rejected`, `warnings` | `activate_products_in_action` |
| `v1/actions/products/deactivate` | `v2/actions/products/deactivate` (только «Промокоды»), для бустингов — `update` с ценой выше лимита | `product_ids` — строки | `deactivate_products_in_action`, `exit_products_from_actions` / `execute_action_plan` |
| `v1/analytics/product-queries` с концом периода «вчера, текущее время» | конец периода — полночь вчерашнего дня | — | `get_search_visibility` |
| `v2/product/info` (удалён, 404) | `v3/product/info/list` с `product_id` списком | Ответ `items[]` с `id`, `offer_id`, `sku` | `get_product_info_by_id` |

Общие приёмы, которые помогли при переходах:
- **Официальный swagger.** Ozon раздаёт его по запросу; сторонний SDK
  [dev-ik/seller-sdk](https://github.com/dev-ik/seller-sdk) держит актуальный snapshot в `docs/ozon/swagger.json`.
  Описания методов в swagger содержат дату переключения и предупреждения о поведении.
- **Проверка без последствий.** Метод записи можно проверить запросом, который ничего не меняет:
  отправить ту же цену, что уже стоит (Ozon вернёт `rejected: "Product update has no changes"`),
  или удалить из акции товар, который в ней не участвует.
- **Сверка результата.** После перевода загрузки на новую версию сравните выборку с уже сохранёнными
  отчётами: при переходе на `v4/posting/fbs/list` списки заказов FBS за август и сентябрь 2026
  совпали с сохранёнными отчётами один в один (125 и 25 заказов).

## Правила для нового кода

1. Для чтения больших объёмов используйте `report_http.create_session()` — получите темп и
   разумные повторы при 429 бесплатно.
2. Пагинация — с защитой от зацикливания: запоминайте увиденные курсоры/`last_id` и падайте при повторе
   (`_get_postings`, `get_transactions_by_posting`).
3. Ошибка любой страницы — исключение, а не пустой результат. Пустой список от упавшего запроса
   выглядит как «нет данных» и ведёт к неверным выводам (пример — `get_fbo_warehouse_ids`, см. [99](99-known-issues.md)).
4. Деньги из новых методов приходят строками внутри объекта — переводите в число в одном месте
   при чтении ответа.
5. Не храните ключи в коде и логах; в сообщениях об ошибках печатайте метод и `message` Ozon, но не заголовки.
