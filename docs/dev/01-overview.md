# 01. Обзор проекта

## Что это

OzonReportX — локальное настольное приложение продавца Ozon (и частично Wildberries).
Оно работает на компьютере пользователя, ходит в API маркетплейсов с ключами из `.env`
и складывает результаты в Excel/JSON-файлы в папке проекта. Сервера у программы нет,
база данных не используется: **состояние программы — это файлы в корне репозитория**.

Основные функции:

| Функция | Где в интерфейсе | Код |
|---|---|---|
| Месячный отчёт по продажам и бизнес-сводка | «Бизнес-сводка» | `scripts/Monthly_sales_report.py`, [04](04-monthly-report.md) |
| Расчёт минимальной и желательной цены | «Цены и скидки» | `scripts/recommended_prices.py` |
| Синхронизация текущих цен и акций в `costs.xlsx` | «Цены и скидки» | `scripts/price_management.py` |
| Управление участием в акциях | «Акции» | `scripts/price_management.py`, `scripts/recommended_prices.py` |
| Автодобавление в акции, учёт минимальной цены | «Акции» | `scripts/auto_actions.py` |
| Заявки покупателей на скидку | «Цены и скидки» | `scripts/price_management.py` |
| Реклама с оплатой за клик (ставки) | «Реклама» | `scripts/ads_cpc.py`, `scripts/performance_api.py` |
| Анализ органических продаж | «Анализ» | `scripts/organic_analysis.py` |
| Поиск и ключевые слова: спрос, семантика, покрытие карточек, позиции | «Поиск» | `scripts/search_*.py`, `scripts/serp_*.py`, [17](17-search.md) |
| План поставки FBO | «Поставки» | `scripts/fbo_supply_plan.py` |
| ABC/XYZ-анализ | «ABC/XYZ-анализ» | `scripts/ABC_XYZ_analytics_report.py`, `scripts/abc_view.py` |
| Баланс Ozon | «Настройки» → «Проверить баланс Ozon»; месячный отчёт | `scripts/balance_report.py` |
| Wildberries: сводка, поставки | переключатель магазина → WB | `scripts/wb_*.py` |
| AI-ассистент (локальный, Ollama) | «AI-ассистент» | `scripts/ai_agent.py`, `ai_tools.py`, `ai_warehouse.py`, `ai_chat.py`, [14](14-ai-assistant.md) |

Разделы интерфейса заданы списком `section_names` в `app_flet.py:7734`:
«Бизнес-сводка», «AI-ассистент», «ABC/XYZ-анализ», «Цены и скидки», «Акции», «Реклама»,
«Анализ», «Поиск», «Поставки», «Остатки», «Настройки», «Файлы». Раздел «Финансы» (выгрузка баланса за период) убран после 3.2.0.

## Структура репозитория

```
OzonReportX/
├── app_flet.py               # настольное приложение (Flet), ~8 тыс. строк, всё в main(page)
├── OzonReportX_UI.bat        # запуск приложения (основной способ), Windows
├── OzonReportX.bat           # консольное меню (старый интерфейс, first_run_setup.py)
├── Export_Data.bat / Import_Data.bat   # перенос данных на другой компьютер
├── *.command                 # те же четыре лаунчера для macOS (bash, должны быть исполняемыми)
├── .gitattributes            # держит *.command и *.sh с LF
├── version.txt               # текущая версия, читается автообновлением
├── .env.example              # шаблон ключей API
├── config/
│   ├── requirements.txt      # минимальные версии зависимостей
│   ├── requirements.lock.txt # точные версии — именно их ставит first_run_setup.py
│   └── show_banner.py        # баннер консольного запуска
├── scripts/                  # вся бизнес-логика; каждый модуль можно запустить отдельно
├── tests/                    # unittest, без сети
├── img/                      # картинки для README и иконка приложения
└── docs/dev/                 # эта документация
```

Модули `scripts/`:

| Модуль | Назначение |
|---|---|
| `Monthly_sales_report.py` | Месячный отчёт: заказы FBS/FBO, начисления, баланс, реклама → `reports/Ozon/<Год>/<Месяц> <Год>.xlsx` |
| `report_http.py` | `ReportSession` и `create_session()`: темп запросов и повторы при 429 для Seller API |
| `balance_report.py` | `/v1/finance/balance`: звёздные товары, хранение FBO, выплаты; выгрузка в `balance reports/` |
| `performance_api.py` | Performance API (реклама): токен, кампании, статистика — для месячного отчёта |
| `recommended_prices.py` | Расчёт цен по марже, работа с акциями (чтение, добавление, вывод), остатки, видимость в поиске; ~2,3 тыс. строк |
| `price_management.py` | Меню и планы действий: синхронизация `costs.xlsx`, планы акций, календарь акций, заявки на скидку |
| `update_prices.py` | Отправка минимальных цен из `costs.xlsx` в Ozon (`/v1/product/import/prices`) |
| `auto_actions.py` | Автодобавление в акции и учёт минимальной цены: чтение, изменение, журнал |
| `ads_cpc.py` | Реклама «оплата за клик»: план ставок по экономике товара и отправка |
| `organic_analysis.py` | Диагностика органических продаж по товарам |
| `search_db.py`, `search_api.py`, `search_collect.py`, `search_semantics.py`, `serp_*.py` | Поиск и ключевые слова: база SQLite с миграциями, сбор поисковой аналитики и карточек, леммы, кластеры, покрытие карточки, приоритет запросов; парсер выдачи по регионам ([17](17-search.md)) |
| `fbo_supply_plan.py` | План поставки FBO по кластерам → `stocks reports/Поставка FBO <дата>.xlsx` |
| `fbo_supply_report.py` | **Устаревший** расчёт поставок по ABC/XYZ; нигде не вызывается (см. [99](99-known-issues.md)) |
| `ABC_XYZ_analytics_report.py` | ABC/XYZ по месячным отчётам → `ABC&XYZ reports/` |
| `abc_view.py` | Чтение готового ABC/XYZ-отчёта для экрана приложения |
| `ui_actions.py` | CLI-обёртка для приложения: `save-margin`, `recommended-prices`, `refresh-pricing`, `update-min-prices`, `discount-requests` |
| `first_run_setup.py` | Подготовка окружения (`.venv`, зависимости) и консольное меню |
| `marketplace_settings.py` | Чтение и сохранение ключей в `.env` (общая логика для приложения и консоли) |
| `_auto_update.py` | Автообновление из GitHub Releases |
| `transfer_data.py` | Экспорт/импорт личных данных в zip |
| `file_io.py`, `file_lock.py` | Атомарная запись файлов и блокировки (см. ниже) |
| `ui_log.py` | Журнал сессии приложения |
| `utils.py` | `log_verbose`, `prompt_yes_no`, `print_step` |
| `ai_chat.py`, `ai_agent.py` | AI-ассистент на Ollama: клиент и агентный цикл ([14](14-ai-assistant.md)) |
| `ai_eval.py` | Контрольные вопросы AI-ассистента на живой модели ([14](14-ai-assistant.md#контрольные-вопросы--ai_evalpy)) |
| `ai_warehouse.py` | Аналитическая база AI-ассистента: все данные программы в одном SQLite и словарь данных ([14](14-ai-assistant.md#аналитическая-база)) |
| `ai_tools.py` | Инструменты AI-ассистента поверх базы и живого Ozon (только чтение) ([14](14-ai-assistant.md#инструменты--ai_toolspy)) |
| `wb_*.py` | Wildberries: финансы, заказы, воронка, возвраты, приёмка, хранение, сводка, план запасов |

## Запуск

### Обычный путь — `OzonReportX_UI.bat`

1. Если нет `.venv\Scripts\python.exe`, вызывается системный `python scripts\first_run_setup.py --setup-only`:
   создаётся `.venv`.
2. Повторно `first_run_setup.py --setup-only`, уже из `.venv`: ставятся зависимости.
3. Запускается `.venv\Scripts\python.exe app_flet.py`.

`--setup-only` делает только `ensure_venv` и `ensure_deps` (`scripts/first_run_setup.py:593`).
Мастер ключей, автообновление и шаблон `costs.xlsx` в этом режиме **не запускаются**:
ключи настраиваются мастером в самом приложении (`needs_setup` → `open_marketplace_settings`,
`app_flet.py:8164`), обновление — кнопкой в «Настройках».

### macOS — `OzonReportX_UI.command`

Делает то же самое, что `.bat`, но с `.venv/bin/python`. Системный Python ищется как `python3.14`,
затем `python3` (команды `python` на macOS нет); в `PATH` добавляются `/opt/homebrew/bin` и `/usr/local/bin`,
потому что при двойном клике в Finder профиль оболочки может не подхватить Homebrew.
Лаунчеры `OzonReportX.command`, `Export_Data.command`, `Import_Data.command` повторяют соответствующие `.bat`.

Код приложения кроссплатформенный: интерпретатор ищется в `.venv/Scripts/python.exe` и `.venv/bin/python`
(запасной вариант — `sys.executable`), файлы открываются через `open`, блокировки — через `fcntl`,
отмена задачи убивает всю группу процессов (см. [02](02-app-architecture.md#отмена)).

Права на запуск `.command` хранятся в git (`100755`). Новый лаунчер добавляйте с
`git update-index --chmod=+x <файл>`. Автообновление распаковывает zipball через `zipfile`, который права
не сохраняет, поэтому `apply_update` заново ставит бит запуска на `*.command` и `*.sh`.

### Консольный путь — `OzonReportX.bat`

Запускает `first_run_setup.py` без флагов: окружение → мастер ключей в терминале
(`marketplace_settings.configure_console`) → проверка обновлений → текстовое меню.
Консольное меню — старый интерфейс; новые функции (анализ, реклама, календарь акций, WB)
есть только в приложении. Флаги `--yes` / `--no` отвечают на все вопросы `prompt_yes_no` автоматически.

### Зависимости

`ensure_deps` ставит `config/requirements.lock.txt` (если его нет — `requirements.txt`)
и запоминает SHA-256 файла в `.venv/.bootstrap_done`. Пока хеш не изменился, установка пропускается.
После изменения зависимостей обновите lock-файл: `.venv\Scripts\python.exe -m pip freeze`
(на macOS — `.venv/bin/python -m pip freeze`). Lock-файл не должен содержать пакетов только для одной ОС.

Python — **3.14** (`.python-version`). Ключевые библиотеки: `flet 0.84`, `pandas 2.3`,
`openpyxl 3.1`, `requests 2.32`, `matplotlib`, `plotly`, `python-dotenv`, `pymorphy3` (леммы запросов).

### Запуск модулей по отдельности

Почти каждый модуль `scripts/` — самостоятельный скрипт (`if __name__ == "__main__"`).
Приложение запускает их **отдельными процессами** через `stream_process` (см. [02](02-app-architecture.md)),
поэтому любой сценарий можно воспроизвести из терминала:

```bash
.venv/Scripts/python.exe scripts/Monthly_sales_report.py --month 9 --year 2026
.venv/Scripts/python.exe scripts/fbo_supply_plan.py --lead-days 14 --cover-days 30 --service-level 0.9
.venv/Scripts/python.exe scripts/ui_actions.py refresh-pricing
```

На macOS вместо `.venv/Scripts/python.exe` — `.venv/bin/python`.

Модули импортируют друг друга двумя способами — как пакет (`from scripts.x import ...`, когда
импортирует `app_flet.py`) и как соседний файл (`from x import ...`, когда скрипт запущен напрямую).
Отсюда повсеместный шаблон `try: from scripts.x ... except ModuleNotFoundError: from x ...`.
`file_lock.py` дополнительно регистрирует себя под обоими именами в `sys.modules`, чтобы
реестр блокировок был один на процесс.

## Окружение: `.env`

| Переменная | Обязательна | Для чего |
|---|---|---|
| `OZON_CLIENT_ID`, `OZON_API_KEY` | для Ozon | Seller API. Заполняются только парой |
| `OZON_PERF_CLIENT_ID`, `OZON_PERF_API_KEY` | нет | Performance API (реклама). `CLIENT_ID` вида `...@advertising.performance.ozon.ru`, `API_KEY` — это client secret |
| `WB_API_TOKEN` | для WB | Токен Wildberries |
| `MARKETPLACE_SETUP_VERSION` | служебная | Пишется мастером настройки. Пока значение ≠ `"1"`, приложение при запуске открывает мастер |
| `OLLAMA_HOST`, `OLLAMA_MODEL` | нет | AI-ассистент, по умолчанию `http://localhost:11434` и `qwen3:4b` |
| `OLLAMA_NUM_CTX`, `OLLAMA_KEEP_ALIVE`, `OLLAMA_THINK`, `OLLAMA_READ_TIMEOUT` | нет | Окно контекста (32768), время удержания модели в памяти (`30m`), режим размышлений, таймаут ответа — см. [14](14-ai-assistant.md#подключение-к-ollama) |
| `SERP_PARSER_ENABLED` | нет | `1` — включить замер позиций в выдаче Ozon ([17](17-search.md#парсер-выдачи)) |
| `SERP_PROXY_*` | нет | Адрес прокси региона для парсера выдачи; имя переменной задаётся при создании профиля |
| `OZON_VERBOSE` | нет | `1` — подробный вывод (`utils.log_verbose`) |
| `OZONREPORTX_UI_PROMPTS` | служебная | Ставит приложение для месячного отчёта: разрешает вопросы о маркетинге при не-TTY stdin (см. [04](04-monthly-report.md#ручной-ввод-маркетинга)) |

`marketplace_settings.save_settings` пишет `.env` атомарно, сохраняя посторонние строки,
и не даёт сохранить только половину пары ключей Ozon. Подключение к API при сохранении **не проверяется**.

Большинство модулей читают ключи **при импорте** (`load_dotenv()` + `os.getenv` на уровне модуля).
Следствия: изменение `.env` подхватывается только новыми процессами (для функций, которые приложение
вызывает в своём процессе, есть `recommended_prices.refresh_api_credentials`, см. [03](03-ozon-api.md)).
`Monthly_sales_report.py` без ключей бросает `RuntimeError` при импорте — это допустимо, потому что
приложение запускает его отдельным процессом и не импортирует.

Модули, которые **импортирует само приложение**, не должны требовать ключей при импорте: иначе приложение
не запустится у пользователя без ключей Ozon (например, только с Wildberries) и не покажет мастер настройки.
`balance_report.py` поэтому собирает заголовки в `_headers()` при каждом запросе и без ключей бросает
`RuntimeError` только в момент запроса (тест `test_startup_and_update.StartWithoutKeysTests`).

## Файлы данных

Всё ниже исключено из Git (`.gitignore`) и считается данными пользователя.

| Путь | Кто пишет | Что внутри |
|---|---|---|
| `.env` | мастер настройки | Ключи API |
| `costs.xlsx` | пользователь + синхронизация цен | Себестоимость, рассчитанные и текущие цены, лист «Акции». Главный справочник товаров |
| `margin_settings.json` | «Цены и скидки» | Минимальная и желаемая маржа |
| `reports/Ozon/<Год>/<Месяц> <Год>.xlsx` | месячный отчёт | Лист «Заказы» + блок показателей P/Q, «Расходы вне заказов», «Кампании» |
| `ABC&XYZ reports/` | ABC/XYZ | `<Месяц Год>-<Месяц Год>.xlsx` |
| `stocks reports/` | планы поставок | Отчёты FBO/WB, кэш заказов `.orders_cache/`, кэш продаж WB `.wb_sales_cache/` |
| `balance reports/` | консольное меню (`balance_report.py`) | JSON и Excel баланса за период |
| `fbo_supply_settings.json`, `wb_supply_settings.json` | «Поставки» | Параметры и зоны складов |
| `ads_settings.json`, `ads reports/` | «Реклама» | Настройки подбора ставок; журнал отправленных изменений ставок (JSONL) |
| `analytics history/` | «Анализ» | `organic_last.json` — последний результат; `snapshots/` — по одному снимку в день (цена, остаток, индекс цен, контент-рейтинг, число акций) |
| `search data/search.db` | сборщик поисковой аналитики | SQLite: товары, запросы и кластеры, статистика запросов и продажи по дням, снимки карточек, регионы и позиции, журнал сбора; `serp_profiles/` — профили регионов (cookies анонимной сессии Ozon), `serp_raw/` — сырые ответы выдачи за 7 дней ([17](17-search.md)) |
| `search_settings.json` | пользователь (необязателен) | Порог кластеров, кривая CTR, зона роста — поверх значений по умолчанию ([17](17-search.md#настройки)) |
| `actions history/auto_actions_log.jsonl` | «Акции» → автоакции | Журнал изменений автодобавления |
| `wb reports/`, `wb_costs.xlsx` | Wildberries | Финансовые отчёты по месяцам и себестоимость WB |
| `.cache/` | приложение | Журналы сессий `.cache/logs/ui-*.log`, картинки графиков, временные файлы тестов; `.cache/ai/warehouse.db` — аналитическая база AI-ассистента (производная, можно удалить) |
| `.<имя>.ozon-lock` | `file_lock.py` | Файлы-замки рядом с защищаемыми файлами |
| `~tmp_*` | `file_io.py` | Временные файлы атомарной записи; при сбое удаляются |

### Атомарная запись и блокировки

- `file_io.atomic_output_path(path)` — пишем во временный `~tmp_*` в той же папке и только после
  успешного завершения делаем `os.replace`. Если запись упала, старый файл остаётся целым.
- `file_io.costs_excel_writer` — заменяет выбранные листы `costs.xlsx`, сохраняя остальные
  (копирует файл во временный, правит его, публикует). Попутно переименовывает старый первый лист в «Основной».
- `file_lock.exclusive_file(path)` — неблокирующая блокировка: реестр `RLock` внутри процесса
  плюс OS-блокировка файла `.<имя>.ozon-lock` (`msvcrt.locking` / `fcntl.flock`) между процессами.
  Повторный вход в том же потоке разрешён. Занято — сразу `FileBusyError` с понятным текстом,
  ожидания нет. Декоратор `locked_costs` держит блокировку `costs.xlsx` на всю операцию
  «прочитать — посчитать — записать».
- `file_lock.file_signature(path)` — `(mtime_ns, size, inode)`; редактор таблицы в приложении
  сравнивает подпись, чтобы не перезаписать `costs.xlsx`, изменённый снаружи.

## Обновление и перенос

**Автообновление** (`scripts/_auto_update.py`): читает `version.txt`, спрашивает GitHub
`releases/latest` репозитория `Alator2001/OzonReportX`, сравнивает версии (`packaging.version`),
делает резервную копию в `../backup_v<версия>`, скачивает zipball и заменяет файлы через
промежуточную папку с откатом при ошибке. Не трогает пользовательские данные из списка `preserve`
(`.env`, `costs.xlsx`, папки отчётов и т. п.); файлы, исключённые из Git, в zipball и так не попадают.

Режимы запуска:

| Как запущен | Поведение |
|---|---|
| Из терминала (консольное меню) | Показывает новую версию, спрашивает «Установить обновление сейчас? [Y/n]», после установки перезапускает процесс через `os.execv` |
| Без терминала, без флагов (кнопка «Проверить обновления» в приложении) | Только показывает новую версию и что изменилось; **ничего не ждёт и не ставит** — вопрос через `input()` в приложении ждал бы ответа вечно |
| `--yes --no-restart` (кнопка «Установить обновление») | Ставит без вопроса и не перезапускает скрипт; просит перезапустить приложение |

Чтобы выпустить версию: обновить `version.txt` → коммит → тег → GitHub Release на этом теге.

**Перенос** (`scripts/transfer_data.py`): `export` упаковывает `.env`, `costs.xlsx`, `wb_costs.xlsx`,
`margin_settings.json` и папки отчётов в `OzonReportX_transfer_<дата>.zip`; `import` распаковывает,
предварительно копируя перезаписываемые файлы в `backup_<дата>/`, и отбрасывает пути вне проекта.
Список переносимого неполный — см. [99](99-known-issues.md).

Особенности для macOS (общие для обеих ОС, тесты — `tests/test_transfer_data.py`):
- `import` принимает и zip, и распакованную папку: Safari по умолчанию распаковывает скачанный zip сам.
- Пропускаются `__MACOSX/`, `._*` и `.DS_Store`. Если все записи лежат в одной лишней папке
  (архив пересжат через Finder), она снимается — кроме случая, когда это папка программы, например `reports`.
- Имена приводятся к NFC и при экспорте, и при импорте: macOS может хранить кириллицу разложенной (NFD:
  «й» = «и» + знак), а приложение ищет месяцы в именах отчётов регулярками с составными буквами.
- `Import_Data.command`: Finder не передаёт перетащенный на `.command` файл, поэтому архив берётся из аргумента,
  затем самый свежий `OzonReportX_transfer_*` в папке проекта, затем в `~/Downloads`, иначе — окно выбора
  файла (`osascript … choose file`). Перед импортом показывается путь и спрашивается подтверждение.
- `Export_Data.command` после экспорта показывает архив в Finder (`open -R`).

## Тесты

```bash
.venv/Scripts/python.exe -m unittest discover -s tests   # Windows
.venv/bin/python -m unittest discover -s tests           # macOS
```

Запускать именно через `discover -s tests`: тесты импортируют друг друга как модули верхнего
уровня (`import test_technical_fixes`), поэтому `python -m unittest tests.test_x` не работает.
`pytest` в зависимостях нет. Сеть в тестах не используется: `test_technical_fixes.OfflineTests.setUp`
подменяет сетевые вызовы, остальные тесты мокают `requests`. Временные файлы — в `.cache/technical-fix-tests/`.
