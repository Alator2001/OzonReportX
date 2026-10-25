# 18. Сборка приложения

Готовое приложение без Python: macOS — `OzonReportX.app`, Windows — папка с `OzonReportX.exe`.
Собирается `flet pack` (PyInstaller + клиент Flet) скриптом `packaging/build.py`; для релиза — GitHub Actions.

## Где код и где данные — `scripts/app_paths.py`

| | Из исходников | В сборке |
|---|---|---|
| `CODE_ROOT` — код, `scripts/`, `img/`, `version.txt` | корень репозитория | `sys._MEIPASS` (`.app/Contents/Frameworks`, `_internal\` на Windows) — только чтение |
| `DATA_ROOT` — `.env`, `costs.xlsx`, `reports/`, настройки, `.cache/` | корень репозитория | папка из «Настройки» → «Папка с данными», иначе `~/Documents/OzonReportX` |
| `python_executable()` — чем запускать скрипты | `.venv` → текущий python | сама сборка (`sys.executable`) |

Порядок для `DATA_ROOT` (`_data_root`, второе значение — `DATA_ROOT_SOURCE`):
1. `OZONREPORTX_HOME` — в обоих режимах (`env`). На macOS из Терминала: `open --env OZONREPORTX_HOME=… OzonReportX.app`
   (обычный `VAR=… open` переменную приложению не передаёт).
2. Только в сборке — папка, выбранная в «Настройки» → «Папка с данными» (`chosen`). Путь хранится вне папки данных:
   `~/Library/Application Support/OzonReportX/data_folder.txt` (macOS), `%APPDATA%\OzonReportX\data_folder.txt` (Windows).
   Если папки больше нет (отключён диск) — берётся папка по умолчанию.
3. По умолчанию (`default`): из исходников — корень репозитория, как раньше; в сборке — `~/Documents/OzonReportX`.

Новая папка действует со следующего запуска. Чтобы до перезапуска скрипты не ушли в новую папку, `launcher.run_app`
записывает текущую в `OZONREPORTX_HOME` — её наследуют все процессы, запущенные из окна. Карточка «Папка с данными»
есть только в сборке: из исходников данные всегда в репозитории.

Папка выбирается системным диалогом (`pick_folder` в `app_flet.py`: `osascript choose folder` на macOS,
`FolderBrowserDialog` через PowerShell на Windows), а не `ft.FilePicker`: клиенту Flet на macOS не хватает
разрешения на доступ к файлам (`ENTITLEMENT_NOT_FOUND`), и его диалог ронял окно программы. Строки для AppleScript
экранируются `json.dumps(…, ensure_ascii=False)` — запись `\uXXXX` AppleScript не понимает.

Правило для нового кода: пути к данным — от `DATA_ROOT` (или от `repo_root`, который приходит параметром), пути
к скриптам — `SCRIPTS_DIR` / `script_command()`, `.env` — `load_dotenv(ENV_FILE)`. Не считать корень от `__file__`:
в сборке это папка с кодом, куда писать нельзя.

## Запуск скриптов в сборке — `launcher.py`

Приложение запускает отчёты и прочее отдельными процессами (`stream_process`, `subprocess.run`). Отдельного
python в сборке нет, поэтому команда — `[сборка, <CODE_ROOT>/scripts/X.py, аргументы…]`: `launcher.py` видит
первым аргументом `.py` и выполняет файл через `runpy.run_path` как `__main__`, добавив `scripts/` и корень
в `sys.path`. Без аргументов — открывает окно (`chdir` в `DATA_ROOT`, `ft.run(app_flet.main)`).

- Исходники `scripts/` кладутся в сборку файлами (`--add-data`), а все модули `scripts.*` — `--hidden-import`,
  чтобы PyInstaller собрал зависимости и тех скриптов, которые приложение только запускает.
- В оконной сборке Windows `sys.stdout` бывает `None` даже при трубе от родителя — `launcher._utf8_stream`
  открывает дескриптор 1/2 сам и включает UTF-8 с построчным сбросом (вывод отчёта идёт в журнал приложения по строкам).
- Сборка всегда «папкой» (`--onedir`): однофайловая распаковывалась бы во временную папку при каждом запуске скрипта.

## Клиент Flet

`flet pack` кладёт клиент Flet в сборку архивом (`flet_desktop/app/flet-macos.tar.gz` / `.zip`) с нашим именем
и иконкой. При первом запуске `flet_desktop.ensure_client_cached` распаковывает его в
`~/.flet/client/flet-desktop-full-<версия>`; интернет не нужен. Если этот кэш уже есть (на машине разработчика —
от `flet run`), берётся он: окно тогда называется «Flet», это нормально.

## Сборка локально

```bash
.venv/bin/pip install -r config/requirements-build.txt
.venv/bin/python packaging/build.py            # версия — из README.md, или --version 3.4.0
```

Результат — `dist/OzonReportX-<версия>-macOS-<arm64|x86_64>.zip` или `dist/OzonReportX-<версия>-Windows.zip`
(~120 МБ). PyInstaller собирает только под текущую ОС и архитектуру. `flet pack` очищает `build/` и `dist/`,
поэтому иконки и сам скрипт лежат в `packaging/`. Иконки (`icon.icns`, `icon.ico`, `icon.png`) сделаны из
`img/ozonreportx-mark.svg`.

Проверка без своих данных: `OZONREPORTX_HOME=/tmp/ozon-test dist/OzonReportX.app/Contents/MacOS/OzonReportX`;
скрипт — тем же файлом с путём `…/Contents/Frameworks/scripts/<скрипт>.py`.

## Релиз — `.github/workflows/build.yml`

На публикацию релиза (или вручную с тегом) собирает на `macos-14` (Apple Silicon), `macos-15-intel` и
`windows-latest`, запускает в сборке `_auto_update.py` (скрипт выполняется, вывод доходит, версия верная)
и прикладывает архивы к релизу (`gh release upload --clobber`).

## Обновление

Сборка не может заменить сама себя: `_auto_update.py` при `FROZEN` сообщает о новой версии и по «Скачать
обновление» открывает страницу релиза. Данные в `Документы/OzonReportX` новая версия подхватывает.

## Ограничения

- **Нет подписи Apple и Microsoft.** macOS-сборка подписана ad-hoc (`codesign -s -`), без этого на Apple Silicon
  она не запустится вовсе; при первом запуске — правой кнопкой → «Открыть». Windows SmartScreen предупреждает
  о неизвестном издателе. Нотаризация требует Apple Developer ID ($99/год), подпись Windows — сертификат.
- **Парсер выдачи (Playwright) в сборку не входит** (`EXCLUDE` в `build.py`): ему нужен свой браузер на ~150 МБ.
  В сборке он недоступен; остальной раздел «Поиск» работает.
- Консольное меню (`OzonReportX.bat` / `.command`) и `first_run_setup.py` — только для версии из исходников.
