# -*- coding: utf-8 -*-
"""
Модуль для чата с ИИ на тему бизнеса Ozon.
"""

import os
from contextlib import ExitStack, closing
import sys
import re
import json
from datetime import date
import requests
import pandas as pd
from pathlib import Path
from typing import Optional, Tuple, Dict, Any, List
from dotenv import load_dotenv
from openpyxl import load_workbook
try:
    from scripts.file_io import read_costs_dataframe
except ModuleNotFoundError:
    from file_io import read_costs_dataframe

# Загружаем переменные окружения
load_dotenv()
OLLAMA_HOST = (os.getenv("OLLAMA_HOST") or "http://localhost:11434").rstrip("/")
OLLAMA_CHAT_URL = f"{OLLAMA_HOST}/api/chat"

# Импортируем функции для работы с отчётами
try:
    script_dir = Path(__file__).resolve().parent
    if str(script_dir) not in sys.path:
        sys.path.insert(0, str(script_dir))
    from recommended_prices import (
        REPORTS_DIR_NAME,
        MONTHS_RU,
        ORDER_SHEET,
    )
    from analytics_models import ArtikulAggregateState, YearlyArtikulSummaryState, WorkflowState
    from analytics_reducers import (
        build_workflow_fallback_answer_core,
        build_yearly_artikul_profit_summary_core,
        extract_year_from_text as extract_year_from_text_core,
        is_yearly_artikul_report_request_core,
    )
    from agent_workflows import (
        build_workflow_state_core,
        detect_workflow_core,
        get_workflow_followup_needs_core,
    )
    from payload_policy import format_model_payload_core
except ImportError:
    REPORTS_DIR_NAME = "reports"
    ORDER_SHEET = "Заказы"
    MONTHS_RU = [
        "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
    ]

    def get_report_path(repo_root: Path, year: int, month: int) -> Path:
        name = f"{MONTHS_RU[month - 1]} {year}.xlsx"
        return repo_root / REPORTS_DIR_NAME / name
    from analytics_models import ArtikulAggregateState, YearlyArtikulSummaryState, WorkflowState
    from analytics_reducers import (
        build_workflow_fallback_answer_core,
        build_yearly_artikul_profit_summary_core,
        extract_year_from_text as extract_year_from_text_core,
        is_yearly_artikul_report_request_core,
    )
    from agent_workflows import (
        build_workflow_state_core,
        detect_workflow_core,
        get_workflow_followup_needs_core,
    )
    from payload_policy import format_model_payload_core

DEFAULT_MODEL = "qwen3:4b"
OLLAMA_MODEL = (os.getenv("OLLAMA_MODEL") or "").strip() or DEFAULT_MODEL
CURRENT_DATE = date.today().isoformat()
CURRENT_YEAR = date.today().year

SYSTEM_PROMPT = r"""
Ты — локальный AI-аналитик OzonReportX.
Текущая дата: """ + CURRENT_DATE + r""".
Если пользователь говорит "этот год", это """ + str(CURRENT_YEAR) + r""" год.

Главная задача: выбрать правильные инструменты и вернуть строго один JSON-объект.

# Формат ответа
Разрешены только 2 формата:
1. {"type":"DATA_REQUEST","needs":[{"tool":"<tool_name>","args":{}}],"reason":"<кратко>"}
2. {"type":"FINAL_ANSWER","answer":"<ответ на русском>"}

Запрещено:
- любой текст до JSON или после JSON;
- поля кроме type, needs, reason, answer;
- форматы {"tool":...}, {"action":...}, {"instrument":...}, {"requests":...}, {"parameters":...}, {"params":...};
- массив вместо объекта.

Если нужны данные, верни только DATA_REQUEST.
Если данных достаточно, верни только FINAL_ANSWER.
Если после получения данных их всё ещё недостаточно, можно вернуть ещё один DATA_REQUEST, но всё равно только в каноническом формате выше.

# Как выбирать инструмент
Используй get_month_summary, если нужен общий разбор одного месяца.
Используй get_month_report_full_data, если нужны все строки, все колонки, детальный отчёт по месяцу или агрегация по заказам.
Используй get_artikul_stats, если пользователь спрашивает про один конкретный артикул.
Используй get_top_profit и get_top_orders, если нужен рейтинг.
Используй list_reports, если сначала нужно понять, какие месяцы вообще доступны.

# Полное описание инструментов
list_reports(args={})
- Аргументы: не нужны.
- Возвращает: список доступных месячных отчётов с полями period, file, path.
- Используй, когда нужно сначала определить, какие месяцы реально есть в системе.
- Особенно важен для запросов вида "за этот год", "за последние месяцы", "за доступные месяцы".

get_month_summary(args={"period":"2026-03"}) или {"period":"Март 2026"}
- Обязательный аргумент: period.
- Возвращает: сводку по месяцу.
- Используй для общего обзора месяца, KPI, выручки, прибыли, маржи, структуры расходов.
- Не используй, если пользователь просит все строки отчёта или полный список заказов.

get_month_report_full_data(args={"period":"Март 2026"}) или {"period":"2026-03"}
- Обязательный аргумент: period.
- Возвращает: полный месячный отчёт целиком, включая строки отчёта, колонки, статистику по артикулам, распределения и агрегаты.
- Это главный инструмент для глубокого анализа месяца и для агрегации по нескольким месяцам.
- Используй его для годовых отчётов по артикулам, продажам и прибыли.
- Если нужно собрать данные за несколько месяцев, верни несколько вызовов get_month_report_full_data, по одному на каждый месяц.

get_top_profit(args={"period":"Март 2026","n":10})
- Обязательный аргумент: period.
- Необязательный аргумент: n, по умолчанию 10.
- Возвращает: top_profit и period.
- Используй, когда пользователь прямо просит топ артикулов по прибыли.

get_top_orders(args={"period":"Март 2026","n":10})
- Обязательный аргумент: period.
- Необязательный аргумент: n, по умолчанию 10.
- Возвращает: top_orders и period.
- Используй, когда пользователь просит топ артикулов по количеству заказов или продаж.

get_artikul_stats(args={"period":"Март 2026","artikul":"1302"})
- Обязательные аргументы: period и artikul.
- Возвращает: stats по одному конкретному артикулу за один период.
- Используй только для одного артикула.
- Не используй для годовых сводок по всем артикулам.

search_artikul(args={"query":"1302"}) или {"query":"часть названия"}
- Обязательный аргумент: query.
- Возвращает: results и count.
- Используй, когда нужно найти артикул по фрагменту текста или когда пользователь не уверен в точном артикуле.

get_costs_columns(args={})
- Аргументы: не нужны.
- Возвращает: список колонок файла costs.xlsx.
- Используй, когда нужно понять структуру файла costs.xlsx.

get_artikul_cost(args={"artikul":"1302"})
- Обязательный аргумент: artikul.
- Возвращает: строку по артикулу из costs.xlsx.
- Используй для проверки себестоимости или полей costs.xlsx по одному артикулу.

get_costs_bulk(args={"artikuls":["1302","1310"]})
- Обязательный аргумент: artikuls.
- Возвращает: results и count по нескольким артикулам из costs.xlsx.
- Используй, когда нужно получить cost-данные сразу для нескольких артикулов.

get_abcxyz_summary(args={"period":"Ноябрь 2025-Декабрь 2025"})
- Обязательный аргумент: period.
- Возвращает: сводку ABC&XYZ отчёта.
- Используй, когда вопрос касается ABC/XYZ анализа.

get_artikul_abcxyz(args={"period":"...","artikul":"1302"})
- Функция пока в разработке и может вернуть error.
- Не выбирай её без необходимости.

get_category_list(args={"period":"...","category":"AX"})
- Функция пока в разработке и может вернуть error.
- Не выбирай её без необходимости.

# Правила для годового отчёта
Если пользователь просит отчёт за год по прибыли или продажам артикулов:
1. Сначала запроси list_reports.
2. После получения list_reports выбери только реально существующие месяцы нужного года.
3. Не запрашивай будущие месяцы, которых нет в list_reports.
4. Затем верни DATA_REQUEST с несколькими get_month_report_full_data, по одному на каждый доступный месяц.
5. После получения месячных отчётов верни FINAL_ANSWER и агрегируй данные по артикулам.

Если в list_reports за """ + str(CURRENT_YEAR) + r""" год есть только Январь, Февраль, Март и Апрель, то используй только эти месяцы.

# Few-shot примеры
Запрос: "Что произошло в марте 2026?"
Ответ:
{"type":"DATA_REQUEST","needs":[{"tool":"get_month_summary","args":{"period":"2026-03"}}],"reason":"Нужна месячная сводка"}

Запрос: "Покажи все данные за март 2026"
Ответ:
{"type":"DATA_REQUEST","needs":[{"tool":"get_month_report_full_data","args":{"period":"Март 2026"}}],"reason":"Нужен полный месячный отчёт"}

Запрос: "Сделай отчёт по прибыли артикулов за этот год"
Первый ответ:
{"type":"DATA_REQUEST","needs":[{"tool":"list_reports","args":{}}],"reason":"Нужно определить доступные месяцы """ + str(CURRENT_YEAR) + r""" года"}

Если после list_reports доступны Январь 2026, Февраль 2026, Март 2026, Апрель 2026, следующий ответ должен быть именно таким:
{"type":"DATA_REQUEST","needs":[
  {"tool":"get_month_report_full_data","args":{"period":"Январь 2026"}},
  {"tool":"get_month_report_full_data","args":{"period":"Февраль 2026"}},
  {"tool":"get_month_report_full_data","args":{"period":"Март 2026"}},
  {"tool":"get_month_report_full_data","args":{"period":"Апрель 2026"}}
],"reason":"Нужны полные месячные отчёты за доступные месяцы """ + str(CURRENT_YEAR) + r""" года"}

# Правила точности
- Не выдумывай периоды, месяцы и поля.
- Если tool вернул error, не придумывай причину.
- Если данных не хватает, запрашивай следующий tool.
- Если пользователь просит отчёт за год, а доступны только часть месяцев, строй отчёт по доступным месяцам и явно скажи это в FINAL_ANSWER.

# Стиль FINAL_ANSWER
- На русском.
- Коротко и по делу.
- Без раскрытия внутренней механики.
"""

def _extract_ollama_content(data: Dict[str, Any]) -> Optional[str]:
    """Пытается извлечь текст ответа из разных форматов payload Ollama."""
    parts: List[str] = []

    message = data.get("message")
    if isinstance(message, dict):
        for key in ("content", "text"):
            value = message.get(key)
            if isinstance(value, str) and value.strip():
                parts.append(value.strip())

    for key in ("response", "content", "text"):
        value = data.get(key)
        if isinstance(value, str) and value.strip():
            parts.append(value.strip())

    thinking = data.get("thinking")
    if isinstance(thinking, str) and thinking.strip():
        parts.append(thinking.strip())

    if not parts and isinstance(message, dict):
        tool_calls = message.get("tool_calls")
        if tool_calls:
            return json.dumps({"tool_calls": tool_calls}, ensure_ascii=False)

    if not parts:
        return None

    unique_parts: List[str] = []
    seen: set[str] = set()
    for part in parts:
        if part not in seen:
            unique_parts.append(part)
            seen.add(part)
    return "\n".join(unique_parts).strip() or None


def _extract_ollama_thinking(data: Dict[str, Any]) -> Optional[str]:
    """Извлекает reasoning/thinking из payload Ollama, если он доступен."""
    parts: List[str] = []
    message = data.get("message")
    if isinstance(message, dict):
        for key in ("thinking", "reasoning"):
            value = message.get(key)
            if isinstance(value, str) and value.strip():
                parts.append(value)

    for key in ("thinking", "reasoning"):
        value = data.get(key)
        if isinstance(value, str) and value.strip():
            parts.append(value)

    if not parts:
        return None

    unique_parts: List[str] = []
    seen: set[str] = set()
    for part in parts:
        if part not in seen:
            unique_parts.append(part)
            seen.add(part)
    return "\n".join(unique_parts) or None


def _extract_ollama_chunk(data: Dict[str, Any]) -> str:
    """Извлекает сырой текстовый chunk без strip, чтобы не терять пробелы при стриминге."""
    message = data.get("message")
    if isinstance(message, dict):
        for key in ("content", "text"):
            value = message.get(key)
            if isinstance(value, str):
                return value
    for key in ("response", "content", "text"):
        value = data.get(key)
        if isinstance(value, str):
            return value
    return ""


def _extract_ollama_thinking_chunk(data: Dict[str, Any]) -> str:
    """Извлекает сырой thinking chunk без strip, чтобы сохранить пробелы и переносы."""
    message = data.get("message")
    if isinstance(message, dict):
        for key in ("thinking", "reasoning"):
            value = message.get(key)
            if isinstance(value, str):
                return value
    for key in ("thinking", "reasoning"):
        value = data.get(key)
        if isinstance(value, str):
            return value
    return ""


def _send_ollama_chat(
    messages: List[Dict[str, str]],
    temperature: float,
    max_tokens: Optional[int],
) -> Tuple[Optional[str], Optional[Dict[str, Any]], Optional[Dict[str, Any]]]:
    """
    Отправляет запрос в локальный Ollama и возвращает текст ответа.
    """
    try:
        payload = {
            "model": OLLAMA_MODEL,
            "messages": messages,
            "stream": False,
            "options": {
                "temperature": temperature,
            },
        }
        if max_tokens is not None:
            payload["options"]["num_predict"] = max_tokens

        response = requests.post(
            OLLAMA_CHAT_URL,
            json=payload,
            timeout=120,
        )
        response.raise_for_status()
        data = response.json()

        usage_info = {
            "prompt_eval_count": data.get("prompt_eval_count"),
            "eval_count": data.get("eval_count"),
        }
        runtime_info = {
            "model": data.get("model", OLLAMA_MODEL),
            "done_reason": data.get("done_reason"),
            "host": OLLAMA_HOST,
            "raw_response": data,
        }
        content = _extract_ollama_content(data)
        if not content:
            return None, usage_info, runtime_info
        return content, usage_info, runtime_info
    except requests.exceptions.RequestException as e:
        print(f"⚠️ Ошибка при запросе к Ollama: {e}")
        return None, None, None
    except Exception as e:
        print(f"⚠️ Неожиданная ошибка при работе с Ollama: {e}")
        return None, None, None


def _stream_ollama_chat(
    messages: List[Dict[str, str]],
    temperature: float,
    max_tokens: Optional[int],
    on_chunk=None,
    on_thinking=None,
    should_cancel=None,
) -> Tuple[Optional[str], Optional[Dict[str, Any]], Optional[Dict[str, Any]]]:
    """
    Потоковый запрос к локальному Ollama. Возвращает полный текст ответа.
    """
    collected: List[str] = []
    thinking_collected: List[str] = []
    final_payload: Optional[Dict[str, Any]] = None
    try:
        payload = {
            "model": OLLAMA_MODEL,
            "messages": messages,
            "stream": True,
            "options": {
                "temperature": temperature,
            },
        }
        if max_tokens is not None:
            payload["options"]["num_predict"] = max_tokens

        with requests.post(
            OLLAMA_CHAT_URL,
            json=payload,
            timeout=120,
            stream=True,
        ) as response:
            response.raise_for_status()
            last_thinking = ""
            for raw_line in response.iter_lines(decode_unicode=True):
                if should_cancel and should_cancel():
                    return None, None, None
                if not raw_line:
                    continue
                data = json.loads(raw_line)
                final_payload = data
                chunk = _extract_ollama_chunk(data)
                thinking_chunk = _extract_ollama_thinking_chunk(data)
                if chunk:
                    collected.append(chunk)
                    if on_chunk:
                        on_chunk(chunk)
                if thinking_chunk:
                    delta = thinking_chunk
                    if last_thinking and thinking_chunk.startswith(last_thinking):
                        delta = thinking_chunk[len(last_thinking):]
                    last_thinking = thinking_chunk
                    if delta:
                        thinking_collected.append(delta)
                        if on_thinking:
                            on_thinking(delta)
                if data.get("done"):
                    break

        usage_info = {
            "prompt_eval_count": (final_payload or {}).get("prompt_eval_count"),
            "eval_count": (final_payload or {}).get("eval_count"),
        }
        runtime_info = {
            "model": (final_payload or {}).get("model", OLLAMA_MODEL),
            "done_reason": (final_payload or {}).get("done_reason"),
            "host": OLLAMA_HOST,
            "raw_response": final_payload,
            "thinking": "".join(thinking_collected).strip() or _extract_ollama_thinking(final_payload or {}),
        }
        content = "".join(collected).strip()
        if not content:
            content = _extract_ollama_content(final_payload or {}) or ""
        if not content:
            return None, usage_info, runtime_info
        return content, usage_info, runtime_info
    except requests.exceptions.RequestException as e:
        print(f"⚠️ Ошибка при потоковом запросе к Ollama: {e}")
        return None, None, None
    except Exception as e:
        print(f"⚠️ Неожиданная ошибка при потоковом запросе к Ollama: {e}")
        return None, None, None


def check_ollama_model() -> Tuple[bool, str]:
    """
    Проверяет доступность Ollama и наличие выбранной модели.
    """
    try:
        response = requests.get(f"{OLLAMA_HOST}/api/tags", timeout=5)
        response.raise_for_status()
        data = response.json()
    except requests.exceptions.RequestException:
        return False, f"Ollama недоступен по адресу {OLLAMA_HOST}"

    models = data.get("models", []) or []
    available = {
        item.get("name")
        for item in models
        if isinstance(item, dict) and item.get("name")
    }
    if OLLAMA_MODEL not in available:
        return False, (
            f"Модель {OLLAMA_MODEL} не найдена в Ollama. "
            f"Скачайте её командой: ollama pull {OLLAMA_MODEL}"
        )

    return True, ""


def load_costs_data(repo_root: Path) -> Optional[Dict[str, Any]]:
    """
    Загружает все данные из costs.xlsx.
    
    Returns:
        Словарь с данными из costs.xlsx или None в случае ошибки
    """
    costs_path = repo_root / "costs.xlsx"
    if not costs_path.exists():
        return None
    
    try:
        df = read_costs_dataframe(costs_path)
        
        # Преобразуем DataFrame в словарь для удобной передачи ИИ
        costs_data = {
            "Всего товаров": len(df),
            "Колонки": list(df.columns),
            "Данные": []
        }
        
        # Преобразуем каждую строку в словарь
        for _, row in df.iterrows():
            row_dict = {}
            for col in df.columns:
                value = row[col]
                # Преобразуем NaN в None
                if pd.isna(value):
                    value = None
                # Преобразуем числа в float для JSON-совместимости
                elif isinstance(value, (int, float)):
                    value = float(value)
                else:
                    value = str(value)
                row_dict[col] = value
            costs_data["Данные"].append(row_dict)
        
        return costs_data
        
    except Exception as e:
        return None


def _normalize_column_name(value: Any) -> str:
    """Приводит имя колонки к устойчивому виду для поиска по меняющимся отчётам."""
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ")
    return re.sub(r"\s+", " ", text).strip()


def _normalize_column_key(value: Any) -> str:
    return _normalize_column_name(value).lower()


def _to_json_compatible(value: Any) -> Any:
    if pd.isna(value):
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value)
    return str(value)


def _build_column_lookup(df: pd.DataFrame) -> Dict[str, str]:
    lookup: Dict[str, str] = {}
    for col in df.columns:
        normalized = _normalize_column_key(col)
        if normalized and normalized not in lookup:
            lookup[normalized] = col
    return lookup


def _find_column(column_lookup: Dict[str, str], *candidates: str, contains: bool = False) -> Optional[str]:
    normalized_candidates = [_normalize_column_key(candidate) for candidate in candidates if candidate]
    if not normalized_candidates:
        return None

    if not contains:
        for candidate in normalized_candidates:
            if candidate in column_lookup:
                return column_lookup[candidate]

    for normalized_name, original_name in column_lookup.items():
        if any(candidate and candidate in normalized_name for candidate in normalized_candidates):
            return original_name
    return None


def _serialize_dataframe_rows(df: pd.DataFrame) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for _, row in df.iterrows():
        row_dict: Dict[str, Any] = {}
        for col in df.columns:
            row_dict[_normalize_column_name(col)] = _to_json_compatible(row[col])
        rows.append(row_dict)
    return rows


def _to_float_or_none(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, str):
        text = value.strip().replace(",", ".")
        if not text or text == "-":
            return None
        value = text
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        return None
    if pd.isna(numeric):
        return None
    return float(numeric)


def _normalize_status_name(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def _normalize_artikul_identifier(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


def _extract_year_from_text(text: str) -> Optional[int]:
    return extract_year_from_text_core(text, CURRENT_YEAR)


def _is_yearly_artikul_report_request(user_message: Optional[str], results: Dict[str, Any]) -> bool:
    return is_yearly_artikul_report_request_core(user_message, results, CURRENT_YEAR)


def detect_workflow(user_message: Optional[str]) -> WorkflowState:
    return detect_workflow_core(user_message, CURRENT_YEAR)


def _extract_available_year_periods_from_list_reports(
    results: Dict[str, Any],
    year: Optional[int],
) -> List[str]:
    list_reports_result = results.get("list_reports")
    if not isinstance(list_reports_result, dict):
        return []
    reports = list_reports_result.get("reports")
    if not isinstance(reports, list):
        return []

    periods_with_sort_keys: List[Tuple[Tuple[int, int], str]] = []
    for item in reports:
        if not isinstance(item, dict):
            continue
        period = str(item.get("period") or "").strip()
        if not period:
            continue
        period_year = _extract_year_from_text(period)
        if year is not None and period_year != year:
            continue

        month_number = 0
        for idx, month_name in enumerate(MONTHS_RU, start=1):
            if period.lower().startswith(month_name.lower()):
                month_number = idx
                break
        if period_year is None:
            period_year = year or 0
        periods_with_sort_keys.append(((period_year, month_number), period))

    periods_with_sort_keys.sort()
    return [period for _, period in periods_with_sort_keys]


def build_workflow_state(user_message: Optional[str], results: Dict[str, Any]) -> WorkflowState:
    return build_workflow_state_core(user_message, results, CURRENT_YEAR, MONTHS_RU)


def get_workflow_followup_needs(user_message: Optional[str], results: Dict[str, Any]) -> List[Dict[str, Any]]:
    return get_workflow_followup_needs_core(user_message, results, CURRENT_YEAR, MONTHS_RU)


def build_yearly_artikul_profit_summary(
    results: Dict[str, Any],
    user_message: Optional[str] = None,
) -> Optional[Dict[str, Any]]:
    return build_yearly_artikul_profit_summary_core(results, user_message, CURRENT_YEAR)


def build_workflow_fallback_answer(user_message: Optional[str], results: Dict[str, Any]) -> Optional[str]:
    return build_workflow_fallback_answer_core(user_message, results, CURRENT_YEAR)


def load_monthly_report_detailed_data(report_path: Path) -> Optional[Dict[str, Any]]:
    """
    Загружает детальные данные из месячного отчёта (лист "Заказы").
    Агрегирует данные по артикулам для экономии токенов.
    
    Returns:
        Словарь с агрегированными данными или None в случае ошибки
    """
    if not report_path.exists():
        return {"error": f"Файл отчёта не найден: {report_path.name}"}

    try:
        df = pd.read_excel(report_path, sheet_name=ORDER_SHEET)
    except ValueError:
        return {"error": f"Лист '{ORDER_SHEET}' не найден в отчёте {report_path.name}"}
    except Exception as e:
        return {"error": f"Ошибка чтения отчёта {report_path.name}: {e}"}

    if df.empty:
        return {"error": f"Лист '{ORDER_SHEET}' пуст в отчёте {report_path.name}"}

    df.columns = [_normalize_column_name(col) for col in df.columns]
    summary_data = load_report_summary(report_path) or {}
    summary_metric_names = {
        key for key in summary_data.keys()
        if key and key != "Период отчёта"
    }

    row_columns = [
        col for col in df.columns
        if col
        and not _normalize_column_key(col).startswith("unnamed:")
        and col not in summary_metric_names
        and not re.fullmatch(r"-?\d+(?:\.\d+)?", str(col))
    ]
    work_df = df[row_columns].copy() if row_columns else df.copy()
    column_lookup = _build_column_lookup(work_df)

    artikul_col = _find_column(column_lookup, "Артикул", "offer_id", "offer id", contains=True)
    status_col = _find_column(column_lookup, "Статус", contains=True)
    schema_col = _find_column(column_lookup, "Схема", contains=True)
    profit_col = _find_column(column_lookup, "Прибыль", contains=True)

    numeric_aliases = [
        "Количество шт.",
        "Цена продажи",
        "Комиссия за продажу Ozon",
        "Логистика (Включает операционные ошибки продавца)",
        "Сумма начисления",
        "Себестоимость",
        "Прибыль",
    ]
    resolved_numeric_cols: List[str] = []
    for alias in numeric_aliases:
        col = _find_column(column_lookup, alias, contains=True)
        if col and col not in resolved_numeric_cols:
            resolved_numeric_cols.append(col)

    report_name = report_path.stem
    detailed_data: Dict[str, Any] = {
        "Период отчёта": report_name,
        "Всего строк отчёта": int(len(work_df)),
        "Колонки": list(work_df.columns),
        "Сводка отчёта": summary_data,
        "Строки отчёта": _serialize_dataframe_rows(work_df),
        "Диагностика колонок": {
            "Артикул": artikul_col,
            "Статус": status_col,
            "Схема": schema_col,
            "Прибыль": profit_col,
            "Числовые колонки": resolved_numeric_cols,
            "Отфильтрованные служебные колонки": [col for col in df.columns if col not in work_df.columns],
        },
    }

    if artikul_col:
        artikul_stats = []
        for artikul in work_df[artikul_col].dropna().unique():
            artikul_df = work_df[work_df[artikul_col] == artikul]
            stats: Dict[str, Any] = {
                "Артикул": str(artikul),
                "Количество заказов": int(len(artikul_df)),
                "Строки отчёта": _serialize_dataframe_rows(artikul_df),
            }

            for col in resolved_numeric_cols:
                numeric_series = pd.to_numeric(artikul_df[col], errors="coerce")
                total = numeric_series.sum()
                avg = numeric_series.mean()
                if pd.notna(total):
                    stats[f"Сумма {col}"] = float(total)
                if pd.notna(avg):
                    stats[f"Средняя {col}"] = float(avg)

            if status_col:
                status_counts = artikul_df[status_col].value_counts(dropna=False).to_dict()
                stats["Распределение по статусам"] = {str(k): int(v) for k, v in status_counts.items()}

            if schema_col:
                schema_counts = artikul_df[schema_col].value_counts(dropna=False).to_dict()
                stats["Распределение по схемам"] = {str(k): int(v) for k, v in schema_counts.items()}

            artikul_stats.append(stats)

        detailed_data["Статистика по артикулам"] = artikul_stats

    if status_col:
        status_counts = work_df[status_col].value_counts(dropna=False).to_dict()
        detailed_data["Общее распределение по статусам"] = {str(k): int(v) for k, v in status_counts.items()}

    if schema_col:
        schema_counts = work_df[schema_col].value_counts(dropna=False).to_dict()
        detailed_data["Общее распределение по схемам"] = {str(k): int(v) for k, v in schema_counts.items()}

    if profit_col and artikul_col:
        profit_df = work_df[[artikul_col, profit_col]].copy()
        profit_df[profit_col] = pd.to_numeric(profit_df[profit_col], errors="coerce")
        top_profit = profit_df.groupby(artikul_col)[profit_col].sum().nlargest(10)
        detailed_data["Топ-10 артикулов по прибыли"] = {
            str(artikul): float(profit) for artikul, profit in top_profit.items() if pd.notna(profit)
        }

    if artikul_col:
        top_orders = work_df[artikul_col].value_counts().head(10)
        detailed_data["Топ-10 артикулов по количеству заказов"] = {
            str(artikul): int(count) for artikul, count in top_orders.items()
        }

    return detailed_data


def load_report_summary(report_path: Path) -> Optional[Dict[str, Any]]:
    """
    Загружает сводку из месячного отчёта.
    Читает итоговые показатели из колонок P и Q листа "Заказы".
    Блок метрик определяется динамически: берём все подряд идущие
    бизнес-показатели сверху листа, пока не встретим длинную серию пустых строк.
    
    Returns:
        Словарь с метриками отчёта или None в случае ошибки
    """
    with ExitStack() as books:
        if not report_path.exists():
            return None
    
        try:
            wb = books.enter_context(closing(load_workbook(report_path, data_only=True)))
            if ORDER_SHEET not in wb.sheetnames:
                return None
        
            ws = wb[ORDER_SHEET]
        
            summary = {}

            def normalize_value(value: Any) -> Any:
                try:
                    if isinstance(value, (int, float)):
                        return float(value)
                    if isinstance(value, str):
                        cleaned = value.replace(",", ".").replace(" ", "")
                        return float(cleaned)
                except (ValueError, TypeError):
                    pass
                return value

            started_metrics = False
            empty_streak = 0
            max_scan_rows = 200
            max_empty_streak = 10

            for row_num in range(1, max_scan_rows + 1):
                metric_name = ws[f"P{row_num}"].value
                value = ws[f"Q{row_num}"].value

                metric_name = str(metric_name).strip() if metric_name is not None else ""

                if metric_name:
                    started_metrics = True
                    empty_streak = 0
                    summary[metric_name] = normalize_value(value)
                    continue

                if value is not None:
                    # Если у строки почему-то нет подписи, но есть значение,
                    # сохраняем его с техническим именем, чтобы не потерять показатель.
                    started_metrics = True
                    empty_streak = 0
                    summary[f"Показатель P{row_num}"] = normalize_value(value)
                    continue

                if started_metrics:
                    empty_streak += 1
                    if empty_streak >= max_empty_streak:
                        break
        
            # Добавляем информацию о периоде отчёта из имени файла
            report_name = report_path.stem  # Без расширения
            summary["Период отчёта"] = report_name
        
            return summary if summary else None
        
        except Exception as e:
            return None


def list_available_reports(repo_root: Path) -> List[Path]:
    """
    Возвращает список доступных месячных отчётов.
    
    Returns:
        Список путей к файлам отчётов, отсортированный по дате (новые первыми)
    """
    reports_dir = repo_root / REPORTS_DIR_NAME
    if not reports_dir.exists():
        return []
    
    reports = []
    for file_path in reports_dir.glob("*.xlsx"):
        if file_path.is_file() and not file_path.name.startswith(("~$", "~tmp_")):
            reports.append(file_path)
    
    # Сортируем по дате изменения (новые первыми)
    reports.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    
    return reports


def list_abc_xyz_reports(repo_root: Path) -> List[Path]:
    """
    Возвращает список доступных ABC&XYZ отчётов.
    
    Returns:
        Список путей к файлам ABC&XYZ отчётов, отсортированный по дате (новые первыми)
    """
    abc_xyz_dir = repo_root / "ABC&XYZ reports"
    if not abc_xyz_dir.exists():
        return []
    
    reports = []
    for file_path in abc_xyz_dir.glob("*.xlsx"):
        if file_path.is_file() and not file_path.name.startswith(("~$", "~tmp_")):
            reports.append(file_path)
    
    # Сортируем по дате изменения (новые первыми)
    reports.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    
    return reports


def load_abc_xyz_summary(report_path: Path) -> Optional[Dict[str, Any]]:
    """
    Загружает сводку из ABC&XYZ отчёта.
    Читает данные из листов ABC, XYZ и Итог.
    
    Returns:
        Словарь с метриками отчёта или None в случае ошибки
    """
    with ExitStack() as books:
        if not report_path.exists():
            return None
    
        try:
            wb = books.enter_context(closing(load_workbook(report_path, data_only=True)))
            summary = {}
        
            # Добавляем информацию о периоде отчёта из имени файла
            report_name = report_path.stem  # Без расширения
            summary["Период отчёта"] = report_name
            summary["Тип отчёта"] = "ABC&XYZ"
        
            # Читаем лист "Итог" для получения общей статистики
            if "Итог" in wb.sheetnames:
                ws = wb["Итог"]
                # Подсчитываем количество артикулов в каждой категории ABCXYZ
                abcxyz_counts = {}
                total_articles = 0
            
                # Ищем колонку с общей оценкой ABCXYZ (обычно последняя или предпоследняя)
                # Пробуем найти заголовок
                header_row = 1
                abcxyz_col = None
            
                for col_idx in range(1, ws.max_column + 1):
                    cell_value = ws.cell(header_row, col_idx).value
                    if cell_value and ("ABCXYZ" in str(cell_value) or "Оценка" in str(cell_value)):
                        abcxyz_col = col_idx
                        break
            
                # Если не нашли по заголовку, пробуем последнюю колонку
                if abcxyz_col is None:
                    abcxyz_col = ws.max_column
            
                # Подсчитываем категории
                for row in range(2, ws.max_row + 1):
                    cell_value = ws.cell(row, abcxyz_col).value
                    if cell_value:
                        category = str(cell_value).strip()
                        if category and category != "Недостаточно данных":
                            abcxyz_counts[category] = abcxyz_counts.get(category, 0) + 1
                            total_articles += 1
            
                summary["Всего артикулов"] = total_articles
                summary["Распределение по категориям ABCXYZ"] = abcxyz_counts
        
            # Читаем лист "ABC" для статистики по прибыли
            if "ABC" in wb.sheetnames:
                try:
                    df_abc = pd.read_excel(report_path, sheet_name="ABC")
                    if not df_abc.empty:
                        # Ищем колонку с прибылью
                        profit_col = None
                        for col in df_abc.columns:
                            if "прибыль" in str(col).lower() or "profit" in str(col).lower():
                                profit_col = col
                                break
                    
                        if profit_col is not None:
                            total_profit = df_abc[profit_col].sum()
                            summary["Общая прибыль (ABC)"] = float(total_profit)
                        
                            # Подсчитываем артикулы по категориям A, B, C
                            if "ABC" in df_abc.columns:
                                abc_dist = df_abc["ABC"].value_counts().to_dict()
                                summary["Распределение ABC"] = {str(k): int(v) for k, v in abc_dist.items()}
                except Exception:
                    pass
        
            # Читаем лист "XYZ" для статистики по стабильности
            if "XYZ" in wb.sheetnames:
                try:
                    df_xyz = pd.read_excel(report_path, sheet_name="XYZ")
                    if not df_xyz.empty:
                        # Подсчитываем артикулы по категориям X, Y1, Y2, Y3, Y, Z
                        if "XYZ" in df_xyz.columns:
                            xyz_dist = df_xyz["XYZ"].value_counts().to_dict()
                            summary["Распределение XYZ"] = {str(k): int(v) for k, v in xyz_dist.items()}
                except Exception:
                    pass
        
            return summary if summary else None
        
        except Exception as e:
            return None


def format_costs_data_for_ai(costs_data: Dict[str, Any]) -> str:
    """
    Форматирует данные из costs.xlsx для передачи ИИ.
    
    Args:
        costs_data: Словарь с данными из costs.xlsx
    
    Returns:
        Отформатированная строка с данными
    """
    if not costs_data:
        return ""
    
    lines = ["\n" + "="*60]
    lines.append("## ДАННЫЕ ИЗ ФАЙЛА СЕБЕСТОИМОСТИ (costs.xlsx)")
    lines.append("="*60)
    lines.append(f"\nВсего товаров: {costs_data.get('Всего товаров', 0)}")
    lines.append(f"Колонки: {', '.join(costs_data.get('Колонки', []))}\n")
    
    # Показываем первые 20 товаров (чтобы не превысить лимиты токенов)
    # Остальные данные доступны для анализа по запросу
    data_items = costs_data.get("Данные", [])
    show_count = min(20, len(data_items))
    
    if show_count > 0:
        lines.append(f"### Данные по товарам (показано {show_count} из {len(data_items)}):\n")
        
        for idx, item in enumerate(data_items[:show_count], 1):
            lines.append(f"**Товар {idx}:**")
            for key, value in item.items():
                if value is not None:
                    if isinstance(value, float):
                        if value >= 1000:
                            lines.append(f"  - {key}: {value:,.0f}")
                        else:
                            lines.append(f"  - {key}: {value:.2f}")
                    else:
                        lines.append(f"  - {key}: {value}")
            lines.append("")
        
        if len(data_items) > show_count:
            lines.append(f"\n*... и ещё {len(data_items) - show_count} товаров (все данные доступны для анализа, просто укажи конкретный артикул)*\n")
    
    return "\n".join(lines)


def format_detailed_report_data_for_ai(detailed_data: Dict[str, Any]) -> str:
    """
    Форматирует детальные данные из месячного отчёта для передачи ИИ.
    
    Args:
        detailed_data: Словарь с детальными данными отчёта
    
    Returns:
        Отформатированная строка с данными
    """
    if not detailed_data:
        return ""
    
    lines = [f"\n### ДЕТАЛЬНЫЕ ДАННЫЕ: {detailed_data.get('Период отчёта', 'Неизвестный период')}\n"]
    
    if "Всего заказов" in detailed_data:
        lines.append(f"- Всего заказов: {detailed_data['Всего заказов']}")
    
    if "Общее распределение по статусам" in detailed_data:
        lines.append("\n**Распределение заказов по статусам:**")
        for status, count in detailed_data["Общее распределение по статусам"].items():
            lines.append(f"  - {status}: {count}")
    
    if "Общее распределение по схемам" in detailed_data:
        lines.append("\n**Распределение заказов по схемам:**")
        for schema, count in detailed_data["Общее распределение по схемам"].items():
            lines.append(f"  - {schema}: {count}")
    
    if "Топ-10 артикулов по прибыли" in detailed_data:
        lines.append("\n**Топ-10 артикулов по прибыли:**")
        for artikul, profit in detailed_data["Топ-10 артикулов по прибыли"].items():
            lines.append(f"  - {artikul}: {profit:,.0f} руб.")
    
    if "Топ-10 артикулов по количеству заказов" in detailed_data:
        lines.append("\n**Топ-10 артикулов по количеству заказов:**")
        for artikul, count in detailed_data["Топ-10 артикулов по количеству заказов"].items():
            lines.append(f"  - {artikul}: {count} заказов")
    
    # Статистика по артикулам (показываем первые 10 для экономии токенов)
    if "Статистика по артикулам" in detailed_data:
        artikul_stats = detailed_data["Статистика по артикулам"]
        show_count = min(10, len(artikul_stats))
        
        lines.append(f"\n**Статистика по артикулам (показано {show_count} из {len(artikul_stats)}):**")
        for stats in artikul_stats[:show_count]:
            lines.append(f"\n- **Артикул {stats.get('Артикул', 'N/A')}:**")
            lines.append(f"  - Заказов: {stats.get('Количество заказов', 0)}")
            
            # Показываем только ключевые метрики (суммы, а не средние, чтобы экономить токены)
            key_metrics = ["Сумма Прибыль", "Сумма Цена продажи", "Сумма Количество шт.", 
                          "Распределение по статусам", "Распределение по схемам"]
            for key in key_metrics:
                if key in stats:
                    value = stats[key]
                    if isinstance(value, dict):
                        lines.append(f"  - {key}: {value}")
                    elif isinstance(value, float):
                        if value >= 1000:
                            lines.append(f"  - {key}: {value:,.0f}")
                        else:
                            lines.append(f"  - {key}: {value:.2f}")
                    else:
                        lines.append(f"  - {key}: {value}")
        
        if len(artikul_stats) > show_count:
            lines.append(f"\n*... и ещё {len(artikul_stats) - show_count} артикулов (все данные доступны для анализа, просто укажи конкретный артикул или период)*")
    
    return "\n".join(lines)


def format_all_reports_summary_for_ai(
    monthly_summaries: List[Dict[str, Any]], 
    abc_xyz_summaries: List[Dict[str, Any]],
    costs_data: Optional[Dict[str, Any]] = None,
    monthly_detailed_data: List[Dict[str, Any]] = None
) -> str:
    """
    Форматирует сводки всех отчётов в текстовый формат для передачи ИИ.
    
    Args:
        monthly_summaries: Список сводок месячных отчётов
        abc_xyz_summaries: Список сводок ABC&XYZ отчётов
        costs_data: Данные из costs.xlsx
        monthly_detailed_data: Список детальных данных из месячных отчётов
    
    Returns:
        Отформатированная строка со всеми сводками
    """
    lines = []
    
    # Данные из costs.xlsx
    if costs_data:
        lines.append(format_costs_data_for_ai(costs_data))
        lines.append("")
    
    # Месячные отчёты - сводки
    if monthly_summaries:
        lines.append("\n" + "="*60)
        lines.append("## МЕСЯЧНЫЕ ОТЧЁТЫ ПО ПРОДАЖАМ (СВОДКИ)")
        lines.append("="*60)
        
        for summary in monthly_summaries:
            lines.append(format_report_summary_for_ai(summary))
            lines.append("")  # Пустая строка между отчётами
    
    # Месячные отчёты - детальные данные
    if monthly_detailed_data:
        lines.append("\n" + "="*60)
        lines.append("## МЕСЯЧНЫЕ ОТЧЁТЫ ПО ПРОДАЖАМ (ДЕТАЛЬНЫЕ ДАННЫЕ)")
        lines.append("="*60)
        
        for detailed in monthly_detailed_data:
            lines.append(format_detailed_report_data_for_ai(detailed))
            lines.append("")  # Пустая строка между отчётами
    
    # ABC&XYZ отчёты
    if abc_xyz_summaries:
        lines.append("\n" + "="*60)
        lines.append("## ABC&XYZ ОТЧЁТЫ (АНАЛИЗ ПРИБЫЛЬНОСТИ И СТАБИЛЬНОСТИ)")
        lines.append("="*60)
        
        for summary in abc_xyz_summaries:
            lines.append(format_abc_xyz_summary_for_ai(summary))
            lines.append("")  # Пустая строка между отчётами
    
    return "\n".join(lines)


def format_abc_xyz_summary_for_ai(summary: Dict[str, Any]) -> str:
    """
    Форматирует сводку ABC&XYZ отчёта в текстовый формат для передачи ИИ.
    
    Args:
        summary: Словарь с метриками ABC&XYZ отчёта
    
    Returns:
        Отформатированная строка со сводкой
    """
    if not summary:
        return ""
    
    lines = [f"\n### ABC&XYZ ОТЧЁТ: {summary.get('Период отчёта', 'Неизвестный период')}\n"]
    
    if "Всего артикулов" in summary:
        lines.append(f"- Всего артикулов: {summary['Всего артикулов']}")
    
    if "Распределение ABC" in summary:
        lines.append("\n**Распределение по прибыльности (ABC):**")
        for category, count in summary["Распределение ABC"].items():
            lines.append(f"  - Категория {category}: {count} артикулов")
    
    if "Распределение XYZ" in summary:
        lines.append("\n**Распределение по стабильности спроса (XYZ):**")
        for category, count in summary["Распределение XYZ"].items():
            lines.append(f"  - Категория {category}: {count} артикулов")
    
    if "Распределение по категориям ABCXYZ" in summary:
        lines.append("\n**Комбинированное распределение (ABCXYZ):**")
        for category, count in sorted(summary["Распределение по категориям ABCXYZ"].items()):
            lines.append(f"  - {category}: {count} артикулов")
    
    if "Общая прибыль (ABC)" in summary:
        profit = summary["Общая прибыль (ABC)"]
        if profit >= 1000:
            lines.append(f"\n- Общая прибыль за период: {profit:,.0f} руб.")
        else:
            lines.append(f"\n- Общая прибыль за период: {profit:.2f} руб.")
    
    return "\n".join(lines)


def format_report_summary_for_ai(summary: Dict[str, Any]) -> str:
    """
    Форматирует сводку отчёта в текстовый формат для передачи ИИ.
    
    Args:
        summary: Словарь с метриками отчёта
    
    Returns:
        Отформатированная строка со сводкой
    """
    if not summary:
        return ""
    
    lines = [f"\n## ДАННЫЕ ИЗ МЕСЯЧНОГО ОТЧЁТА: {summary.get('Период отчёта', 'Неизвестный период')}\n"]
    
    # Группируем метрики по категориям
    financial_metrics = [
        "Общая выручка",
        "Чистая прибыль",
        "Итоговая себестоимость",
        "COGS (валовая прибыль)",
        "Операционные расходы",
    ]
    
    margin_metrics = [
        "Рентабельность по чистой прибыли (Net Profit Margin) %",
        "Gross Profit Margin Рентабельность по валовой прибыли %",
    ]
    
    order_metrics = [
        "Общее количество заказов",
        "Количество доставленных заказов",
        "Количество отменённых заказов",
        "Средний чек",
    ]
    
    cost_metrics = [
        "Продвижение Ozon",
        "Звёздные товары",
        "Внешний маркетинг",
        "Комиссии Ozon %",
        "Логистика %",
    ]
    
    def format_value(key: str, value: Any) -> str:
        """Форматирует значение метрики для вывода."""
        if isinstance(value, (int, float)):
            if "процент" in key.lower() or "%" in key:
                return f"{value:.2f}%"
            elif value >= 1000:
                return f"{value:,.0f} руб."
            else:
                return f"{value:.2f} руб."
        return str(value)
    
    if any(k in summary for k in financial_metrics):
        lines.append("### Финансовые показатели:")
        for metric in financial_metrics:
            if metric in summary:
                lines.append(f"- {metric}: {format_value(metric, summary[metric])}")
    
    if any(k in summary for k in margin_metrics):
        lines.append("\n### Рентабельность:")
        for metric in margin_metrics:
            if metric in summary:
                lines.append(f"- {metric}: {format_value(metric, summary[metric])}")
    
    if any(k in summary for k in order_metrics):
        lines.append("\n### Статистика заказов:")
        for metric in order_metrics:
            if metric in summary:
                lines.append(f"- {metric}: {format_value(metric, summary[metric])}")
    
    if any(k in summary for k in cost_metrics):
        lines.append("\n### Расходы и комиссии:")
        for metric in cost_metrics:
            if metric in summary:
                lines.append(f"- {metric}: {format_value(metric, summary[metric])}")
    
    return "\n".join(lines)


# ============================================================================
# ИНСТРУМЕНТЫ (TOOLS) - функции для доступа к данным по запросу
# ============================================================================

class Tools:
    """Класс с инструментами для доступа к данным."""
    
    def __init__(self, repo_root: Path):
        self.repo_root = repo_root
        self._costs_df = None  # Кеш для costs.xlsx
        self._costs_columns = None
        self._costs_signature = None
    
    def _get_costs_df(self) -> Optional[pd.DataFrame]:
        """Загружает costs.xlsx с кешированием."""
        costs_path = self.repo_root / "costs.xlsx"
        try:
            stat = costs_path.stat()
            signature = (stat.st_mtime_ns, stat.st_size)
            if self._costs_df is None or signature != self._costs_signature:
                self._costs_df = read_costs_dataframe(costs_path)
                self._costs_signature = signature
        except Exception:
            self._costs_df = None
            self._costs_signature = None
            return None
        return self._costs_df
    
    def _normalize_period(self, period: str) -> Optional[Path]:
        """Нормализует период и возвращает путь к файлу отчёта."""
        # Пробуем разные форматы: "2025-12", "Декабрь 2025", "12 2025"
        reports = list_available_reports(self.repo_root)
        
        # Формат "2025-12"
        if re.match(r'^\d{4}-\d{2}$', period):
            year, month = period.split('-')
            month_num = int(month)
            if 1 <= month_num <= 12:
                name = f"{MONTHS_RU[month_num - 1]} {year}.xlsx"
                for r in reports:
                    if r.name == name:
                        return r
        
        # Формат "Декабрь 2025" или частичное совпадение
        period_lower = period.lower()
        for r in reports:
            if period_lower in r.stem.lower() or r.stem.lower() in period_lower:
                return r
        
        return None
    
    def list_reports(self, args: Dict[str, Any]) -> Dict[str, Any]:
        """Возвращает список доступных месячных отчётов."""
        reports = list_available_reports(self.repo_root)
        return {
            "reports": [
                {"period": r.stem, "file": r.name, "path": str(r)}
                for r in reports
            ],
            "count": len(reports)
        }
    
    def list_abcxyz_reports(self, args: Dict[str, Any]) -> Dict[str, Any]:
        """Возвращает список доступных ABC&XYZ отчётов."""
        reports = list_abc_xyz_reports(self.repo_root)
        return {
            "reports": [
                {"period": r.stem, "file": r.name, "path": str(r)}
                for r in reports
            ],
            "count": len(reports)
        }
    
    def get_month_summary(self, args: Dict[str, Any]) -> Dict[str, Any]:
        """Возвращает сводку за месяц."""
        period = args.get("period", "")
        if not period:
            return {"error": "Параметр period обязателен"}
        
        report_path = self._normalize_period(period)
        if not report_path:
            return {"error": f"Отчёт за период '{period}' не найден"}
        
        summary = load_report_summary(report_path)
        if summary:
            return summary
        return {"error": "Не удалось загрузить сводку"}

    def get_month_report_full_data(self, args: Dict[str, Any]) -> Dict[str, Any]:
        """Возвращает полный набор данных месячного отчёта."""
        period = args.get("period", "")
        if not period:
            return {"error": "Параметр period обязателен"}

        report_path = self._normalize_period(period)
        if not report_path:
            return {"error": f"Отчёт за период '{period}' не найден"}

        detailed = load_monthly_report_detailed_data(report_path)
        if not detailed:
            return {"error": "Не удалось загрузить данные отчёта"}
        return detailed
    
    def get_top_profit(self, args: Dict[str, Any]) -> Dict[str, Any]:
        """Возвращает топ-N артикулов по прибыли."""
        period = args.get("period", "")
        n = args.get("n", 10)
        
        if not period:
            return {"error": "Параметр period обязателен"}
        
        report_path = self._normalize_period(period)
        if not report_path:
            return {"error": f"Отчёт за период '{period}' не найден"}
        
        detailed = load_monthly_report_detailed_data(report_path)
        if not detailed:
            return {"error": "Не удалось загрузить данные"}
        if detailed.get("error"):
            return detailed
        
        top_profit = detailed.get("Топ-10 артикулов по прибыли", {})
        # Ограничиваем до n
        sorted_items = sorted(top_profit.items(), key=lambda x: x[1], reverse=True)[:n]
        return {"top_profit": dict(sorted_items), "period": period}
    
    def get_top_orders(self, args: Dict[str, Any]) -> Dict[str, Any]:
        """Возвращает топ-N артикулов по количеству заказов."""
        period = args.get("period", "")
        n = args.get("n", 10)
        
        if not period:
            return {"error": "Параметр period обязателен"}
        
        report_path = self._normalize_period(period)
        if not report_path:
            return {"error": f"Отчёт за период '{period}' не найден"}
        
        detailed = load_monthly_report_detailed_data(report_path)
        if not detailed:
            return {"error": "Не удалось загрузить данные"}
        if detailed.get("error"):
            return detailed
        
        top_orders = detailed.get("Топ-10 артикулов по количеству заказов", {})
        # Ограничиваем до n
        sorted_items = sorted(top_orders.items(), key=lambda x: x[1], reverse=True)[:n]
        return {"top_orders": dict(sorted_items), "period": period}
    
    def get_artikul_stats(self, args: Dict[str, Any]) -> Dict[str, Any]:
        """Возвращает статистику по артикулу за период."""
        period = args.get("period", "")
        artikul = args.get("artikul", "")
        
        if not period or not artikul:
            return {"error": "Параметры period и artikul обязательны"}
        
        report_path = self._normalize_period(period)
        if not report_path:
            return {"error": f"Отчёт за период '{period}' не найден"}
        
        detailed = load_monthly_report_detailed_data(report_path)
        if not detailed:
            return {"error": "Не удалось загрузить данные"}
        if detailed.get("error"):
            return detailed
        
        artikul_stats = detailed.get("Статистика по артикулам", [])
        for stats in artikul_stats:
            if str(stats.get("Артикул", "")) == str(artikul):
                return {"artikul": artikul, "period": period, "stats": stats}
        
        return {"error": f"Артикул '{artikul}' не найден в отчёте за '{period}'"}
    
    def search_artikul(self, args: Dict[str, Any]) -> Dict[str, Any]:
        """Поиск артикула по запросу."""
        query = args.get("query", "").lower()
        if not query:
            return {"error": "Параметр query обязателен"}
        
        df = self._get_costs_df()
        if df is None:
            return {"error": "Файл costs.xlsx не найден"}
        
        # Ищем в колонке артикула
        results = []
        for col in df.columns:
            if "артикул" in col.lower() or "artikul" in col.lower() or "offer_id" in col.lower():
                matches = df[df[col].astype(str).str.lower().str.contains(query, na=False)]
                for _, row in matches.iterrows():
                    results.append({col: str(row[col]) for col in df.columns})
                break
        
        return {"results": results[:20], "count": len(results)}  # Ограничиваем 20 результатами
    
    def get_costs_columns(self, args: Dict[str, Any]) -> Dict[str, Any]:
        """Возвращает список колонок в costs.xlsx."""
        df = self._get_costs_df()
        if df is None:
            return {"error": "Файл costs.xlsx не найден"}
        return {"columns": list(df.columns), "count": len(df.columns)}
    
    def get_artikul_cost(self, args: Dict[str, Any]) -> Dict[str, Any]:
        """Возвращает данные по артикулу из costs.xlsx."""
        artikul = args.get("artikul", "")
        if not artikul:
            return {"error": "Параметр artikul обязателен"}
        
        df = self._get_costs_df()
        if df is None:
            return {"error": "Файл costs.xlsx не найден"}
        
        # Ищем артикул
        for col in df.columns:
            if "артикул" in col.lower() or "artikul" in col.lower() or "offer_id" in col.lower():
                matches = df[df[col].astype(str) == str(artikul)]
                if not matches.empty:
                    row = matches.iloc[0]
                    return {col: str(row[col]) if pd.notna(row[col]) else None for col in df.columns}
                break
        
        return {"error": f"Артикул '{artikul}' не найден в costs.xlsx"}
    
    def get_costs_bulk(self, args: Dict[str, Any]) -> Dict[str, Any]:
        """Возвращает данные по нескольким артикулам."""
        artikuls = args.get("artikuls", [])
        if not artikuls:
            return {"error": "Параметр artikuls обязателен"}
        
        df = self._get_costs_df()
        if df is None:
            return {"error": "Файл costs.xlsx не найден"}
        
        results = []
        for col in df.columns:
            if "артикул" in col.lower() or "artikul" in col.lower() or "offer_id" in col.lower():
                matches = df[df[col].astype(str).isin([str(a) for a in artikuls])]
                for _, row in matches.iterrows():
                    results.append({col: str(row[col]) if pd.notna(row[col]) else None for col in df.columns})
                break
        
        return {"results": results, "count": len(results)}
    
    def get_abcxyz_summary(self, args: Dict[str, Any]) -> Dict[str, Any]:
        """Возвращает сводку ABC&XYZ отчёта."""
        period = args.get("period", "")
        if not period:
            return {"error": "Параметр period обязателен"}
        
        reports = list_abc_xyz_reports(self.repo_root)
        report_path = None
        for r in reports:
            if period in r.stem:
                report_path = r
                break
        
        if not report_path:
            return {"error": f"ABC&XYZ отчёт за период '{period}' не найден"}
        
        summary = load_abc_xyz_summary(report_path)
        if summary:
            return summary
        return {"error": "Не удалось загрузить сводку"}
    
    def get_artikul_abcxyz(self, args: Dict[str, Any]) -> Dict[str, Any]:
        """Возвращает категорию артикула в ABC&XYZ."""
        # Упрощённая реализация - можно расширить
        return {"error": "Функция в разработке"}
    
    def get_category_list(self, args: Dict[str, Any]) -> Dict[str, Any]:
        """Возвращает список артикулов в категории."""
        # Упрощённая реализация - можно расширить
        return {"error": "Функция в разработке"}


def parse_ai_response(response_text: str) -> Optional[Dict[str, Any]]:
    """
    Парсит ответ ИИ и извлекает DATA_REQUEST или FINAL_ANSWER.
    
    Returns:
        Словарь с type и данными, или None если не удалось распарсить
    """
    # Убираем лишние пробелы в начале и конце
    response_text = response_text.strip()
    
    def normalize_parsed_response(parsed: Any) -> Optional[Dict[str, Any]]:
        def coerce_tool_call(item: Any) -> List[Dict[str, Any]]:
            if not isinstance(item, dict):
                return []

            tool_name = item.get("tool")
            if not isinstance(tool_name, str) or not tool_name.strip():
                tool_name = item.get("instrument")
            if not isinstance(tool_name, str) or not tool_name.strip():
                tool_name = item.get("action")
            if not isinstance(tool_name, str) or not tool_name.strip():
                tool_name = item.get("method")

            args = item.get("args")
            if not isinstance(args, dict):
                args = item.get("parameters")
            if not isinstance(args, dict):
                args = item.get("params")
            if not isinstance(args, dict):
                args = item.get("arguments")
            if not isinstance(args, dict):
                args = {}

            if isinstance(tool_name, str) and tool_name.strip() and not args:
                reserved = {
                    "tool", "instrument", "action", "method",
                    "args", "parameters", "params", "arguments",
                }
                args = {
                    key: value
                    for key, value in item.items()
                    if key not in reserved
                }

            return expand_args_for_tool(tool_name, args)

        def expand_args_for_tool(tool_name: str, args: Any) -> List[Dict[str, Any]]:
            if not isinstance(tool_name, str) or not tool_name.strip():
                return []
            clean_tool = tool_name.strip()
            if not isinstance(args, dict):
                args = {}

            months_value = args.get("months")
            year_value = args.get("year")
            if clean_tool == "get_month_report_full_data" and isinstance(months_value, list):
                month_map = {
                    "january": "Январь",
                    "february": "Февраль",
                    "march": "Март",
                    "april": "Апрель",
                    "may": "Май",
                    "june": "Июнь",
                    "july": "Июль",
                    "august": "Август",
                    "september": "Сентябрь",
                    "october": "Октябрь",
                    "november": "Ноябрь",
                    "december": "Декабрь",
                }
                expanded: List[Dict[str, Any]] = []
                for item in months_value:
                    if not isinstance(item, str):
                        continue
                    month_text = item.strip()
                    if not month_text:
                        continue
                    lower_month = month_text.lower()
                    if re.match(r"^[а-яa-z]+\s+\d{4}$", lower_month):
                        month_name, year_part = lower_month.split(maxsplit=1)
                        month_ru = month_map.get(month_name, month_name.capitalize())
                        expanded.append({"tool": clean_tool, "args": {"period": f"{month_ru} {year_part}"}})
                        continue
                    month_ru = month_map.get(lower_month)
                    if month_ru and year_value is not None:
                        expanded.append({"tool": clean_tool, "args": {"period": f"{month_ru} {year_value}"}})
                if expanded:
                    return expanded

            return [{"tool": clean_tool, "args": args}]

        if not isinstance(parsed, dict):
            return None

        if "type" not in parsed and isinstance(parsed.get("needs"), list):
            reason = parsed.get("reason")
            if not isinstance(reason, str):
                reason = ""
            generated_needs: List[Dict[str, Any]] = []
            for item in parsed.get("needs", []):
                generated_needs.extend(coerce_tool_call(item))
            return {"type": "DATA_REQUEST", "needs": generated_needs, "reason": reason}

        if "type" in parsed:
            parsed_type = parsed.get("type")
            normalized_type = str(parsed_type).strip().upper() if parsed_type is not None else ""
            if normalized_type == "FINAL_ANSWER":
                answer = parsed.get("answer")
                if not isinstance(answer, str):
                    answer = parsed.get("content")
                if not isinstance(answer, str):
                    answer = parsed.get("text")
                return {"type": "FINAL_ANSWER", "answer": answer or ""}

            if normalized_type == "DATA_REQUEST":
                needs = parsed.get("needs")
                if isinstance(needs, list):
                    normalized_needs: List[Dict[str, Any]] = []
                    for item in needs:
                        normalized_needs.extend(coerce_tool_call(item))
                    needs = normalized_needs
                else:
                    requests_list = parsed.get("requests")
                    if isinstance(requests_list, list):
                        needs = []
                        for item in requests_list:
                            needs.extend(coerce_tool_call(item))
                if not isinstance(needs, list):
                    needs = []
                if not needs:
                    data_block = parsed.get("data")
                    if isinstance(data_block, dict):
                        months = data_block.get("months")
                        year = data_block.get("year")
                        if isinstance(months, list) and year is not None:
                            month_map = {
                                "january": "Январь",
                                "february": "Февраль",
                                "march": "Март",
                                "april": "Апрель",
                                "may": "Май",
                                "june": "Июнь",
                                "july": "Июль",
                                "august": "Август",
                                "september": "Сентябрь",
                                "october": "Октябрь",
                                "november": "Ноябрь",
                                "december": "Декабрь",
                            }
                            generated_needs = []
                            for month in months:
                                if not isinstance(month, str):
                                    continue
                                month_ru = month_map.get(month.strip().lower())
                                if not month_ru:
                                    continue
                                generated_needs.append(
                                    {
                                        "tool": "get_month_report_full_data",
                                        "args": {"period": f"{month_ru} {year}"},
                                    }
                                )
                            if generated_needs:
                                needs = generated_needs
                reason = parsed.get("reason")
                if not isinstance(reason, str):
                    reason = ""
                return {"type": "DATA_REQUEST", "needs": needs, "reason": reason}

        # Fallback для legacy-схемы некоторых моделей:
        # {"tool":"list_reports","parameters":{}}
        tool_name = parsed.get("tool")
        if not isinstance(tool_name, str) or not tool_name.strip():
            tool_name = parsed.get("action")
        if isinstance(tool_name, str) and tool_name.strip():
            args = parsed.get("args")
            if not isinstance(args, dict):
                args = parsed.get("parameters")
            if not isinstance(args, dict):
                args = parsed.get("params")
            if not isinstance(args, dict):
                args = parsed.get("arguments")
            if not isinstance(args, dict):
                args = {}
            reason = parsed.get("reason")
            if not isinstance(reason, str) or not reason.strip():
                reason = f"Нужны данные через инструмент {tool_name}"
            return {"type": "DATA_REQUEST", "needs": expand_args_for_tool(tool_name, args), "reason": reason}

        # Fallback для dict вида {"get_month_report_full_data":[{"period":"Январь 2026"}, ...]}
        generated_needs: List[Dict[str, Any]] = []
        for key, value in parsed.items():
            if not isinstance(key, str) or not key.strip():
                continue
            if key in {"type", "reason", "answer", "content", "text", "data", "needs", "requests", "tool", "action", "args", "parameters", "params", "arguments"}:
                continue
            if isinstance(value, list):
                for item in value:
                    generated_needs.extend(expand_args_for_tool(key, item))
            elif isinstance(value, dict):
                generated_needs.extend(expand_args_for_tool(key, value))
        if generated_needs:
            reason = parsed.get("reason")
            if not isinstance(reason, str) or not reason.strip():
                reason = "Нужны дополнительные данные из инструментов"
            return {"type": "DATA_REQUEST", "needs": generated_needs, "reason": reason}
        return None

    # Пробуем распарсить весь текст как JSON (если это чистый JSON)
    try:
        parsed = json.loads(response_text)
        normalized = normalize_parsed_response(parsed)
        if normalized:
            return normalized
    except json.JSONDecodeError:
        pass
    
    # Пробуем найти JSON в code blocks (```json ... ```)
    json_match = re.search(r'```json\s*(\{.*?\})\s*```', response_text, re.DOTALL)
    if json_match:
        try:
            parsed = json.loads(json_match.group(1))
            normalized = normalize_parsed_response(parsed)
            if normalized:
                return normalized
        except json.JSONDecodeError:
            pass
    
    # Пробуем найти JSON объект с балансом скобок
    # Ищем начало JSON объекта
    start_idx = response_text.find('{')
    if start_idx != -1:
        # Находим соответствующий закрывающий символ
        brace_count = 0
        in_string = False
        escape_next = False
        
        for i in range(start_idx, len(response_text)):
            char = response_text[i]
            
            if escape_next:
                escape_next = False
                continue
            
            if char == '\\':
                escape_next = True
                continue
            
            if char == '"' and not escape_next:
                in_string = not in_string
                continue
            
            if not in_string:
                if char == '{':
                    brace_count += 1
                elif char == '}':
                    brace_count -= 1
                    if brace_count == 0:
                        # Нашли полный JSON объект
                        json_str = response_text[start_idx:i+1]
                        try:
                            parsed = json.loads(json_str)
                            normalized = normalize_parsed_response(parsed)
                            if normalized:
                                return normalized
                        except json.JSONDecodeError:
                            pass
                        break
    
    # Если не нашли JSON - возвращаем None (не делаем fallback на FINAL_ANSWER)
    return None


def execute_tools(tools: Tools, needs: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Выполняет запросы к инструментам.
    
    Args:
        tools: Экземпляр класса Tools
        needs: Список запросов [{"tool": "...", "args": {...}}]
    
    Returns:
        Словарь с результатами выполнения инструментов.
        Если один инструмент вызывается несколько раз - результаты объединяются в список.
    """
    results = {}
    
    for idx, need in enumerate(needs):
        tool_name = need.get("tool", "")
        args = need.get("args", {})
        
        if not hasattr(tools, tool_name):
            key = f"{tool_name}_{idx}" if tool_name in results else tool_name
            results[key] = {"error": f"Инструмент '{tool_name}' не найден"}
            continue
        
        try:
            tool_func = getattr(tools, tool_name)
            result = tool_func(args)
            
            # Если инструмент уже вызывался - создаём список результатов
            if tool_name in results:
                # Если уже список - добавляем, иначе преобразуем в список
                if isinstance(results[tool_name], list):
                    results[tool_name].append(result)
                else:
                    results[tool_name] = [results[tool_name], result]
            else:
                results[tool_name] = result
        except Exception as e:
            key = f"{tool_name}_{idx}" if tool_name in results else tool_name
            results[key] = {"error": str(e)}
    
    return results


def format_tool_results(
    results: Dict[str, Any],
    user_message: Optional[str] = None,
    workflow_state: Optional[WorkflowState] = None,
) -> str:
    """Фасад над payload policy для передачи данных инструментов обратно в модель."""
    return format_model_payload_core(results, user_message, workflow_state, CURRENT_YEAR)


def build_chat_messages(
    user_message: str,
    conversation_history: list = None,
    data_block: Optional[str] = None,
    context_state: Optional[str] = None,
    allow_followup_data_requests: bool = True,
) -> List[Dict[str, str]]:
    system_prompt = SYSTEM_PROMPT

    if context_state:
        system_prompt += f"\n\n## Контекст состояния:\n{context_state}\n"

    messages = [{"role": "system", "content": system_prompt}]

    if conversation_history:
        recent_history = conversation_history[-6:] if len(conversation_history) > 6 else conversation_history
        messages.extend(recent_history)

    if data_block:
        messages.append(
            {
                "role": "user",
                "content": (
                    (
                        "Ниже переданы данные, уже полученные из инструментов.\n"
                        "Если этих данных достаточно, верни FINAL_ANSWER.\n"
                        "Если данных всё ещё недостаточно, верни DATA_REQUEST только с дополнительными нужными инструментами.\n"
                        if allow_followup_data_requests
                        else "Данные уже получены. На их основе верни только FINAL_ANSWER.\n"
                    )
                    + f"Вопрос пользователя: {user_message}\n"
                    + f"Данные инструментов: {data_block}"
                ),
            }
        )
    else:
        messages.append({"role": "user", "content": user_message})

    return messages


def chat_with_ai(user_message: str, conversation_history: list = None, data_block: Optional[str] = None, context_state: Optional[str] = None, temperature: float = 0.7, max_tokens: Optional[int] = None, allow_followup_data_requests: bool = True) -> Tuple[Optional[str], Optional[Dict[str, Any]], Optional[Dict[str, Any]]]:
    """
    ?????????? ????????? ? ????????? Ollama ? ???????? ?????.
    """
    messages = build_chat_messages(user_message, conversation_history, data_block, context_state, allow_followup_data_requests=allow_followup_data_requests)
    return _send_ollama_chat(messages, temperature, max_tokens)


def stream_chat_with_ai(
    user_message: str,
    conversation_history: list = None,
    data_block: Optional[str] = None,
    context_state: Optional[str] = None,
    temperature: float = 0.7,
    max_tokens: Optional[int] = None,
    allow_followup_data_requests: bool = True,
    on_chunk=None,
    on_thinking=None,
    should_cancel=None,
) -> Tuple[Optional[str], Optional[Dict[str, Any]], Optional[Dict[str, Any]]]:
    messages = build_chat_messages(
        user_message,
        conversation_history,
        data_block,
        context_state,
        allow_followup_data_requests=allow_followup_data_requests,
    )
    return _stream_ollama_chat(
        messages,
        temperature,
        max_tokens,
        on_chunk=on_chunk,
        on_thinking=on_thinking,
        should_cancel=should_cancel,
    )

def start_chat_session():
    """
    Запускает интерактивную сессию чата с ИИ.
    """
    print("\n· OzonReportX AI — чат с ИИ. Вопросы по бизнесу Ozon. Выход: выход / exit / quit\n")

    ok, message = check_ollama_model()
    if not ok:
        print(f"⚠️ {message}")
        return
    
    # Определяем корень репозитория
    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parent
    
    # Создаём экземпляр инструментов
    tools = Tools(repo_root)
    
    # Память сессии (server-side, без токенов)
    session_memory = {
        "selected_period": None,
        "selected_artikul": None,
        "default_margin_settings": None,
        "last_comparison_periods": []
    }
    
    print(f"✅ Готов. Модель: {OLLAMA_MODEL} | Ollama: {OLLAMA_HOST}\n")
    
    conversation_history = []
    
    while True:
        try:
            # Получаем сообщение от пользователя
            user_input = input("Вы: ").strip()
            
            if not user_input:
                continue
            
            # Проверяем команды выхода
            if user_input.lower() in ('выход', 'exit', 'quit', 'q'):
                print("\n👋 Выход в меню.\n")
                break
            
            # Формируем контекст состояния
            context_state_lines = []
            if session_memory["selected_period"]:
                context_state_lines.append(f"selected_period={session_memory['selected_period']}")
            if session_memory["selected_artikul"]:
                context_state_lines.append(f"selected_artikul={session_memory['selected_artikul']}")
            workflow_state = detect_workflow(user_input)
            if workflow_state.name != "generic":
                context_state_lines.append(f"workflow={workflow_state.name}")
            context_state = "\n".join(context_state_lines) if context_state_lines else None
            
            # РЕЖИМ A: План/Уточнение
            print("   …", end="\r")
            ai_response, usage_info, rate_limit_info = chat_with_ai(
                user_input, 
                conversation_history, 
                data_block=None,
                context_state=context_state,
                temperature=0.7,
                max_tokens=None
            )
            
            print(" " * 50, end="\r")
            
            if not ai_response:
                print("\n⚠️ Не удалось получить ответ от ИИ. Попробуйте ещё раз.\n")
                continue
            
            # Инициализируем rate_limit_info, если его нет
            if rate_limit_info is None:
                rate_limit_info = {}
            
            # Парсим ответ ИИ
            parsed_response = parse_ai_response(ai_response)
            
            # Если парсер вернул None - делаем один ретрай с просьбой переформатировать
            if parsed_response is None:
                print("   🔄 Переформатирую ответ...", end="\r")
                retry_messages = [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": f"""Твой предыдущий ответ не был валидным JSON. Переформатируй его в правильный формат.

Твой предыдущий ответ:
{ai_response}

Верни ТОЛЬКО валидный JSON объект (без Markdown, без ```) в формате:
- {{"type": "DATA_REQUEST", "needs": [...], "reason": "..."}} или
- {{"type": "FINAL_ANSWER", "answer": "..."}}"""}
                ]
                
                try:
                    retry_ai_response, _, _ = _send_ollama_chat(retry_messages, 0.2, 500)
                    if retry_ai_response:
                        parsed_response = parse_ai_response(retry_ai_response)
                        if parsed_response:
                            ai_response = retry_ai_response
                except Exception:
                    pass
                
                print(" " * 50, end="\r")
                
                # Если после ретрая всё ещё None - fallback на FINAL_ANSWER
                if parsed_response is None:
                    parsed_response = {"type": "FINAL_ANSWER", "answer": "Извините, произошла ошибка при обработке ответа. Попробуйте переформулировать вопрос."}
            
            if parsed_response and parsed_response.get("type") == "DATA_REQUEST":
                # ИИ запросил данные - выполняем инструменты (скрыто от пользователя)
                needs = parsed_response.get("needs", [])
                reason = parsed_response.get("reason", "")
                
                # Показываем только индикатор загрузки
                print("   …", end="\r")
                
                # Выполняем инструменты
                tool_results = execute_tools(tools, needs)
                merged_results = dict(tool_results)

                deterministic_needs = get_workflow_followup_needs(user_input, merged_results)
                if deterministic_needs:
                    additional_prefetch_results = execute_tools(tools, deterministic_needs)
                    for key, value in additional_prefetch_results.items():
                        if key not in merged_results:
                            merged_results[key] = value
                        else:
                            existing = merged_results[key]
                            if isinstance(existing, list):
                                if isinstance(value, list):
                                    existing.extend(value)
                                else:
                                    existing.append(value)
                            else:
                                merged_results[key] = [existing, value] if not isinstance(value, list) else [existing, *value]

                workflow_state = build_workflow_state(user_input, merged_results)
                data_block = format_tool_results(
                    merged_results,
                    user_message=user_input,
                    workflow_state=workflow_state,
                )
                
                # Обновляем память сессии на основе запросов
                # Если запрошено несколько периодов - сохраняем последний
                periods_found = []
                for need in needs:
                    args = need.get("args", {})
                    if "period" in args:
                        periods_found.append(args["period"])
                    if "artikul" in args:
                        session_memory["selected_artikul"] = args["artikul"]
                
                # Сохраняем последний период (или первый, если один)
                if periods_found:
                    session_memory["selected_period"] = periods_found[-1]
                    # Сохраняем список периодов для сравнения
                    if len(periods_found) > 1:
                        session_memory["last_comparison_periods"] = periods_found
                
                # РЕЖИМ B: Ответ с данными (temperature=0.2 для точности формата)
                print("🤖 Анализирую данные...", end="\r")
                
                ai_response_final, usage_info2, rate_limit_info2 = chat_with_ai(
                    user_input,
                    conversation_history,
                    data_block=data_block,
                    context_state=context_state,
                    temperature=0.2,  # Низкая температура для режима B
                    max_tokens=None,
                    allow_followup_data_requests=not workflow_state.force_final_answer,
                )
                
                print(" " * 50, end="\r")
                
                if ai_response_final:
                    parsed_final = parse_ai_response(ai_response_final)
                    
                    # Если ИИ снова вернул DATA_REQUEST после получения данных - обрабатываем рекурсивно (макс 2 уровня)
                    max_iterations = 2
                    iteration = 0
                    current_response = ai_response_final
                    current_parsed = parsed_final
                    
                    while current_parsed and current_parsed.get("type") == "DATA_REQUEST" and iteration < max_iterations:
                        iteration += 1
                        print("   …", end="\r")
                        
                        # Выполняем дополнительные запросы
                        additional_needs = current_parsed.get("needs", [])
                        additional_results = execute_tools(tools, additional_needs)
                        additional_workflow_state = build_workflow_state(user_input, additional_results)
                        additional_data = format_tool_results(
                            additional_results,
                            user_message=user_input,
                            workflow_state=additional_workflow_state,
                        )
                        
                        # Объединяем с предыдущими данными
                        combined_data = f"{data_block}\n\nДОПОЛНИТЕЛЬНЫЕ ДАННЫЕ:\n{additional_data}"
                        
                        # Запрашиваем финальный ответ с объединёнными данными
                        current_response, _, rate_limit_info2 = chat_with_ai(
                            user_input,
                            conversation_history,
                            data_block=combined_data,
                            context_state=context_state,
                            temperature=0.2,  # Низкая температура для режима B
                            max_tokens=None,
                            allow_followup_data_requests=not additional_workflow_state.force_final_answer,
                        )
                        
                        if current_response:
                            current_parsed = parse_ai_response(current_response)
                        else:
                            break
                    
                    print(" " * 50, end="\r")
                    
                    # Извлекаем финальный ответ
                    if current_parsed and current_parsed.get("type") == "FINAL_ANSWER":
                        final_answer = current_parsed.get("answer", current_response)
                    elif current_parsed and current_parsed.get("type") == "DATA_REQUEST":
                        # Если ИИ всё ещё запрашивает данные после всех итераций - принудительно извлекаем ответ
                        final_answer = build_workflow_fallback_answer(user_input, merged_results) or current_response
                        # Убираем JSON из ответа
                        if "{" in final_answer and '"type"' in final_answer:
                            json_start = final_answer.find('{')
                            if json_start > 0:
                                final_answer = final_answer[:json_start].strip()
                            else:
                                json_end = final_answer.rfind('}')
                                if json_end < len(final_answer) - 1:
                                    final_answer = final_answer[json_end+1:].strip()
                        
                        if not final_answer or len(final_answer) < 10:
                            final_answer = build_workflow_fallback_answer(user_input, merged_results) or "Проанализировал полученные данные. Пожалуйста, уточните ваш вопрос или запросите конкретные метрики."
                    else:
                        final_answer = current_response or build_workflow_fallback_answer(user_input, merged_results)
                    
                    # Показываем только финальный ответ пользователю
                    print(f"\n🤖 ИИ: {final_answer}\n")
                    
                    # Сохраняем в историю (только финальный ответ, не DATA_REQUEST)
                    conversation_history.append({"role": "user", "content": user_input})
                    conversation_history.append({"role": "assistant", "content": final_answer})
                    
                    # Используем данные о лимитах из финального запроса
                    if rate_limit_info2:
                        rate_limit_info = rate_limit_info2
                else:
                    print("\n⚠ Не удалось получить ответ от ИИ.\n")
                    continue
                    
            else:
                # Финальный ответ без запроса данных
                if parsed_response and parsed_response.get("type") == "FINAL_ANSWER":
                    final_answer = parsed_response.get("answer", ai_response)
                else:
                    # Если ответ не в формате JSON, используем его как есть
                    final_answer = ai_response
                    # Убираем JSON из ответа, если он там случайно оказался
                    if "{" in final_answer and '"type"' in final_answer and '"DATA_REQUEST"' in final_answer:
                        # Это ошибка - ИИ вернул DATA_REQUEST вместо FINAL_ANSWER
                        # Пробуем извлечь текстовую часть
                        json_start = final_answer.find('{')
                        if json_start > 0:
                            final_answer = final_answer[:json_start].strip()
                        if not final_answer:
                            final_answer = "Для ответа на ваш вопрос мне нужны данные. Пожалуйста, уточните, какие именно данные вас интересуют."
                
                print(f"\n🤖 ИИ: {final_answer}\n")
                
                # Сохраняем в историю
                conversation_history.append({"role": "user", "content": user_input})
                conversation_history.append({"role": "assistant", "content": final_answer})
            
            # Ограничиваем историю последними 6 сообщениями (3 пары вопрос-ответ)
            if len(conversation_history) > 6:
                conversation_history = conversation_history[-6:]
            
            # Выводим информацию о модели и лимитах из ?????????? Ollama
            model = rate_limit_info.get("model", OLLAMA_MODEL) if rate_limit_info else OLLAMA_MODEL
            info_lines = [f"?? ??????: {model}"]
            host = rate_limit_info.get("host") if rate_limit_info else OLLAMA_HOST
            if host:
                info_lines.append(f"Host: {host}")
            print(" | ".join(info_lines) + "\n")
                
        except KeyboardInterrupt:
            print("\n👋 Выход в меню.\n")
            break
        except EOFError:
            print("\n👋 Выход в меню.\n")
            break
        except Exception as e:
            print(f"\n⚠ Ошибка: {e}\n")


def main():
    """Точка входа для запуска модуля как скрипта."""
    start_chat_session()


if __name__ == "__main__":
    main()

