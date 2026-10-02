from __future__ import annotations

import os
import asyncio
import json
import queue
import re
import requests
import subprocess
import sys
import threading
import time
from dataclasses import dataclass
from collections import deque
from uuid import uuid4
from datetime import date, datetime, timedelta
from pathlib import Path

import flet as ft
from openpyxl import load_workbook
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scripts import balance_report
from scripts.file_io import save_workbook_atomic
from scripts.file_lock import exclusive_file, file_signature
from scripts.ui_log import SessionLog
from scripts.wb_dashboard import build_dashboard as build_wb_dashboard
from scripts.marketplace_settings import FIELDS, HELP, read_settings, save_settings, needs_setup

try:
    import scripts.ai_chat as ai_chat_module
    from scripts.ai_chat import (
        OLLAMA_HOST,
        OLLAMA_MODEL,
        Tools as AITools,
        stream_chat_with_ai,
        chat_with_ai,
        check_ollama_model,
        execute_tools,
        format_tool_results,
        parse_ai_response,
        build_workflow_state,
        get_workflow_followup_needs,
        build_workflow_fallback_answer,
    )
    AI_IMPORT_ERROR = None
except Exception as exc:
    ai_chat_module = None
    OLLAMA_HOST = "http://localhost:11434"
    OLLAMA_MODEL = "qwen3:4b"
    AITools = None
    stream_chat_with_ai = None
    chat_with_ai = None
    check_ollama_model = None
    execute_tools = None
    format_tool_results = None
    parse_ai_response = None
    build_workflow_state = None
    get_workflow_followup_needs = None
    build_workflow_fallback_answer = None
    AI_IMPORT_ERROR = str(exc)

try:
    from scripts.price_management import (
        build_discount_request_plan,
        process_discount_request_item,
        process_discount_request_plan,
    )
    PRICING_IMPORT_ERROR = None
except Exception as exc:
    build_discount_request_plan = None
    process_discount_request_item = None
    process_discount_request_plan = None
    PRICING_IMPORT_ERROR = str(exc)


ROOT = Path(__file__).resolve().parent
SCRIPTS = ROOT / "scripts"
PYTHON = next(
    (p for p in [ROOT / ".venv" / "Scripts" / "python.exe", ROOT / ".venv" / "bin" / "python"] if p.exists()),
    Path("python"),
)

BG = "#F5F5F7"
SURFACE = "#FFFFFF"
PRIMARY = "#0E6B5C"
PRIMARY_SOFT = "#E7F2EF"
ACCENT = "#A56B36"
TEXT = "#1D1D1F"
MUTED = "#6E6E73"
DANGER = "#A24438"
BORDER = "#D2D2D7"
COSTS_MAIN_SHEET = "Основной"
COSTS_MAIN_COLUMNS = [
    "Артикул",
    "Себестоимость",
    "Минимальная цена продажи",
    "Желательная цена продажи",
    "Текущая цена на Ozon",
    "Цена с учётом акций и скидок",
    "Текущая ожидаемая рентабельность",
]
COSTS_MAIN_COLUMN_EXPANDS = {
    "Артикул": 12,
    "Себестоимость": 11,
    "Минимальная цена продажи": 14,
    "Желательная цена продажи": 14,
    "Текущая цена на Ozon": 14,
    "Цена с учётом акций и скидок": 16,
    "Текущая ожидаемая рентабельность": 14,
}
PRICING_CALCULATED_COLUMNS = ("Минимальная цена продажи", "Желательная цена продажи")
PRICING_SYNC_COLUMNS = ("Текущая цена на Ozon", "Цена с учётом акций и скидок", "Текущая ожидаемая рентабельность")
EDITABLE_COSTS_COLUMNS = {"Артикул", "Себестоимость"}
MONTHS = [
    "Январь",
    "Февраль",
    "Март",
    "Апрель",
    "Май",
    "Июнь",
    "Июль",
    "Август",
    "Сентябрь",
    "Октябрь",
    "Ноябрь",
    "Декабрь",
]
MONTH_INDEX = {name.lower(): idx for idx, name in enumerate(MONTHS, start=1)}
PRIMARY_TEXT_COLORS = {"#1D1D1F", "#F5F5F7", "#1D2A2E", "#F3F4F6", "#202123", "#ECECF1", "#1F1F22"}
MUTED_TEXT_COLORS = {"#6E6E73", "#A1A1A6", "#6F7A76", "#B6C0BC"}


@dataclass
class FolderSpec:
    title: str
    path: Path
    pattern: str = "*.xlsx"


def open_path(path: Path) -> None:
    if not path.exists():
        return
    if os.name == "nt":
        os.startfile(str(path))
        return
    subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(path)])


def stream_process(
    cmd: list[str],
    on_line,
    on_prompt=None,
    extra_env: dict[str, str] | None = None,
    proc_holder: dict[str, subprocess.Popen | None] | None = None,
    cancel_event: threading.Event | None = None,
) -> tuple[int, str]:
    if cancel_event is not None and cancel_event.is_set():
        return 130, "Остановлено пользователем."
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUNBUFFERED"] = "1"
    if extra_env:
        env.update(extra_env)
    proc = subprocess.Popen(
        cmd,
        cwd=ROOT,
        env=env,
        stdout=subprocess.PIPE,
        stdin=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )
    if proc_holder is not None:
        proc_holder["proc"] = proc
    try:
        if cancel_event is not None and cancel_event.is_set():
            proc.kill()
        lines = deque(maxlen=500)
        assert proc.stdout is not None
        for line in proc.stdout:
            lines.append(line)
            on_line(line)
            if on_prompt and proc.stdin is not None:
                lower = line.lower()
                label = None
                if "введите сумму затрат на продвижение ozon" in lower:
                    label = "Продвижение Ozon"
                elif "введите сумму затрат на внешний маркетинг" in lower:
                    label = "Внешний маркетинг"
                if label:
                    answer = on_prompt(label, line.strip())
                    try:
                        proc.stdin.write(f"{answer}\n")
                        proc.stdin.flush()
                    except (BrokenPipeError, OSError):
                        pass
        proc.wait()
        return proc.returncode, "".join(lines).strip()
    finally:
        if proc.poll() is None:
            proc.kill()
            proc.wait()
        if proc.stdout is not None:
            proc.stdout.close()
        if proc.stdin is not None:
            proc.stdin.close()
        if proc_holder is not None and proc_holder.get("proc") is proc:
            proc_holder["proc"] = None


def files(folder: Path, pattern: str = "*.xlsx") -> list[Path]:
    return sorted((p for p in folder.glob(pattern) if p.is_file() and not p.name.startswith(("~$", "~tmp_"))), key=lambda p: p.stat().st_mtime, reverse=True) if folder.exists() else []


def _period_sort_key(path: Path) -> tuple[int, int, int]:
    stem = path.stem.lower()
    matches = re.findall(r"(январь|февраль|март|апрель|май|июнь|июль|август|сентябрь|октябрь|ноябрь|декабрь)\s+(\d{4})", stem)
    if matches:
        month_name, year_text = matches[-1]
        return (int(year_text), MONTH_INDEX.get(month_name, 0), 1)
    try:
        return (datetime.fromtimestamp(path.stat().st_mtime).year, datetime.fromtimestamp(path.stat().st_mtime).month, 0)
    except OSError:
        return (0, 0, 0)


def files_by_relevance(folder: Path, pattern: str = "*.xlsx") -> list[Path]:
    if not folder.exists():
        return []
    batch = [p for p in folder.glob(pattern) if p.is_file() and not p.name.startswith(("~$", "~tmp_"))]
    return sorted(batch, key=lambda p: (_period_sort_key(p), p.stat().st_mtime if p.exists() else 0), reverse=True)


def report_path_for_period(month: int, year: int) -> Path:
    return ROOT / "reports" / f"{MONTHS[month - 1]} {year}.xlsx"


def iter_periods(month: int, year: int, count: int) -> list[tuple[int, int]]:
    periods: list[tuple[int, int]] = []
    cursor_month = month
    cursor_year = year
    for _ in range(count):
        periods.append((cursor_month, cursor_year))
        cursor_month -= 1
        if cursor_month == 0:
            cursor_month = 12
            cursor_year -= 1
    periods.reverse()
    return periods


def _format_number(value) -> str:
    if value is None:
        return "—"
    if isinstance(value, float):
        if value.is_integer():
            return f"{int(value):,}".replace(",", " ")
        return f"{value:,.2f}".replace(",", " ").replace(".", ",")
    if isinstance(value, int):
        return f"{value:,}".replace(",", " ")
    return str(value)


def _format_currency(value) -> str:
    if value is None:
        return "—"
    if isinstance(value, (int, float)):
        return f"{_format_number(float(value))} ₽"
    return str(value)


def _format_percent(value) -> str:
    if value is None:
        return "—"
    if isinstance(value, (int, float)):
        return f"{float(value):.2f}%".replace(".", ",")
    return str(value)


def _format_costs_preview_value(column: str, value) -> str:
    if value is None:
        return "—"
    if column == "Артикул":
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        text = str(value).strip()
        return text[:-2] if text.endswith(".0") else (text or "—")
    numeric = as_float(value)
    if numeric is not None:
        if "рентабельность" in column.lower():
            percent_value = numeric * 100 if abs(numeric) <= 1 else numeric
            return _format_percent(percent_value)
        return _format_currency(numeric)
    text = str(value).strip()
    return text or "—"


def _format_costs_editor_value(column: str, value) -> str:
    if value is None:
        return ""
    if column == "Артикул":
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        text = str(value).strip()
        return text[:-2] if text.endswith(".0") else text
    numeric = as_float(value)
    if numeric is not None:
        return str(int(numeric)) if float(numeric).is_integer() else str(numeric).replace(".", ",")
    return str(value).strip()


def _excel_color_to_hex(color) -> str | None:
    if color is None:
        return None
    color_type = getattr(color, "type", None)
    rgb = getattr(color, "rgb", None)
    if color_type == "rgb" and isinstance(rgb, str):
        hex_value = rgb[-6:].upper()
        if len(hex_value) == 6:
            return f"#{hex_value}"
    return None


def _excel_fill_color(cell) -> str | None:
    fill = getattr(cell, "fill", None)
    if fill is None or getattr(fill, "patternType", None) != "solid":
        return None
    return _excel_color_to_hex(getattr(fill, "fgColor", None)) or _excel_color_to_hex(getattr(fill, "startColor", None))


def _excel_font_color(cell) -> str | None:
    font = getattr(cell, "font", None)
    if font is None:
        return None
    return _excel_color_to_hex(getattr(font, "color", None))


def _hex_to_rgb(color: str | None) -> tuple[int, int, int] | None:
    if not color or not isinstance(color, str):
        return None
    value = color.lstrip("#")
    if len(value) != 6:
        return None
    try:
        return tuple(int(value[index:index + 2], 16) for index in (0, 2, 4))
    except ValueError:
        return None


def _excel_text_color(background_color: str | None, fallback_dark: bool) -> str:
    rgb = _hex_to_rgb(background_color)
    if rgb is None:
        return "#F3F4F6" if fallback_dark else TEXT
    red, green, blue = rgb
    luminance = (0.299 * red + 0.587 * green + 0.114 * blue) / 255
    return "#1D1D1F" if luminance >= 0.62 else "#F5F5F7"


def _build_costs_cell(column: str, cell_payload: dict[str, object] | None, dark: bool, expand: int) -> ft.Control:
    payload = cell_payload or {}
    value = payload.get("value")
    fill_color = payload.get("fill_color") if isinstance(payload.get("fill_color"), str) else None
    font_color = payload.get("font_color") if isinstance(payload.get("font_color"), str) else None
    resolved_text_color = font_color or _excel_text_color(fill_color, dark)
    return ft.Container(
        expand=expand,
        padding=ft.Padding.symmetric(horizontal=10, vertical=8),
        border_radius=10,
        bgcolor=fill_color,
        data="preserve_excel_fill" if fill_color else None,
        content=ft.Text(
            _format_costs_preview_value(column, value),
            size=12,
            color=resolved_text_color,
            weight=ft.FontWeight.W_700 if column == "Артикул" else ft.FontWeight.W_500,
            selectable=True,
            data="preserve_excel_text" if (fill_color or font_color) else None,
        ),
    )


def _pricing_conditional_fill(
    column: str,
    row: dict[str, dict[str, object]],
    min_margin_pct: float = 23.0,
    desired_margin_pct: float = 30.0,
) -> str | None:
    min_price_payload = row.get("Минимальная цена продажи") or {}
    min_price = as_float(min_price_payload.get("value"))
    value_payload = row.get(column) or {}
    value = as_float(value_payload.get("value"))
    if value is None:
        return None

    if column in {"Текущая цена на Ozon", "Цена с учётом акций и скидок"}:
        if value <= 0 or min_price is None:
            return None
        return "#FFC7CE" if value < min_price else "#C6EFCE"

    if column == "Текущая ожидаемая рентабельность":
        profitability = value * 100 if abs(value) <= 1 else value
        if profitability < min_margin_pct:
            return "#FFC7CE"
        if min_margin_pct <= profitability <= desired_margin_pct:
            return "#C6EFCE"
    return None


def _apply_costs_conditional_formats(
    rows: list[dict[str, object]],
    min_margin_pct: float = 23.0,
    desired_margin_pct: float = 30.0,
) -> None:
    for row in rows:
        typed_row = row if isinstance(row, dict) else {}
        for column in COSTS_MAIN_COLUMNS:
            payload = typed_row.get(column)
            if not isinstance(payload, dict):
                continue
            if payload.get("fill_color"):
                continue
            conditional_fill = _pricing_conditional_fill(column, typed_row, min_margin_pct=min_margin_pct, desired_margin_pct=desired_margin_pct)
            if conditional_fill:
                payload["fill_color"] = conditional_fill


def read_costs_main_rows(
    costs_path: Path,
    min_margin_pct: float = 23.0,
    desired_margin_pct: float = 30.0,
) -> tuple[list[dict[str, object]], str | None]:
    if not costs_path.exists():
        return [], "Файл costs.xlsx не найден."

    try:
        wb = load_workbook(costs_path, data_only=True)
    except Exception as exc:
        return [], f"Не удалось открыть costs.xlsx: {exc}"

    try:
        if COSTS_MAIN_SHEET not in wb.sheetnames:
            return [], "Лист «Основной» не найден в costs.xlsx."
        ws = wb[COSTS_MAIN_SHEET]
        header_cells = next(ws.iter_rows(min_row=1, max_row=1), None)
        if not header_cells:
            return [], "Лист «Основной» пуст."

        header_map = {
            str(cell.value).strip().lower(): idx
            for idx, cell in enumerate(header_cells)
            if cell.value is not None and str(cell.value).strip()
        }
        missing = [column for column in COSTS_MAIN_COLUMNS if column.lower() not in header_map]
        if missing:
            return [], f"В листе «Основной» нет колонок: {', '.join(missing)}."

        rows: list[dict[str, object]] = []
        for row in ws.iter_rows(min_row=2):
            item: dict[str, object] = {}
            item["__row_number"] = row[0].row if row else None
            has_payload = False
            for column in COSTS_MAIN_COLUMNS:
                cell_index = header_map[column.lower()]
                cell = row[cell_index] if cell_index < len(row) else None
                value = cell.value if cell is not None else None
                if value not in (None, ""):
                    has_payload = True
                item[column] = {
                    "value": value,
                    "fill_color": _excel_fill_color(cell) if cell is not None else None,
                    "font_color": _excel_font_color(cell) if cell is not None else None,
                }
            if has_payload:
                rows.append(item)
        _apply_costs_conditional_formats(rows, min_margin_pct=min_margin_pct, desired_margin_pct=desired_margin_pct)
        return rows, None
    except Exception as exc:
        return [], f"Не удалось прочитать лист «Основной»: {exc}"
    finally:
        wb.close()


def write_costs_main_rows(costs_path: Path, updates: list[dict[str, object]], expected_signature=None, saved_signature=None) -> str | None:
    try:
        with exclusive_file(costs_path):
            if expected_signature is not None and file_signature(costs_path) != expected_signature:
                return "Файл costs.xlsx изменился после открытия таблицы. Правки сохранены в редакторе; обновите данные перед повторным сохранением."
            error = _write_costs_main_rows(costs_path, updates)
            if error is None and saved_signature is not None:
                saved_signature["signature"] = file_signature(costs_path)
            return error
    except Exception as exc:
        return f"Не удалось сохранить costs.xlsx: {exc}"


def _write_costs_main_rows(costs_path: Path, updates: list[dict[str, object]]) -> str | None:
    if not costs_path.exists():
        return "Файл costs.xlsx не найден."

    try:
        wb = load_workbook(costs_path)
    except Exception as exc:
        return f"Не удалось открыть costs.xlsx для сохранения: {exc}"

    try:
        if COSTS_MAIN_SHEET not in wb.sheetnames:
            return "Лист «Основной» не найден в costs.xlsx."
        ws = wb[COSTS_MAIN_SHEET]
        header_cells = next(ws.iter_rows(min_row=1, max_row=1), None)
        if not header_cells:
            return "Лист «Основной» пуст."

        header_map = {
            str(cell.value).strip().lower(): idx + 1
            for idx, cell in enumerate(header_cells)
            if cell.value is not None and str(cell.value).strip()
        }
        missing = [column for column in EDITABLE_COSTS_COLUMNS if column.lower() not in header_map]
        if missing:
            return f"В листе «Основной» нет колонок для сохранения: {', '.join(sorted(missing))}."

        for update in updates:
            row_number = update.get("row_number")
            values = update.get("values")
            if not isinstance(row_number, int) or row_number < 2 or not isinstance(values, dict):
                continue
            for column in EDITABLE_COSTS_COLUMNS:
                if column not in values:
                    continue
                ws.cell(row=row_number, column=header_map[column.lower()]).value = values[column]

        save_workbook_atomic(wb, costs_path)
        return None
    except Exception as exc:
        return f"Не удалось сохранить costs.xlsx: {exc}"
    finally:
        wb.close()


def read_business_metrics(report_path: Path) -> tuple[dict[str, object], str | None]:
    if not report_path.exists():
        return {}, None
    wb = load_workbook(report_path, data_only=True, read_only=True)
    try:
        ws = wb["Заказы"] if "Заказы" in wb.sheetnames else wb.active
        metrics: dict[str, object] = {}
        for label, value in ws.iter_rows(min_row=1, max_row=70, min_col=16, max_col=17, values_only=True):
            if label:
                metrics[str(label).strip()] = value
        needs_fallback = any(
            key not in metrics
            for key in (
                "Количество возвращённых заказов",
                "Количество заказов в доставке",
                "Комиссии Ozon сумма",
                "Логистика сумма",
            )
        )
        if needs_fallback:
            returned_count = 0
            delivering_count = 0
            commission_total = 0.0
            logistics_total = 0.0
            for row in ws.iter_rows(min_row=2, max_col=8, values_only=True):
                status_value = row[0]
                commission_value = row[6]
                logistics_value = row[7]
                if status_value is not None and str(status_value).strip().lower() == "returned":
                    returned_count += 1
                if status_value is not None and str(status_value).strip().lower() == "delivering":
                    delivering_count += 1
                try:
                    if commission_value is not None and str(commission_value).strip() not in ("", "-"):
                        commission_total += abs(float(commission_value))
                except (TypeError, ValueError):
                    pass
                try:
                    if logistics_value is not None and str(logistics_value).strip() not in ("", "-"):
                        logistics_total += abs(float(logistics_value))
                except (TypeError, ValueError):
                    pass
            metrics.setdefault("Количество возвращённых заказов", returned_count)
            metrics.setdefault("Количество заказов в доставке", delivering_count)
            metrics.setdefault("Комиссии Ozon сумма", commission_total)
            metrics.setdefault("Логистика сумма", logistics_total)
        built_at = datetime.fromtimestamp(report_path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        return metrics, built_at
    finally:
        wb.close()


def read_campaign_metrics(report_path: Path) -> tuple[dict[str, float], list[dict[str, object]]]:
    if not report_path.exists():
        return {}, []
    wb = load_workbook(report_path, data_only=True, read_only=True)
    try:
        if "Кампании" not in wb.sheetnames:
            return {}, []
        ws = wb["Кампании"]
        rows = ws.iter_rows(values_only=True)
        headers = next(rows, ())
        header_names = [str(header).strip() if header else f"column_{idx}" for idx, header in enumerate(headers, start=1)]

        campaigns: list[dict[str, object]] = []
        totals = {
            "spend": 0.0,
            "impressions": 0.0,
            "clicks": 0.0,
            "orders": 0.0,
            "revenue": 0.0,
        }

        def safe_metric(row_map: dict[str, object], key: str) -> float:
            return as_float(row_map.get(key)) or 0.0

        for row in rows:
            row_map = dict(zip(header_names, row))
            spend = safe_metric(row_map, "Расход за период (руб.)")
            if spend <= 0:
                continue

            impressions = safe_metric(row_map, "Показы")
            clicks = safe_metric(row_map, "Клики")
            orders = safe_metric(row_map, "Заказы (шт.)")
            revenue = safe_metric(row_map, "Заказы (руб.)")
            ctr = safe_metric(row_map, "CTR (%)")
            cpc = safe_metric(row_map, "Средняя цена клика (руб.)")
            drr = safe_metric(row_map, "ДРР (%)")
            budget = safe_metric(row_map, "Бюджет (руб.)")
            daily_budget = safe_metric(row_map, "Дневной бюджет (руб.)")
            weekly_budget = safe_metric(row_map, "Недельный бюджет (руб.)")

            calc_cpm = (spend / impressions * 1000.0) if impressions > 0 else 0.0
            calc_cpa = (spend / orders) if orders > 0 else None
            calc_roas = (revenue / spend) if spend > 0 else None
            calc_cr = (orders / clicks * 100.0) if clicks > 0 else 0.0
            calc_avg_order_value = (revenue / orders) if orders > 0 else None
            budget_burn = (spend / budget * 100.0) if budget > 0 else None
            daily_budget_burn = (spend / daily_budget * 100.0) if daily_budget > 0 else None
            weekly_budget_burn = (spend / weekly_budget * 100.0) if weekly_budget > 0 else None

            if clicks >= 15 and orders <= 0:
                efficiency_label = "Трафик без заказов"
                efficiency_tone = "#F6DFDB"
                efficiency_reason = "Есть заметный трафик, но клики не конвертируются в заказ: кликов >= 15, заказов = 0."
            elif (
                calc_roas is not None
                and calc_avg_order_value is not None
                and calc_cpa is not None
                and orders >= 3
                and calc_roas >= 5
                and drr <= 20
                and calc_cr >= 5
                and calc_cpa <= calc_avg_order_value * 0.25
            ):
                efficiency_label = "Можно масштабировать"
                efficiency_tone = "#DCEFE5"
                efficiency_reason = "Сильная экономика кампании: ROAS >= 5, ДРР <= 20%, CR >= 5%, заказов >= 3 и CPA не выше 25% среднего чека."
            elif (
                calc_roas is not None
                and orders >= 1
                and calc_roas >= 3
                and drr <= 35
                and calc_cr >= 2
            ):
                efficiency_label = "Рабочая кампания"
                efficiency_tone = "#F4EDDB"
                efficiency_reason = "Кампания окупается и приносит заказы, но ещё не дотягивает до уровня масштабирования."
            elif (
                (calc_roas is not None and calc_roas < 2)
                or drr > 45
                or calc_cr < 1.5
                or (calc_avg_order_value is not None and calc_cpa is not None and calc_cpa > calc_avg_order_value * 0.45)
            ):
                efficiency_label = "Нужна оптимизация"
                efficiency_tone = "#F8E9DD"
                efficiency_reason = "Слабая эффективность по одной или нескольким метрикам: низкий ROAS, высокий ДРР, слабый CR или слишком дорогой заказ."
            else:
                efficiency_label = "Под наблюдением"
                efficiency_tone = "#EEF2F8"
                efficiency_reason = "Есть сигналы эффективности, но данных или запаса по метрикам пока недостаточно для уверенного решения."

            campaign = dict(row_map)
            campaign.update(
                {
                    "__spend": spend,
                    "__impressions": impressions,
                    "__clicks": clicks,
                    "__orders": orders,
                    "__revenue": revenue,
                    "__ctr": ctr,
                    "__cpc": cpc,
                    "__drr": drr,
                    "__cpm": calc_cpm,
                    "__cpa": calc_cpa,
                    "__roas": calc_roas,
                    "__cr": calc_cr,
                    "__avg_order_value": calc_avg_order_value,
                    "__budget_burn": budget_burn,
                    "__daily_budget_burn": daily_budget_burn,
                    "__weekly_budget_burn": weekly_budget_burn,
                    "__efficiency_label": efficiency_label,
                    "__efficiency_tone": efficiency_tone,
                    "__efficiency_reason": efficiency_reason,
                }
            )
            campaigns.append(campaign)

            totals["spend"] += spend
            totals["impressions"] += impressions
            totals["clicks"] += clicks
            totals["orders"] += orders
            totals["revenue"] += revenue

        if not campaigns:
            return {}, []

        campaigns.sort(key=lambda item: (item.get("__spend") or 0.0), reverse=True)

        summary = {
            "campaigns_count": float(len(campaigns)),
            "spend": totals["spend"],
            "impressions": totals["impressions"],
            "clicks": totals["clicks"],
            "orders": totals["orders"],
            "revenue": totals["revenue"],
            "ctr": (totals["clicks"] / totals["impressions"] * 100.0) if totals["impressions"] > 0 else 0.0,
            "cpc": (totals["spend"] / totals["clicks"]) if totals["clicks"] > 0 else 0.0,
            "cpm": (totals["spend"] / totals["impressions"] * 1000.0) if totals["impressions"] > 0 else 0.0,
            "cr": (totals["orders"] / totals["clicks"] * 100.0) if totals["clicks"] > 0 else 0.0,
            "cpa": (totals["spend"] / totals["orders"]) if totals["orders"] > 0 else 0.0,
            "roas": (totals["revenue"] / totals["spend"]) if totals["spend"] > 0 else 0.0,
            "drr": (totals["spend"] / totals["revenue"] * 100.0) if totals["revenue"] > 0 else 0.0,
            "avg_order_value": (totals["revenue"] / totals["orders"]) if totals["orders"] > 0 else 0.0,
        }
        return summary, campaigns
    finally:
        wb.close()


def margin_defaults() -> tuple[str, str]:
    path = ROOT / "margin_settings.json"
    if not path.exists():
        return "0.25", "0.30"
    try:
        import json

        data = json.loads(path.read_text(encoding="utf-8"))
        return f"{float(data.get('min_margin', 0.25)):.2f}", f"{float(data.get('desired_margin', 0.30)):.2f}"
    except Exception:
        return "0.25", "0.30"


def valid_month_year(month: str, year: str) -> str | None:
    try:
        m, y = int(month), int(year)
    except ValueError:
        return "Месяц и год должны быть числами."
    if not 1 <= m <= 12:
        return "Месяц должен быть в диапазоне 1-12."
    if not 2000 <= y <= 2100:
        return "Год должен быть в диапазоне 2000-2100."
    return None


def valid_abc(a_m: str, a_y: str, b_m: str, b_y: str) -> str | None:
    error = valid_month_year(a_m, a_y) or valid_month_year(b_m, b_y)
    if error:
        return error
    if (int(a_y), int(a_m)) > (int(b_y), int(b_m)):
        return "Начало периода должно быть не позже конца."
    return None


def valid_balance(a: str, b: str) -> str | None:
    try:
        da = datetime.strptime(a, "%Y-%m-%d").date()
        db = datetime.strptime(b, "%Y-%m-%d").date()
    except ValueError:
        return "Даты должны быть в формате YYYY-MM-DD."
    if da > db:
        return "Дата начала не может быть позже даты окончания."
    if (db - da).days + 1 > 30:
        return "Период отчёта по балансу не может быть больше 30 дней."
    return None


def valid_margin(a: str, b: str) -> str | None:
    try:
        ma, mb = float(a.replace(",", ".")), float(b.replace(",", "."))
    except ValueError:
        return "Маржа должна быть числом в формате 0.25."
    if not (0 < ma < 1 and 0 < mb < 1):
        return "Значения маржи должны быть в диапазоне (0, 1)."
    if ma >= mb:
        return "Минимальная маржа должна быть меньше желаемой."
    return None


def valid_money(value: str, label: str) -> str | None:
    try:
        amount = float((value or "0").replace(",", "."))
    except ValueError:
        return f"{label} должен быть числом."
    if amount < 0:
        return f"{label} не может быть отрицательным."
    return None


def card(
    title: str,
    subtitle: str,
    body: list[ft.Control],
    tone: str = SURFACE,
    width: int | None = None,
    expand: bool = False,
    col: int | float | dict[str, int] = 12,
) -> ft.Control:
    header_controls: list[ft.Control] = []
    if title.strip():
        header_controls.append(ft.Text(title, size=18, weight=ft.FontWeight.W_700, color=TEXT))
    if subtitle.strip():
        header_controls.append(ft.Text(subtitle, size=12, color=MUTED))
    return ft.Container(
        width=width,
        expand=expand,
        col=col,
        padding=22,
        bgcolor=tone,
        border_radius=24,
        border=ft.Border.all(1, BORDER),
        shadow=ft.BoxShadow(blur_radius=22, spread_radius=0, color="#0F000000", offset=ft.Offset(0, 8)),
        content=ft.Column(
            [*header_controls, *body],
            spacing=10,
        ),
    )


def metric(
    label: str,
    value: str,
    caption: str,
    tone: str,
    width: int = 220,
    col: int | float = 3,
    tooltip: str | None = None,
) -> ft.Control:
    caption_controls: list[ft.Control] = []
    if caption.strip():
        caption_controls.append(ft.Text(caption, size=12, color=MUTED))
    return ft.Container(
        width=width,
        col=col,
        padding=18,
        bgcolor=tone,
        border_radius=22,
        border=ft.Border.all(1, "#E5E5EA"),
        tooltip=tooltip,
        content=ft.Column(
            [
                ft.Text(label, size=11, color=MUTED, weight=ft.FontWeight.W_600),
                ft.Text(value, size=22, weight=ft.FontWeight.W_700, color=TEXT),
                *caption_controls,
            ],
            spacing=3,
        ),
    )


def hero_metric(
    label: str,
    value: str,
    caption: str,
    tone: str,
    accent: str,
    width: int = 320,
    col: int | float = 4,
    tooltip: str | None = None,
) -> ft.Control:
    accent_controls: list[ft.Control] = []
    if caption.strip():
        accent_controls.append(
            ft.Row(
                [
                    ft.Container(width=10, height=10, bgcolor=accent, border_radius=10),
                    ft.Text(caption, size=12, color=MUTED),
                ],
                spacing=6,
            )
        )
    return ft.Container(
        width=width,
        col=col,
        padding=24,
        bgcolor=tone,
        border_radius=26,
        border=ft.Border.all(1, BORDER),
        tooltip=tooltip,
        content=ft.Column(
            [
                ft.Text(label, size=11, color=MUTED, weight=ft.FontWeight.W_700),
                ft.Text(value, size=32, weight=ft.FontWeight.W_800, color=TEXT),
                *accent_controls,
            ],
            spacing=5,
        ),
    )


def kv_row(label: str, value: str) -> ft.Control:
    return ft.Row(
        [
            ft.Text(label, expand=True, color=MUTED, size=12),
            ft.Text(value, color=TEXT, size=12, weight=ft.FontWeight.W_700),
        ],
        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
    )


def kv_row_compact(label: str, value: str, tooltip: str | None = None) -> ft.Control:
    return ft.Column(
        [
            ft.Text(label, color=MUTED, size=9, weight=ft.FontWeight.W_600),
            ft.Text(value, color=TEXT, size=13, weight=ft.FontWeight.W_700),
        ],
        spacing=0,
        tooltip=tooltip,
    )


def fact_chip(
    label: str,
    value: str,
    tone: str = "#FCFAF6",
    col: int | float = 6,
    tooltip: str | None = None,
    variant: str = "default",
) -> ft.Control:
    is_secondary = variant == "secondary"
    return ft.Container(
        col=col,
        padding=8 if is_secondary else 10,
        bgcolor=tone,
        border_radius=14 if is_secondary else 16,
        border=ft.Border.all(1, "#E5E5EA"),
        tooltip=tooltip,
        content=ft.Column(
            [
                ft.Text(label, color=MUTED, size=8 if is_secondary else 9, weight=ft.FontWeight.W_600),
                ft.Text(value, color=TEXT, size=12 if is_secondary else 13, weight=ft.FontWeight.W_700),
            ],
            spacing=1 if is_secondary else 2,
        ),
    )


def report_list_item(name: str, timestamp: str, on_open) -> ft.Control:
    return ft.Container(
        padding=ft.Padding.only(left=14, right=10, top=12, bottom=12),
        border_radius=18,
        bgcolor="#F7F7FA",
        border=ft.Border.all(1, "#E5E5EA"),
        content=ft.Row(
            [
                ft.Column(
                    [
                        ft.Text(name, size=13, color=TEXT, weight=ft.FontWeight.W_600),
                        ft.Text(timestamp, size=11, color=MUTED),
                    ],
                    spacing=2,
                    expand=True,
                ),
                ft.TextButton("Открыть", on_click=on_open),
            ],
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        ),
    )


def as_float(value) -> float | None:
    try:
        if value is None:
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def choose_metric_tone(
    value: float | None,
    *,
    good_at_least: float | None = None,
    warn_at_least: float | None = None,
    good_at_most: float | None = None,
    warn_at_most: float | None = None,
    dark: bool = False,
    neutral_light: str = "#EEF2F8",
    neutral_dark: str = "#293240",
) -> str:
    if value is None:
        return neutral_dark if dark else neutral_light
    if good_at_least is not None and value >= good_at_least:
        return "#214039" if dark else "#EEF7F4"
    if warn_at_least is not None and value >= warn_at_least:
        return "#454033" if dark else "#F8F5EC"
    if good_at_most is not None and value <= good_at_most:
        return "#214039" if dark else "#EEF7F4"
    if warn_at_most is not None and value <= warn_at_most:
        return "#454033" if dark else "#F8F5EC"
    return "#4A3236" if dark else "#FAECEA"


def semantic_surface(
    kind: str,
    *,
    dark: bool = False,
) -> str:
    palettes = {
        "neutral": ("#FBFCFF", "#20262B"),
        "neutral_soft": ("#F7F8FC", "#24292E"),
        "neutral_alt": ("#F3F5FA", "#272C31"),
        "positive": ("#EEF7F4", "#214039"),
        "warning": ("#F8F5EC", "#454033"),
        "danger": ("#FAECEA", "#4A3236"),
        "info": ("#F2F5FC", "#2C3744"),
    }
    light, dark_value = palettes.get(kind, palettes["neutral"])
    return dark_value if dark else light


def chart_definition(chart_kind: str) -> tuple[str, str, str, str]:
    if chart_kind == "revenue":
        return "Оборот по месяцам", "Выручка", "оборот", "#4FB59E"
    if chart_kind == "margin":
        return "Рентабельность по месяцам", "Net Margin", "рентабельность", "#D9B44A"
    return "Чистая прибыль по месяцам", "Чистая прибыль", "прибыль", "#E58A45"


def iter_period_range(from_month: int, from_year: int, to_month: int, to_year: int) -> list[tuple[int, int]]:
    periods: list[tuple[int, int]] = []
    cursor_month = from_month
    cursor_year = from_year
    while (cursor_year, cursor_month) <= (to_year, to_month):
        periods.append((cursor_month, cursor_year))
        cursor_month += 1
        if cursor_month == 13:
            cursor_month = 1
            cursor_year += 1
    return periods


def themed_color(color: str | None, dark: bool) -> str | None:
    if color is None:
        return None
    light_to_dark = {
        "#1D1D1F": "#F5F5F7",
        "#6E6E73": "#A1A1A6",
        "#1D2A2E": "#F3F4F6",
        "#6F7A76": "#B6C0BC",
        "#FFFDF8": "#232A2F",
        "#F4F1EA": "#1A1F23",
        "#FCFAF6": "#20262B",
        "#F7F7FA": "#262A2F",
        "#EAF3F0": "#243934",
        "#EAF0F5": "#27323D",
        "#F5EDE4": "#3A312B",
        "#D9EEE8": "#27443E",
        "#E8F1EE": "#24343A",
        "#EEF3E9": "#2A342D",
        "#F7F3EC": "#2A2F33",
        "#F5E3D8": "#4A302E",
        "#F5DDD8": "#4A302E",
        "#F6E7DA": "#43342C",
        "#E8EEF8": "#2B3440",
        "#EFE6F5": "#352D3C",
        "#E6F1EE": "#263632",
        "#F2EBDD": "#433A2E",
        "#F4E4E1": "#453131",
        "#F3F7F1": "#24342C",
        "#F8F1EA": "#3A312A",
        "#EEF2F8": "#293240",
        "#F1ECE3": "#38332C",
        "#F3F1EC": "#2C302B",
        "#F3F0E8": "#32302C",
        "#F2EDE7": "#332E2B",
        "#F1EEF8": "#2E3040",
        "#F0F4EA": "#2A342D",
        "#F8E9DD": "#433129",
        "#E8F3EE": "#1E3A32",
        "#F6E3E0": "#4A2626",
        "#DCEFE5": "#24443B",
        "#F4EDDB": "#4A422D",
        "#F6DFDB": "#4A302E",
        "#FBFCFF": "#20262B",
        "#F7F8FC": "#24292E",
        "#F3F5FA": "#272C31",
        "#EEF7F4": "#214039",
        "#F8F5EC": "#454033",
        "#FAECEA": "#4A3236",
        "#F2F5FC": "#2C3744",
        "#E6DED1": "#3A3F43",
        "#D2D2D7": "#3A3A3C",
        "#E5E5EA": "#3A3A3C",
        "#ECECF1": "#2C2C2E",
        "#F5F5F7": "#23262A",
        "#E7F2EF": "#233933",
        "#D6D0C4": "#3D474B",
        "#F7F7F8": "#212121",
        "#E9EAED": "#2F2F31",
        "#202123": "#ECECF1",
        "#1F1F22": "#ECECF1",
        "#D9D9DE": "#3A3A3C",
        "#FFFFFF": "#2B2B2D",
    }
    dark_to_light = {v: k for k, v in light_to_dark.items()}
    return light_to_dark.get(color, color) if dark else dark_to_light.get(color, color)


def main(page: ft.Page) -> None:
    page.title = "OzonReportX"
    page.theme_mode = ft.ThemeMode.DARK
    page.bgcolor = "#1A1F23"
    page.padding = 24
    try:
        page.window.maximized = True
        page.window.resizable = True
    except Exception:
        page.window_width = 1600
        page.window_height = 960
    try:
        page.window_resizable = True
    except Exception:
        pass
    page.window_min_width = 1200
    page.window_min_height = 860
    page.scroll = ft.ScrollMode.HIDDEN
    current_dark = {"value": True}

    today = date.today()
    prev = today.replace(day=1) - timedelta(days=1)
    years = [str(y) for y in range(today.year - 2, today.year + 2)]
    month_opts = [ft.dropdown.Option(str(i), f"{i}. {m}") for i, m in enumerate(MONTHS, start=1)]
    m_min, m_des = margin_defaults()

    monthly_m = ft.Dropdown(label="Месяц", options=month_opts, value=str(today.month), width=230)
    monthly_y = ft.Dropdown(label="Год", options=[ft.dropdown.Option(y) for y in years], value=str(today.year), width=170)
    dashboard_m = ft.Dropdown(label="Месяц сводки", options=month_opts, value=str(today.month), width=230)
    dashboard_y = ft.Dropdown(label="Год сводки", options=[ft.dropdown.Option(y) for y in years], value=str(today.year), width=170)
    dashboard_chart_type = ft.Dropdown(
        label="График",
        options=[
            ft.dropdown.Option("profit", "Прибыль / Месяц"),
            ft.dropdown.Option("revenue", "Оборот / Месяц"),
            ft.dropdown.Option("margin", "Рентабельность / Месяц"),
        ],
        value=None,
        width=250,
    )
    chart_from_month_value = prev.month - 4 if prev.month > 4 else prev.month + 8
    chart_from_year_value = prev.year if prev.month > 4 else prev.year - 1
    dashboard_chart_from_m = ft.Dropdown(label="С месяца", options=month_opts, value=str(chart_from_month_value), width=190)
    dashboard_chart_from_y = ft.Dropdown(label="Год", options=[ft.dropdown.Option(y) for y in years], value=str(chart_from_year_value), width=150)
    dashboard_chart_to_m = ft.Dropdown(label="По месяц", options=month_opts, value=str(prev.month), width=190)
    dashboard_chart_to_y = ft.Dropdown(label="Год", options=[ft.dropdown.Option(y) for y in years], value=str(prev.year), width=150)
    for chart_control in [
        dashboard_chart_type,
        dashboard_chart_from_m,
        dashboard_chart_from_y,
        dashboard_chart_to_m,
        dashboard_chart_to_y,
    ]:
        chart_control.color = TEXT
        chart_control.bgcolor = "#FBFCFF"
        chart_control.border_color = "#D9D9DE"
        chart_control.focused_border_color = PRIMARY
        chart_control.label_style = ft.TextStyle(size=11, color=MUTED)
    abc_fm = ft.Dropdown(label="Месяц от", options=month_opts, value=str(prev.month), width=190)
    abc_fy = ft.Dropdown(label="Год от", options=[ft.dropdown.Option(y) for y in years], value=str(prev.year), width=150)
    abc_tm = ft.Dropdown(label="Месяц до", options=month_opts, value=str(today.month), width=190)
    abc_ty = ft.Dropdown(label="Год до", options=[ft.dropdown.Option(y) for y in years], value=str(today.year), width=150)
    bal_from = ft.TextField(label="Дата от", value=today.replace(day=1).isoformat(), width=210)
    bal_to = ft.TextField(label="Дата до", value=today.isoformat(), width=210)
    margin_min = ft.TextField(label="Минимальная маржа", value=m_min, width=210)
    margin_des = ft.TextField(label="Желаемая маржа", value=m_des, width=210)
    state = ft.Text("Готово", color=PRIMARY, weight=ft.FontWeight.W_700, size=12)
    status_detail = ft.Text("Интерфейс готов к работе.", color=MUTED, size=9)
    status_indicator = ft.Container(width=8, height=8, border_radius=999, bgcolor=PRIMARY)
    spin = ft.ProgressRing(visible=False, width=18, height=18, stroke_width=2.4)
    cancel_button = ft.Button("Отмена", visible=False, bgcolor=DANGER, color="white", height=30)
    log = ft.ListView(spacing=6, auto_scroll=True, expand=True)
    log_selectable = ft.TextField(
        multiline=True,
        read_only=True,
        min_lines=30,
        max_lines=30,
        value="Интерфейс готов к работе.",
        border_radius=16,
        content_padding=ft.Padding.symmetric(horizontal=16, vertical=16),
    )
    log_store = SessionLog(ROOT / ".cache" / "logs" / f"ui-{datetime.now():%Y%m%d-%H%M%S}-{uuid4().hex}.log")
    log_store.reset("Интерфейс готов к работе.")
    log_state = {"lines": log_store.lines}
    dashboard_freshness = ft.Text("Актуальность данных: —", size=12, color=MUTED)
    dashboard_summary = ft.Column(spacing=10)
    dashboard_metric_wrap = ft.ResponsiveRow(run_spacing=14, spacing=14)
    dashboard_details = ft.Column(spacing=12)
    dashboard_promotion = ft.Column(spacing=12)
    dashboard_chart = ft.Image(src="", visible=False, border_radius=18)
    dashboard_chart_note = ft.Text("Выберите метрику и период, затем постройте график.", size=11, color=MUTED, visible=False)
    dashboard_chart_state_indicator = ft.Container(width=40, height=40, border_radius=14, alignment=ft.Alignment(0, 0))
    dashboard_chart_state_title = ft.Text("Постройте график", size=15, weight=ft.FontWeight.W_700, color=TEXT)
    dashboard_chart_state_caption = ft.Text("Выберите метрику и период, затем постройте график.", size=12, color=MUTED)
    dashboard_chart_state_card = ft.Container(
        padding=18,
        border_radius=22,
        bgcolor=semantic_surface("neutral_soft", dark=False),
        border=ft.Border.all(1, BORDER),
        content=ft.Row(
            [
                dashboard_chart_state_indicator,
                ft.Column(
                    [dashboard_chart_state_title, dashboard_chart_state_caption],
                    spacing=4,
                    expand=True,
                ),
            ],
            spacing=14,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        ),
    )
    dashboard_chart_meta_title = ft.Text("", size=14, weight=ft.FontWeight.W_700, color=TEXT)
    dashboard_chart_meta_caption = ft.Text("", size=11, color=MUTED)
    dashboard_chart_meta_card = ft.Container(
        visible=False,
        padding=14,
        border_radius=20,
        bgcolor=semantic_surface("neutral_soft", dark=False),
        border=ft.Border.all(1, BORDER),
        content=ft.Row(
            [
                ft.Container(width=10, height=10, border_radius=999, bgcolor=PRIMARY),
                ft.Column([dashboard_chart_meta_title, dashboard_chart_meta_caption], spacing=3, expand=True),
            ],
            spacing=10,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        ),
    )
    dashboard_chart_panel = ft.Column(
        [dashboard_chart_meta_card, dashboard_chart_state_card, dashboard_chart, dashboard_chart_note],
        spacing=10,
    )
    dashboard_chart_state = {"built": False, "mode": "empty"}
    chart_pending_action = {"type": None}
    active_proc: dict[str, subprocess.Popen | None] = {"proc": None}
    job_lock = threading.Lock()
    job_cancel_event = threading.Event()
    cancel_state = {"requested": False, "title": ""}
    status_ui_ready = {"value": False}
    monthly_reports_list = ft.Column(spacing=10, scroll=ft.ScrollMode.AUTO, height=360)
    abc_reports_list = ft.Column(spacing=10, scroll=ft.ScrollMode.AUTO, height=360)
    files_col = ft.Column(spacing=14)
    pricing_costs_status = ft.Text("Читаю лист «Основной»...", size=11, color=MUTED, weight=ft.FontWeight.W_600)
    pricing_costs_status_shell = ft.Container(
        padding=ft.Padding.symmetric(horizontal=12, vertical=8),
        border_radius=999,
        bgcolor=semantic_surface("neutral_alt", dark=False),
        border=ft.Border.all(1, "#E5E5EA"),
        content=pricing_costs_status,
    )
    pricing_costs_list = ft.Column(spacing=8, scroll=ft.ScrollMode.AUTO)
    pricing_costs_save_button = ft.Button("Сохранить изменения", bgcolor=PRIMARY, color="white", disabled=True)
    pricing_costs_editors: dict[int, dict[str, ft.TextField]] = {}
    pricing_costs_dirty = {"value": False, "revision": 0}
    pricing_costs_snapshot = {"signature": None}
    ai_diag_title = ft.Text("Локальная модель готова к работе", size=18, weight=ft.FontWeight.W_800, color=TEXT)
    ai_diag_subtitle = ft.Text("Проверка Ollama и модели выполняется локально на устройстве пользователя.", size=12, color=MUTED)
    ai_status_dot = ft.Container(width=12, height=12, border_radius=999, bgcolor="#D9B44A")
    ai_status_label = ft.Text("Проверка статуса AI...", size=12, weight=ft.FontWeight.W_700, color=TEXT)
    ai_runtime_host = ft.Text(f"Ollama host: {OLLAMA_HOST}", size=11, color=MUTED)
    ai_runtime_model = ft.Text(f"Модель: {OLLAMA_MODEL}", size=11, color=MUTED)
    ai_runtime_note = ft.Text("AI-ассистент работает локально и не отправляет отчёты во внешний API.", size=12, color=MUTED)
    ai_model_select = ft.Dropdown(label="Модель", width=180, value=OLLAMA_MODEL, options=[])
    ai_phase_label = ft.Text("Текущий этап: idle", size=12, weight=ft.FontWeight.W_800, color=TEXT)
    ai_phase_detail = ft.Text("Модель ожидает новый запрос.", size=12, color=MUTED)
    ai_trace = ft.Column(spacing=8, scroll=ft.ScrollMode.AUTO, height=210)
    ai_messages = ft.ListView(spacing=18, auto_scroll=True, expand=True)
    ai_empty_state = ft.Container(
        alignment=ft.Alignment(0, 0),
        expand=True,
        content=ft.Container(
            width=760,
            padding=28,
            border_radius=32,
            bgcolor="#FFFFFF",
            border=ft.Border.all(1, "#E5E5EA"),
            content=ft.Column(
                [
                    ft.Text("Чем помочь по Ozon?", size=30, weight=ft.FontWeight.W_900, color=TEXT, text_align=ft.TextAlign.CENTER),
                    ft.Text(
                        "Спросите про прибыль, расходы, риски или конкретный месяц. Ассистент сам подтянет нужные данные из локальных отчётов.",
                        size=13,
                        color=MUTED,
                        text_align=ft.TextAlign.CENTER,
                        width=560,
                    ),
                ],
                spacing=10,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            ),
        ),
    )
    ai_messages.controls.append(ai_empty_state)
    ai_input = ft.TextField(
        hint_text="Спросите про прибыль, расходы, риски...",
        multiline=True,
        min_lines=1,
        max_lines=6,
        shift_enter=True,
        border_radius=24,
        content_padding=ft.Padding.symmetric(horizontal=18, vertical=16),
        expand=True,
    )
    ai_send_button = ft.IconButton(icon=ft.Icons.ARROW_UPWARD_ROUNDED, tooltip="Отправить")
    ai_reset_button = ft.TextButton("Очистить контекст")
    ai_refresh_button = ft.TextButton("Проверить Ollama")
    ai_tools_button = ft.PopupMenuButton(content=ft.Text("Инструменты", size=12, weight=ft.FontWeight.W_700, color=TEXT))
    ai_new_chat_button = ft.TextButton("Новый диалог")
    ai_top_title = ft.Text("AI-ассистент", size=18, weight=ft.FontWeight.W_800, color=TEXT)
    ai_top_subtitle = ft.Text("Локальный анализ на Ollama", size=12, color=MUTED)
    ai_quick_month_button = ft.TextButton("Разобрать месяц")
    ai_quick_pressure_button = ft.TextButton("Что давит на прибыль")
    ai_quick_brief_button = ft.TextButton("Краткий разбор")
    ai_quick_risk_button = ft.TextButton("Риски месяца")
    ai_busy_label = ft.Text("Ассистент готов.", size=12, color=MUTED)
    ai_busy_ring = ft.ProgressRing(width=16, height=16, stroke_width=2.4, visible=False)
    ai_view: ft.Container | None = None
    ai_header_card: ft.Container | None = None
    ai_composer_card: ft.Container | None = None
    ai_state = {
        "busy": False,
        "conversation_history": [],
        "session_memory": {
            "selected_period": None,
            "selected_artikul": None,
            "last_comparison_periods": [],
        },
        "available": False,
        "status_message": "",
        "selected_model": OLLAMA_MODEL,
        "available_models": [],
        "request_seq": 0,
        "active_request_id": None,
        "cancelled_request_ids": set(),
        "phase": "idle",
        "trace": [],
        "debug_enabled": False,
        "stream_controls": {},
        "request_started_at": {},
    }
    ai_tools = AITools(ROOT) if AITools else None

    folder_specs = [
        FolderSpec("Месячные отчёты", ROOT / "reports"),
        FolderSpec("ABC/XYZ аналитика", ROOT / "ABC&XYZ reports"),
        FolderSpec("Поставки", ROOT / "stocks reports"),
        FolderSpec("Баланс", ROOT / "balance reports"),
    ]

    def toast(msg: str, error: bool = False) -> None:
        set_status("Ошибка" if error else "Готово", msg, DANGER if error else PRIMARY, busy=False)
        snack = ft.SnackBar(ft.Text(msg), bgcolor=DANGER if error else PRIMARY, open=True)
        page.overlay.append(snack)
        page.update()

    def _log_line_control(line: str) -> ft.Text:
        return ft.Text(
            line if line else " ",
            size=12,
            color="#F3F4F6" if current_dark["value"] else TEXT,
            selectable=True,
            font_family="Consolas",
        )

    def _sync_log_theme() -> None:
        line_color = "#F3F4F6" if current_dark["value"] else TEXT
        for control in log.controls:
            if isinstance(control, ft.Text):
                control.color = line_color
        log_selectable.color = line_color
        log_selectable.bgcolor = "#232A2F" if current_dark["value"] else "#FFFDF8"
        log_selectable.border_color = "#3D474B" if current_dark["value"] else "#D6D0C4"

    def _render_log_lines(lines: list[str]) -> None:
        payload = lines or ["Команда завершилась без вывода."]
        log.controls = [_log_line_control(line) for line in payload]
        log_selectable.value = "\n".join(payload)
        _sync_log_theme()

    def set_log(text: str, persist: bool = True) -> None:
        normalized = (text or "Команда завершилась без вывода.").replace("\r\n", "\n").replace("\r", "\n")
        log_store.reset(normalized, persist=persist)
        _render_log_lines(list(log_store.lines))

    def append_log(text: str) -> None:
        normalized = (text or "").replace("\r\n", "\n").replace("\r", "\n")
        extra_lines = normalized.split("\n") if normalized else [""]
        log_store.append(normalized)
        log.controls.extend(_log_line_control(line) for line in extra_lines[-500:])
        log.controls = log.controls[-500:]
        log_selectable.value = "\n".join(log_store.lines)

    def get_full_log_text() -> str:
        return log_store.full_text().strip() or "Лог пока пуст."

    def set_status(title: str, detail: str = "", color: str | None = None, busy: bool = False) -> None:
        tone = color or PRIMARY
        state.value = title
        state.color = tone
        detail_text = (detail or "").replace("\n", " ").strip()
        if len(detail_text) > 52:
            detail_text = detail_text[:49].rstrip() + "..."
        status_detail.value = detail_text
        status_indicator.bgcolor = tone
        if tone == DANGER:
            status_shell.bgcolor = semantic_surface("danger", dark=current_dark["value"])
            status_shell.border = ft.Border.all(1, "#CFA8A1" if not current_dark["value"] else "#5A3B40")
        elif busy:
            status_shell.bgcolor = semantic_surface("info", dark=current_dark["value"])
            status_shell.border = ft.Border.all(1, "#C8D4E5" if not current_dark["value"] else "#344455")
        elif tone == PRIMARY:
            status_shell.bgcolor = semantic_surface("positive", dark=current_dark["value"])
            status_shell.border = ft.Border.all(1, "#C8DED6" if not current_dark["value"] else "#2C5047")
        else:
            status_shell.bgcolor = "#1F252A" if current_dark["value"] else "#FBFCFF"
            status_shell.border = ft.Border.all(1, "#353D42" if current_dark["value"] else "#E1E5EE")
        spin.visible = busy
        cancel_button.visible = busy or active_proc["proc"] is not None
        if status_ui_ready["value"] and page.controls:
            page.update(status_shell, status_indicator, state, status_detail, spin, cancel_button)

    def ai_surface_tokens() -> dict[str, str]:
        dark = current_dark["value"]
        return {
            "canvas": "#1D1D1F" if dark else "#F5F5F7",
            "composer": "#26282C" if dark else "#FFFFFF",
            "composer_border": "#383C41" if dark else "#D9D9DE",
            "text": "#F5F5F7" if dark else "#1D1D1F",
            "muted": "#A1A1A6" if dark else "#6E6E73",
            "assistant": "#F5F5F7" if dark else "#1D1D1F",
            "assistant_bg": "#26282C" if dark else "#FFFFFF",
            "assistant_border": "#383C41" if dark else "#E5E5EA",
            "user_bg": "#303236" if dark else "#ECEEF3",
            "user_text": "#F5F5F7" if dark else "#1D1D1F",
            "toolbar": "#26282C" if dark else "#FFFFFF",
            "toolbar_border": "#383C41" if dark else "#E5E7EB",
        }

    def ai_centered(control: ft.Control, width: int = 1040) -> ft.Control:
        return ft.Row(
            [ft.Container(content=control, width=width)],
            alignment=ft.MainAxisAlignment.CENTER,
        )

    def ai_markdown(text: str) -> ft.Control:
        dark = current_dark["value"]
        return ft.Markdown(
            value=text or "",
            selectable=True,
            extension_set=ft.MarkdownExtensionSet.GITHUB_FLAVORED,
            code_theme=ft.MarkdownCodeTheme.A11Y_DARK if dark else ft.MarkdownCodeTheme.A11Y_LIGHT,
            soft_line_break=True,
            shrink_wrap=True,
            fit_content=True,
        )

    def ai_message_bubble(role: str, text: str, note: str | None = None) -> ft.Control:
        tokens = ai_surface_tokens()
        is_user = role == "user"
        if is_user:
            return ai_centered(
                ft.Row(
                    [
                        ft.Container(
                            padding=ft.Padding.symmetric(horizontal=18, vertical=14),
                            border_radius=22,
                            bgcolor=tokens["user_bg"],
                            data="ai_user_bubble",
                            content=ft.Text(text, size=14, color=tokens["user_text"], selectable=True, data="ai_user_text"),
                        )
                    ],
                    alignment=ft.MainAxisAlignment.END,
                )
            )
        return ai_centered(
            ft.Container(
                padding=ft.Padding.symmetric(horizontal=22, vertical=18),
                border_radius=24,
                bgcolor=tokens["assistant_bg"],
                border=ft.Border.all(1, tokens["assistant_border"]),
                data="ai_assistant_bubble",
                content=ft.Column(
                    [ai_markdown(text)]
                    + ([ft.Text(note, size=11, color=tokens["muted"], data="ai_note_text")] if note else []),
                    spacing=10,
                    horizontal_alignment=ft.CrossAxisAlignment.START,
                ),
            )
        )

    def ai_context_state() -> str | None:
        lines: list[str] = []
        memory = ai_state["session_memory"]
        if memory["selected_period"]:
            lines.append(f"selected_period={memory['selected_period']}")
        if memory["selected_artikul"]:
            lines.append(f"selected_artikul={memory['selected_artikul']}")
        if memory["last_comparison_periods"]:
            periods = ", ".join(memory["last_comparison_periods"])
            lines.append(f"last_comparison_periods={periods}")
        return "\n".join(lines) if lines else None

    def ai_append_message(role: str, text: str, note: str | None = None) -> None:
        if ai_messages.controls and ai_messages.controls[0] is ai_empty_state:
            ai_messages.controls.clear()
        ai_messages.controls.append(ai_message_bubble(role, text, note))

    def ai_format_elapsed(seconds: float | None) -> str | None:
        if seconds is None:
            return None
        if seconds < 1:
            return f"Ответ за {seconds * 1000:.0f} мс"
        if seconds < 10:
            return f"Ответ за {seconds:.1f} с"
        return f"Ответ за {seconds:.0f} с"

    def ai_append_streaming_placeholder(note: str | None = None) -> dict[str, ft.Control]:
        if ai_messages.controls and ai_messages.controls[0] is ai_empty_state:
            ai_messages.controls.clear()
        tokens = ai_surface_tokens()
        text_control = ft.Text("", size=15, color=tokens["assistant"], selectable=True)
        content_host = ft.Container(
            padding=ft.Padding.symmetric(horizontal=22, vertical=18),
            border_radius=24,
            bgcolor=tokens["assistant_bg"],
            border=ft.Border.all(1, tokens["assistant_border"]),
            data="ai_assistant_bubble",
            content=text_control,
        )
        note_control = ft.Text(note or "", size=11, color=tokens["muted"], data="ai_note_text")
        thinking_current = ft.Text(
            "",
            size=12,
            color=tokens["muted"],
            selectable=True,
        )
        thinking_anchor = ft.Text("", size=1, color=tokens["muted"])
        thinking_list = ft.ListView(
            controls=[thinking_current, thinking_anchor],
            spacing=3,
            auto_scroll=True,
            height=66,
            width=956,
        )
        thinking_toggle = ft.TextButton("Развернуть", visible=False)
        thinking_title = ft.Text("Ход рассуждения", size=11, weight=ft.FontWeight.W_700, color=tokens["muted"])
        thinking_wrap = ft.Container(
            visible=False,
            width=1040,
            padding=ft.Padding.symmetric(horizontal=14, vertical=12),
            border_radius=20,
            bgcolor=tokens["assistant_bg"],
            border=ft.Border.all(1, tokens["assistant_border"]),
            data="ai_thinking_wrap",
            content=ft.Column(
                [
                    ft.Row(
                        [thinking_title, thinking_toggle],
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    ),
                    thinking_list,
                ],
                spacing=6,
            ),
        )

        def toggle_thinking(_e=None) -> None:
            expanded = bool(getattr(thinking_toggle, "data", False))
            thinking_toggle.data = not expanded
            thinking_toggle.text = "Свернуть" if not expanded else "Развернуть"
            thinking_list.height = 220 if not expanded else 66
            if status_ui_ready["value"] and page.controls:
                page.update(thinking_wrap)

        thinking_toggle.on_click = toggle_thinking
        assistant_control = ai_centered(
            ft.Column(
                [thinking_wrap, content_host, note_control],
                spacing=8,
                horizontal_alignment=ft.CrossAxisAlignment.START,
            )
        )
        ai_messages.controls.append(assistant_control)
        return {
            "text": text_control,
            "content_host": content_host,
            "note": note_control,
            "thinking_list": thinking_list,
            "thinking_current": thinking_current,
            "thinking_anchor": thinking_anchor,
            "thinking_toggle": thinking_toggle,
            "thinking_wrap": thinking_wrap,
        }

    def ai_has_conversation() -> bool:
        return not (len(ai_messages.controls) == 1 and ai_messages.controls[0] is ai_empty_state)

    def ai_apply_chat_theme() -> None:
        tokens = ai_surface_tokens()
        if ai_view is not None:
            ai_view.bgcolor = tokens["canvas"]
        if ai_composer_card is not None:
            ai_composer_card.bgcolor = tokens["composer"]
            ai_composer_card.border = ft.Border.all(1, tokens["composer_border"])
        if ai_header_card is not None:
            ai_header_card.bgcolor = tokens["assistant_bg"]
            ai_header_card.border = ft.Border.all(1, tokens["assistant_border"])
        ai_send_button.icon_color = "white"
        ai_send_button.style = ft.ButtonStyle(
            bgcolor=PRIMARY,
            shape=ft.CircleBorder(),
            padding=ft.Padding.all(12),
        )

        for ctrl in [ai_refresh_button, ai_reset_button, ai_new_chat_button]:
            ctrl.style = ft.ButtonStyle(
                color=tokens["text"],
                bgcolor="transparent",
                shape=ft.RoundedRectangleBorder(radius=999),
                padding=ft.Padding.symmetric(horizontal=14, vertical=9),
            )

        ai_input.bgcolor = tokens["composer"]
        ai_input.border_color = tokens["composer_border"]
        ai_input.focused_border_color = PRIMARY
        ai_input.color = tokens["text"]
        ai_input.cursor_color = PRIMARY
        ai_input.hint_style = ft.TextStyle(color=tokens["muted"], size=14)
        ai_model_select.color = tokens["text"]
        ai_model_select.border_color = "transparent"
        ai_model_select.focused_border_color = PRIMARY
        ai_model_select.label_style = ft.TextStyle(color=tokens["muted"], size=11)
        ai_model_select.bgcolor = "transparent"
        ai_tools_button.content = ft.Text("Инструменты", size=12, weight=ft.FontWeight.W_700, color=tokens["text"])
        ai_tools_button.menu_bgcolor = "#22272B" if current_dark["value"] else "#FBFCFF"
        ai_top_title.color = tokens["text"]
        ai_top_subtitle.color = tokens["muted"]
        ai_busy_label.color = tokens["muted"]
        ai_diag_subtitle.color = tokens["muted"]
        ai_empty_state.content.bgcolor = tokens["assistant_bg"]
        ai_empty_state.content.border = ft.Border.all(1, tokens["assistant_border"])
        if isinstance(ai_empty_state.content, ft.Container) and isinstance(ai_empty_state.content.content, ft.Column):
            for child in ai_empty_state.content.content.controls:
                if isinstance(child, ft.Text):
                    child.color = tokens["text"] if child.size and child.size >= 20 else tokens["muted"]

        def refresh_ai_control_theme(control) -> None:
            if control is None:
                return
            if isinstance(control, ft.Container):
                if control.data == "ai_user_bubble":
                    control.bgcolor = tokens["user_bg"]
                    control.border = None
                elif control.data in {"ai_assistant_bubble", "ai_thinking_wrap"}:
                    control.bgcolor = tokens["assistant_bg"]
                    if control.border is not None:
                        control.border = ft.Border.all(1, tokens["assistant_border"])
            if isinstance(control, ft.Text):
                if control.data == "ai_user_text":
                    control.color = tokens["user_text"]
                elif control.data == "ai_note_text":
                    control.color = tokens["muted"]
            for attr in ("content", "title", "subtitle", "leading", "trailing"):
                if hasattr(control, attr):
                    refresh_ai_control_theme(getattr(control, attr))
            for attr in ("controls", "actions"):
                if hasattr(control, attr):
                    for child in getattr(control, attr) or []:
                        refresh_ai_control_theme(child)

        refresh_ai_control_theme(ai_messages)
        for payload in ai_state["stream_controls"].values():
            if isinstance(payload, dict):
                thinking_list = payload.get("thinking_list")
                if isinstance(thinking_list, ft.ListView):
                    for control in thinking_list.controls:
                        if isinstance(control, ft.Text):
                            control.color = tokens["muted"]

    def ai_refresh_tools_menu() -> None:
        debug_label = "Debug режим ✓" if ai_state["debug_enabled"] else "Debug режим"
        menu_text_color = "#F5F5F7" if current_dark["value"] else "#1D1D1F"
        ai_tools_button.items = [
            ft.PopupMenuItem(
                content=ft.Row(
                    [
                        ft.Icon(ft.Icons.CHECK if ai_state["debug_enabled"] else ft.Icons.BUG_REPORT_OUTLINED, size=18, color=menu_text_color),
                        ft.Text(debug_label, size=14, color=menu_text_color, weight=ft.FontWeight.W_600),
                    ],
                    spacing=12,
                ),
                on_click=lambda _e: ai_toggle_debug(),
            ),
            ft.PopupMenuItem(
                content=ft.Row(
                    [
                        ft.Icon(ft.Icons.SYNC, size=18, color=menu_text_color),
                        ft.Text("Проверить Ollama", size=14, color=menu_text_color, weight=ft.FontWeight.W_600),
                    ],
                    spacing=12,
                ),
                on_click=lambda _e: ai_refresh_status(),
            ),
            ft.PopupMenuItem(
                content=ft.Row(
                    [
                        ft.Icon(ft.Icons.DELETE_OUTLINE, size=18, color=menu_text_color),
                        ft.Text("Очистить контекст", size=14, color=menu_text_color, weight=ft.FontWeight.W_600),
                    ],
                    spacing=12,
                ),
                on_click=ai_reset,
            ),
        ]
        if status_ui_ready["value"] and page.controls:
            page.update(ai_tools_button)

    def ai_toggle_debug() -> None:
        ai_state["debug_enabled"] = not ai_state["debug_enabled"]
        ai_refresh_tools_menu()
        if ai_state["debug_enabled"]:
            append_log("[AI DEBUG] Debug режим включён.")
            toast("Debug режим включён. Лог доступен по F2.")
        else:
            toast("Debug режим выключен.")

    def ai_trace_row(title: str, detail: str) -> ft.Control:
        return ft.Container(
            padding=12,
            border_radius=16,
            bgcolor="#FFFDF8" if not current_dark["value"] else "#232A2F",
            border=ft.Border.all(1, "#D6D0C4" if not current_dark["value"] else "#3D474B"),
            content=ft.Column(
                [
                    ft.Text(title, size=12, weight=ft.FontWeight.W_800, color=TEXT),
                    ft.Text(detail, size=11, color=MUTED, max_lines=4, overflow=ft.TextOverflow.ELLIPSIS),
                ],
                spacing=4,
            ),
        )

    def ai_set_phase(phase: str, detail: str) -> None:
        ai_state["phase"] = phase
        ai_phase_label.value = f"Текущий этап: {phase}"
        ai_phase_detail.value = detail
        if status_ui_ready["value"] and page.controls:
            page.update(ai_phase_label, ai_phase_detail)

    def ai_add_trace(title: str, detail: str) -> None:
        clean_detail = (detail or "").strip() or "—"
        ai_state["trace"].append((title, clean_detail))
        if len(ai_state["trace"]) > 10:
            ai_state["trace"] = ai_state["trace"][-10:]
        ai_trace.controls = [ai_trace_row(item_title, item_detail) for item_title, item_detail in ai_state["trace"]]
        if status_ui_ready["value"] and page.controls:
            page.update(ai_trace)

    def ai_apply_selected_model(model_name: str | None) -> None:
        model = (model_name or "").strip() or ai_state["selected_model"] or OLLAMA_MODEL
        ai_state["selected_model"] = model
        ai_runtime_model.value = f"Модель: {model}"
        ai_model_select.value = model
        if ai_chat_module is not None:
            ai_chat_module.OLLAMA_MODEL = model

    def ai_debug_log(title: str, payload: object = "") -> None:
        if not ai_state["debug_enabled"]:
            return
        if isinstance(payload, (dict, list)):
            body = json.dumps(payload, ensure_ascii=False, indent=2)
        elif payload is None:
            body = ""
        else:
            body = str(payload)
        block = f"[AI DEBUG] {title}"
        if body:
            block += f"\n{body}"
        summary = body.splitlines()[0][:180] if body else "Событие без деталей"
        title_map = {
            "Новый запрос": ("prepare", "Собираю payload и фиксирую активную модель."),
            "Контекст сессии": ("context", "Собираю контекст диалога и выбранный период."),
            "Этап A: запрос к модели": ("plan", "Модель решает, нужны ли данные из инструментов."),
            "DATA_REQUEST": ("tools", "AI запросил данные из локальных инструментов."),
            "Результаты инструментов": ("tools-ready", "Инструменты вернули данные для анализа."),
            "Сформированный data block": ("compose", "Подготавливаю данные для финального ответа модели."),
            "Этап B: финальный запрос к модели": ("answer", "Модель строит итоговый аналитический ответ."),
            "Финальный ответ для UI": ("done", "Ответ готов и будет показан в чате."),
            "Пользователь отменил AI-запрос": ("cancelled", "Текущий AI-запрос отменён пользователем."),
        }
        phase_payload = title_map.get(title)
        ui_queue.put(("ai_step", title, summary, phase_payload))
        ui_queue.put(("log_append", block))

    def ai_load_installed_models() -> tuple[list[str], str | None]:
        try:
            response = requests.get(f"{OLLAMA_HOST}/api/tags", timeout=5)
            response.raise_for_status()
            data = response.json()
        except requests.exceptions.RequestException as exc:
            return [], f"Ollama недоступен по адресу {OLLAMA_HOST}: {exc}"

        models = []
        for item in data.get("models", []) or []:
            if isinstance(item, dict) and item.get("name"):
                models.append(str(item["name"]))
        models = sorted(set(models))
        return models, None

    def ai_refresh_status(push_update: bool = True) -> None:
        selected_from_ui = (ai_model_select.value or "").strip()
        if selected_from_ui:
            ai_apply_selected_model(selected_from_ui)
        if AI_IMPORT_ERROR:
            ai_state["available"] = False
            ai_state["status_message"] = f"Не удалось загрузить AI-модуль: {AI_IMPORT_ERROR}"
            ai_model_select.options = []
        elif not check_ollama_model:
            ai_state["available"] = False
            ai_state["status_message"] = "AI-модуль недоступен."
        else:
            models, error = ai_load_installed_models()
            ai_state["available_models"] = models
            ai_model_select.options = [ft.dropdown.Option(model, model) for model in models]
            if error:
                ai_state["available"] = False
                ai_state["status_message"] = error
            else:
                current_model = ai_state["selected_model"]
                if models:
                    if current_model not in models:
                        ai_apply_selected_model(models[0])
                    else:
                        ai_apply_selected_model(current_model)
                ok, message = check_ollama_model()
                ai_state["available"] = ok
                ai_state["status_message"] = message if message else "Ollama запущен, модель установлена, локальный анализ доступен."

        ai_model_select.disabled = ai_state["busy"] or not bool(ai_state["available_models"])
        if ai_state["available_models"]:
            ai_runtime_note.value = f"Установлено моделей: {len(ai_state['available_models'])}. Можно переключать без перезапуска приложения."
        else:
            ai_runtime_note.value = "AI-ассистент работает локально и не отправляет отчёты во внешний API."

        ai_diag_title.value = "Локальная модель готова к работе" if ai_state["available"] else "Требуется внимание к AI-окружению"
        ai_diag_subtitle.value = ai_state["status_message"]
        ai_status_dot.bgcolor = "#4FB59E" if ai_state["available"] else DANGER
        ai_status_label.value = "Статус: готово" if ai_state["available"] else "Статус: требуется настройка"
        if push_update and status_ui_ready["value"] and page.controls:
            page.update(ai_diag_title, ai_diag_subtitle, ai_status_dot, ai_status_label, ai_runtime_host, ai_runtime_model, ai_runtime_note, ai_model_select)

    def ai_set_busy(busy: bool, detail: str) -> None:
        ai_state["busy"] = busy
        ai_busy_ring.visible = busy
        ai_busy_label.value = detail
        ai_send_button.disabled = busy
        ai_input.disabled = busy
        ai_reset_button.disabled = busy
        ai_refresh_button.disabled = busy
        ai_model_select.disabled = busy or not bool(ai_state["available_models"])
        if not busy and ai_state["phase"] not in ("done", "cancelled"):
            ai_set_phase("idle", "Модель ожидает новый запрос.")
        if status_ui_ready["value"] and page.controls:
            page.update(ai_busy_ring, ai_busy_label, ai_send_button, ai_input, ai_reset_button, ai_refresh_button, ai_model_select)

    def ai_reset(_e=None) -> None:
        if ai_state["active_request_id"] is not None:
            ai_state["cancelled_request_ids"].add(ai_state["active_request_id"])
            ai_state["request_started_at"].pop(ai_state["active_request_id"], None)
        ai_state["active_request_id"] = None
        ai_state["conversation_history"] = []
        ai_state["session_memory"] = {
            "selected_period": None,
            "selected_artikul": None,
            "last_comparison_periods": [],
        }
        ai_state["trace"] = []
        ai_trace.controls = []
        ai_messages.controls = [ai_empty_state]
        ai_sync_layout()
        ai_set_phase("idle", "Контекст и диагностическая лента очищены.")
        ai_busy_label.value = "Контекст очищен. Ассистент готов к новому диалогу."
        if status_ui_ready["value"] and page.controls:
            page.update(ai_body, ai_messages, ai_messages_shell, ai_welcome_shell, ai_composer_shell, ai_busy_label, ai_trace, ai_phase_label, ai_phase_detail)

    def ai_model_changed(e: ft.ControlEvent) -> None:
        selected = (e.control.value or "").strip()
        if not selected:
            return
        ai_apply_selected_model(selected)
        ai_busy_label.value = f"Активная модель: {selected}"
        if status_ui_ready["value"] and page.controls:
            page.update(ai_runtime_model, ai_busy_label, ai_model_select)

    def ai_submit(prompt_text: str) -> None:
        prompt = (prompt_text or "").strip()
        if not prompt or ai_state["busy"]:
            return
        ai_apply_selected_model(ai_model_select.value or ai_state["selected_model"])
        ai_state["request_seq"] += 1
        request_id = ai_state["request_seq"]
        ai_state["active_request_id"] = request_id
        ai_state["request_started_at"][request_id] = time.perf_counter()
        runtime_model = ai_chat_module.OLLAMA_MODEL if ai_chat_module is not None else ai_state["selected_model"]
        ai_debug_log("Новый запрос", {"request_id": request_id, "prompt": prompt, "model": ai_state["selected_model"], "runtime_model": runtime_model, "host": OLLAMA_HOST})
        ai_append_message("user", prompt)
        ai_input.value = ""
        ai_sync_layout()
        ai_set_busy(True, "Ассистент читает вопрос и строит план анализа.")
        set_status("AI анализирует запрос", prompt, ACCENT, busy=True)
        page.update(ai_body, ai_messages, ai_messages_shell, ai_welcome_shell, ai_composer_shell, ai_input, state, status_detail, spin, cancel_button)

        def worker() -> None:
            def is_cancelled() -> bool:
                return request_id in ai_state["cancelled_request_ids"] or ai_state["active_request_id"] != request_id

            def merge_tool_results(base: dict, incoming: dict) -> dict:
                merged = dict(base)
                for key, value in (incoming or {}).items():
                    if key not in merged:
                        merged[key] = value
                    else:
                        existing = merged[key]
                        if isinstance(existing, list):
                            existing.append(value)
                        else:
                            merged[key] = [existing, value]
                return merged

            if not ai_state["available"] or not ai_tools or not chat_with_ai or not parse_ai_response:
                ai_debug_log("AI недоступен", ai_state["status_message"] or "AI недоступен.")
                ui_queue.put(("ai_finish", request_id, prompt, None, ai_state["status_message"] or "AI недоступен."))
                return

            try:
                ai_apply_selected_model(ai_model_select.value or ai_state["selected_model"])
                if is_cancelled():
                    ai_debug_log("AI запрос отменён до старта", {"request_id": request_id})
                    return
                context_state = ai_context_state()
                ai_debug_log("Контекст сессии", context_state or "(пусто)")
                final_answer = None
                final_runtime = {}
                accumulated_tool_results = {}
                max_steps = 4

                for step_idx in range(1, max_steps + 1):
                    workflow_state = build_workflow_state(prompt, accumulated_tool_results) if build_workflow_state else None
                    has_tool_data = bool(accumulated_tool_results)
                    phase_label = f"Шаг {step_idx}"
                    if workflow_state and workflow_state.name == "yearly_artikul_profit_report":
                        if workflow_state.stage == "planning":
                            phase_note = "Определяю годовой workflow..."
                        elif workflow_state.stage == "periods_resolved":
                            phase_note = "Собираю месячные отчёты по workflow..."
                        elif workflow_state.stage == "aggregation_ready":
                            phase_note = "Формирую итог по агрегированной сводке..."
                        else:
                            phase_note = "Обрабатываю workflow..."
                    else:
                        phase_note = "Планирую ответ..." if not has_tool_data else ("Собираю дополнительные данные..." if step_idx < max_steps else "Формирую ответ...")
                    phase_temperature = 0.7 if not has_tool_data else 0.2
                    current_data_block = format_tool_results(
                        accumulated_tool_results,
                        user_message=prompt,
                        workflow_state=workflow_state,
                    ) if has_tool_data else None
                    deterministic_needs = get_workflow_followup_needs(prompt, accumulated_tool_results) if get_workflow_followup_needs else []
                    if deterministic_needs:
                        ai_debug_log(
                            f"{phase_label}: workflow step",
                            {
                                "workflow": workflow_state.name if workflow_state else None,
                                "workflow_stage": workflow_state.stage if workflow_state else None,
                                "needs": deterministic_needs,
                            },
                        )
                        if has_tool_data:
                            ai_debug_log(f"{phase_label}: workflow data block", current_data_block)
                        tool_results = execute_tools(ai_tools, deterministic_needs)
                        if is_cancelled():
                            ai_debug_log("AI запрос отменён после workflow tool-вызовов", {"request_id": request_id, "step": step_idx})
                            return
                        ai_debug_log("Результаты инструментов", tool_results)
                        accumulated_tool_results = merge_tool_results(accumulated_tool_results, tool_results)

                        periods_found = []
                        for need in deterministic_needs:
                            args = need.get("args", {})
                            if "period" in args:
                                periods_found.append(args["period"])
                            if "artikul" in args:
                                ai_state["session_memory"]["selected_artikul"] = args["artikul"]
                        if periods_found:
                            ai_state["session_memory"]["selected_period"] = periods_found[-1]
                            if len(periods_found) > 1:
                                ai_state["session_memory"]["last_comparison_periods"] = periods_found
                        continue

                    ai_debug_log(
                        f"{phase_label}: model step",
                        {
                            "temperature": phase_temperature,
                            "max_tokens": "unlimited",
                            "streaming": True,
                            "has_data_block": has_tool_data,
                            "workflow": workflow_state.name if workflow_state else None,
                            "workflow_stage": workflow_state.stage if workflow_state else None,
                            "force_final_answer": workflow_state.force_final_answer if workflow_state else False,
                        },
                    )
                    if has_tool_data:
                        ai_debug_log(f"{phase_label}: model data block", current_data_block)

                    if stream_chat_with_ai:
                        ui_queue.put(("ai_stream_start", request_id, phase_note))
                        ai_response, _, runtime_info = stream_chat_with_ai(
                            prompt,
                            ai_state["conversation_history"],
                            data_block=current_data_block,
                            context_state=context_state,
                            temperature=phase_temperature,
                            max_tokens=None,
                            allow_followup_data_requests=not (workflow_state.force_final_answer if workflow_state else False),
                            on_chunk=None,
                            on_thinking=lambda chunk: ui_queue.put(("ai_thinking_chunk", request_id, chunk)),
                            should_cancel=is_cancelled,
                        )
                    else:
                        ai_response, _, runtime_info = chat_with_ai(
                            prompt,
                            ai_state["conversation_history"],
                            data_block=current_data_block,
                            context_state=context_state,
                            temperature=phase_temperature,
                            max_tokens=None,
                            allow_followup_data_requests=not (workflow_state.force_final_answer if workflow_state else False),
                        )

                    final_runtime = runtime_info or final_runtime
                    if is_cancelled():
                        ai_debug_log("AI запрос отменён во время многошагового цикла", {"request_id": request_id, "step": step_idx})
                        return
                    if not ai_response:
                        ai_debug_log("Пустой ответ от модели", runtime_info or {})
                        ui_queue.put(("ai_finish", request_id, prompt, None, "Не удалось получить ответ от Ollama."))
                        return

                    ai_debug_log(f"Сырой ответ модели, {phase_label}", ai_response)
                    ai_debug_log(f"Runtime, {phase_label}", runtime_info or {})
                    parsed = parse_ai_response(ai_response)
                    ai_debug_log(f"Распарсенный ответ, {phase_label}", parsed or "(не удалось распарсить JSON)")

                    if parsed and parsed.get("type") == "FINAL_ANSWER":
                        final_answer = parsed.get("answer", ai_response)
                        break

                    if parsed and parsed.get("type") == "DATA_REQUEST":
                        needs = parsed.get("needs", [])
                        ai_debug_log("DATA_REQUEST", {"step": step_idx, "reason": parsed.get("reason", ""), "needs": needs})
                        tool_results = execute_tools(ai_tools, needs)
                        if is_cancelled():
                            ai_debug_log("AI запрос отменён после tool-вызовов", {"request_id": request_id, "step": step_idx})
                            return
                        ai_debug_log("Результаты инструментов", tool_results)
                        accumulated_tool_results = merge_tool_results(accumulated_tool_results, tool_results)

                        periods_found = []
                        for need in needs:
                            args = need.get("args", {})
                            if "period" in args:
                                periods_found.append(args["period"])
                            if "artikul" in args:
                                ai_state["session_memory"]["selected_artikul"] = args["artikul"]
                        if periods_found:
                            ai_state["session_memory"]["selected_period"] = periods_found[-1]
                            if len(periods_found) > 1:
                                ai_state["session_memory"]["last_comparison_periods"] = periods_found
                        continue

                    if workflow_state and workflow_state.force_final_answer and build_workflow_fallback_answer:
                        fallback_answer = build_workflow_fallback_answer(prompt, accumulated_tool_results)
                        final_answer = fallback_answer or ai_response
                    else:
                        final_answer = ai_response
                    break

                if final_answer is None and accumulated_tool_results:
                    fallback_answer = build_workflow_fallback_answer(prompt, accumulated_tool_results) if build_workflow_fallback_answer else None
                    if fallback_answer:
                        final_answer = fallback_answer
                    else:
                        ui_queue.put(("ai_finish", request_id, prompt, None, "Модель не завершила анализ финальным ответом за допустимое число шагов."))
                        return
                if final_answer is None:
                    final_answer = "Не удалось завершить AI-анализ."

                if is_cancelled():
                    ai_debug_log("AI запрос отменён перед публикацией результата", {"request_id": request_id})
                    return
                ai_debug_log("Финальный ответ для UI", final_answer)
                ui_queue.put(("ai_finish", request_id, prompt, final_answer, final_runtime.get("model", ai_state["selected_model"])))
            except Exception as exc:
                ai_debug_log("Исключение AI-пайплайна", repr(exc))
                ui_queue.put(("ai_finish", request_id, prompt, None, f"Ошибка AI-анализа: {exc}"))

        page.run_thread(worker)

    def reveal_path(path: Path, label: str | None = None) -> None:
        target_label = label or path.name or str(path)
        set_status("Открытие файла", target_label, PRIMARY, busy=False)
        open_path(path)

    def cancel_current_action(_e=None) -> None:
        proc = active_proc["proc"]
        if ai_state["busy"] and ai_state["active_request_id"] is not None:
            request_id = ai_state["active_request_id"]
            ai_state["cancelled_request_ids"].add(request_id)
            ai_state["request_started_at"].pop(request_id, None)
            ai_state["active_request_id"] = None
            ai_set_busy(False, "AI-запрос отменён. Можно отправить новый вопрос.")
            ai_append_message("assistant", "Текущий AI-запрос остановлен пользователем.", "Запрос отменён")
            ai_debug_log("Пользователь отменил AI-запрос", {"request_id": request_id})
            set_status("AI остановлен", "Текущий запрос отменён пользователем.", DANGER, busy=False)
            if status_ui_ready["value"] and page.controls:
                page.update(ai_messages, ai_busy_label, state, status_detail, spin, cancel_button)
            return
        if proc is None and not job_lock.locked():
            set_status("Нет задачи", "Нечего отменять.", MUTED, busy=False)
            return
        cancel_state["requested"] = True
        job_cancel_event.set()
        set_status("Отмена...", cancel_state["title"] or "Останавливаем задачу.", DANGER, busy=True)
        if prompt_dialog is not None and prompt_dialog.open:
            prompt_dialog.open = False
            holder = prompt_state["holder"]
            event = prompt_state["event"]
            if isinstance(holder, dict):
                holder["value"] = "0"
            if isinstance(event, threading.Event):
                event.set()
            if page.controls:
                page.update(prompt_dialog)
        if proc is None:
            return
        try:
            if os.name == "nt":
                subprocess.run(
                    ["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                    capture_output=True,
                    text=True,
                    timeout=10,
                )
            else:
                proc.terminate()
        except Exception:
            try:
                proc.kill()
            except Exception:
                pass

    prompt_label = ft.Text("", color=TEXT, size=14, weight=ft.FontWeight.W_700)
    prompt_hint = ft.Text("", color=MUTED, size=12)
    prompt_input = ft.TextField(label="Сумма маркетинга", value="0", autofocus=True)
    prompt_context = ft.Text("", color=MUTED, size=12)
    prompt_state: dict[str, object | None] = {"holder": None, "event": None}
    prompt_dialog: ft.AlertDialog | None = None
    ui_queue: queue.Queue = queue.Queue()

    def set_dashboard_chart_feedback(kind: str, title: str, caption: str = "", meta_caption: str = "") -> None:
        dashboard_chart_state["mode"] = kind
        icon_specs = {
            "empty": (ft.Icons.AUTO_GRAPH_OUTLINED, semantic_surface("info", dark=False), "#4A6FA5"),
            "loading": (ft.Icons.HOURGLASS_TOP_ROUNDED, semantic_surface("neutral_alt", dark=False), PRIMARY),
            "error": (ft.Icons.ERROR_OUTLINE_ROUNDED, semantic_surface("danger", dark=False), DANGER),
            "ready": (ft.Icons.CHECK_CIRCLE_OUTLINE_ROUNDED, semantic_surface("positive", dark=False), PRIMARY),
        }
        icon_name, card_tone, accent_color = icon_specs.get(kind, icon_specs["empty"])
        dashboard_chart_state_indicator.bgcolor = card_tone
        dashboard_chart_state_indicator.content = (
            ft.ProgressRing(width=18, height=18, stroke_width=2.2, color=accent_color)
            if kind == "loading"
            else ft.Icon(icon_name, size=20, color=accent_color)
        )
        dashboard_chart_state_card.bgcolor = semantic_surface("neutral_soft", dark=False)
        dashboard_chart_state_card.border = ft.Border.all(1, BORDER)
        dashboard_chart_state_title.value = title
        dashboard_chart_state_caption.value = caption
        dashboard_chart_state_card.visible = kind != "ready"
        dashboard_chart_meta_card.visible = kind == "ready"
        if kind == "ready":
            dashboard_chart_meta_title.value = title
            dashboard_chart_meta_caption.value = meta_caption or caption
        if current_dark["value"] and "apply_control_theme" in locals():
            apply_control_theme(dashboard_chart_panel, True)

    set_dashboard_chart_feedback("empty", "Постройте график", "Выберите метрику и период, затем постройте график.")

    def infer_prompt_period(source_line: str) -> str | None:
        line = source_line or ""
        month_match = re.search(r"\b(0?[1-9]|1[0-2])\b", line)
        year_match = re.search(r"\b(20\d{2})\b", line)
        if month_match and year_match:
            month_num = int(month_match.group(1))
            if 1 <= month_num <= 12:
                return f"{MONTHS[month_num - 1]} {year_match.group(1)}"

        named_month_pattern = "|".join(MONTHS)
        named_match = re.search(rf"({named_month_pattern})\s+(20\d{{2}})", line, flags=re.IGNORECASE)
        if named_match:
            month_name = named_match.group(1)
            year = named_match.group(2)
            normalized = next((m for m in MONTHS if m.lower() == month_name.lower()), month_name)
            return f"{normalized} {year}"

        if cancel_state["title"] == "Месячный отчёт":
            month_value = int(monthly_m.value or today.month)
            year_value = int(monthly_y.value or today.year)
            return f"{MONTHS[month_value - 1]} {year_value}"

        if cancel_state["title"] == "Обновление бизнес-сводки":
            month_value = int(dashboard_m.value or today.month)
            year_value = int(dashboard_y.value or today.year)
            return f"{MONTHS[month_value - 1]} {year_value}"

        return None

    def apply_prompt_dialog_theme() -> None:
        dark = current_dark["value"]
        prompt_input.color = "#F3F4F6" if dark else TEXT
        prompt_input.bgcolor = "#232A2F" if dark else "#FFFDF8"
        prompt_input.border_color = "#3D474B" if dark else "#D6D0C4"
        prompt_input.focused_border_color = PRIMARY
        prompt_input.label_style = ft.TextStyle(color="#B6C0BC" if dark else MUTED, size=12)
        prompt_label.color = "#F3F4F6" if dark else TEXT
        prompt_hint.color = "#B6C0BC" if dark else MUTED
        prompt_context.color = "#B6C0BC" if dark else MUTED
        if prompt_dialog is not None:
            prompt_dialog.bgcolor = "#232A2F" if dark else "#FFFDF8"
            prompt_dialog.title = ft.Text(
                "Нужны расходы на маркетинг",
                color="#F3F4F6" if dark else TEXT,
                weight=ft.FontWeight.W_800,
            )

    def apply_discount_dialog_theme() -> None:
        dark = current_dark["value"]
        divider_color = "#3D474B" if dark else "#D6D0C4"
        if isinstance(discount_dialog_state.get("plan"), dict):
            render_discount_dialog()
        discount_dialog_badge.color = "#B6C0BC" if dark else MUTED
        discount_dialog_status.color = "#B6C0BC" if dark else MUTED
        if discount_dialog is not None:
            discount_dialog.bgcolor = "#232A2F" if dark else "#FFFDF8"
            discount_dialog.title = ft.Text(
                "Заявки на скидку",
                color="#F3F4F6" if dark else TEXT,
                weight=ft.FontWeight.W_800,
            )
            content = discount_dialog.content
            if isinstance(content, ft.Container) and isinstance(content.content, ft.Column):
                for child in content.content.controls:
                    if isinstance(child, ft.Divider):
                        child.color = divider_color

    def request_prompt(label: str, source_line: str) -> str:
        holder: dict[str, str] = {"value": "0"}
        event = threading.Event()
        ui_queue.put(("prompt", label, source_line, holder, event))
        while not event.wait(timeout=0.1):
            if job_cancel_event.is_set():
                return "0"
        return holder["value"]

    def collect_dashboard_chart_data(
        from_month: int,
        from_year: int,
        to_month: int,
        to_year: int,
    ) -> tuple[list[str], list[float], list[float], list[float], float]:
        periods = iter_period_range(from_month, from_year, to_month, to_year)
        labels: list[str] = []
        revenue_values: list[float] = []
        profit_values: list[float] = []
        margin_values: list[float] = []
        latest_source_mtime = 0.0
        for month_value, year_value in periods:
            period_report_path = report_path_for_period(month_value, year_value)
            metrics, _built = read_business_metrics(period_report_path)
            if not metrics:
                continue
            try:
                latest_source_mtime = max(latest_source_mtime, period_report_path.stat().st_mtime)
            except OSError:
                pass
            revenue = as_float(metrics.get("Общая выручка"))
            profit = as_float(metrics.get("Чистая прибыль"))
            margin = as_float(metrics.get("Рентабельность по чистой прибыли (Net Profit Margin) %"))
            if revenue is None and profit is None and margin is None:
                continue
            labels.append(f"{MONTHS[month_value - 1][:3]} {str(year_value)[-2:]}")
            revenue_values.append(revenue or 0.0)
            profit_values.append(profit or 0.0)
            margin_values.append(margin or 0.0)
        return labels, revenue_values, profit_values, margin_values, latest_source_mtime

    def render_dashboard_chart(from_month: int, from_year: int, to_month: int, to_year: int) -> None:
        labels, revenue_values, profit_values, margin_values, latest_source_mtime = collect_dashboard_chart_data(from_month, from_year, to_month, to_year)

        if not labels:
            dashboard_chart.visible = False
            dashboard_chart_note.value = "В выбранном диапазоне нет отчётов с бизнес-показателями."
            dashboard_chart_state["built"] = False
            set_dashboard_chart_feedback("empty", "Нет данных для графика", "В выбранном диапазоне пока нет готовых отчётов с бизнес-показателями.")
            return

        fig_bg = "#20262B"
        ax_bg = "#232A2F"
        text_color = "#F3F4F6"
        muted_color = "#B6C0BC"
        grid_color = "#3D474B"
        chart_kind = dashboard_chart_type.value or "profit"
        chart_dir = ROOT / ".cache"
        chart_dir.mkdir(exist_ok=True)

        plt.close("all")
        fig, ax = plt.subplots(figsize=(9.6, 4.4), dpi=140)
        fig.patch.set_facecolor(fig_bg)
        ax.set_facecolor(ax_bg)
        title, legend_label, note_label, line_color = chart_definition(chart_kind)
        if chart_kind == "revenue":
            series = revenue_values
        elif chart_kind == "margin":
            series = margin_values
        else:
            series = profit_values
        ax.plot(labels, series, color=line_color, linewidth=2.8, marker="o", label=legend_label)
        ax.grid(axis="y", color=grid_color, alpha=0.35, linewidth=0.8)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color(grid_color)
        ax.spines["bottom"].set_color(grid_color)
        ax.tick_params(axis="x", colors=muted_color, labelsize=9)
        ax.tick_params(axis="y", colors=muted_color, labelsize=9)
        ax.set_title(title, loc="left", color=text_color, fontsize=13, pad=12)
        legend = ax.legend(
            frameon=False,
            loc="upper center",
            bbox_to_anchor=(0.5, -0.18),
            ncol=1,
            borderaxespad=0.0,
        )
        for text in legend.get_texts():
            text.set_color(text_color)
        fig.subplots_adjust(bottom=0.26, top=0.86, left=0.07, right=0.98)
        cache_stamp = int(latest_source_mtime) if latest_source_mtime else int(datetime.now().timestamp())
        chart_render_path = chart_dir / f"dashboard_trend_{chart_kind}_{cache_stamp}.png"
        fig.savefig(chart_render_path, facecolor=fig_bg, bbox_inches="tight")
        plt.close(fig)

        dashboard_chart.src = str(chart_render_path.resolve())
        dashboard_chart.visible = True
        dashboard_chart.width = 980
        dashboard_chart.height = 420
        dashboard_chart_note.value = f"Период: {MONTHS[from_month - 1]} {from_year} - {MONTHS[to_month - 1]} {to_year}. Метрика: {note_label}."
        dashboard_chart_state["built"] = True
        set_dashboard_chart_feedback(
            "ready",
            title,
            f"{MONTHS[from_month - 1]} {from_year} - {MONTHS[to_month - 1]} {to_year}",
            f"Метрика: {note_label}. Период: {MONTHS[from_month - 1]} {from_year} - {MONTHS[to_month - 1]} {to_year}.",
        )

    def open_interactive_dashboard_chart(_e=None) -> None:
        chart_error = valid_dashboard_chart_period()
        if chart_error:
            toast(chart_error, True)
            return
        missing_periods = get_missing_chart_reports()
        if missing_periods:
            build_missing_chart_reports(missing_periods, "interactive")
            return
        from_month = int(dashboard_chart_from_m.value or prev.month)
        from_year = int(dashboard_chart_from_y.value or prev.year)
        to_month = int(dashboard_chart_to_m.value or prev.month)
        to_year = int(dashboard_chart_to_y.value or prev.year)
        labels, revenue_values, profit_values, margin_values, latest_source_mtime = collect_dashboard_chart_data(from_month, from_year, to_month, to_year)
        if not labels:
            set_dashboard_chart_feedback("empty", "Нет данных для интерактивного графика", "Сначала подготовьте отчёты за выбранный диапазон.")
            toast("Для интерактивного графика пока нет данных.", True)
            return
        try:
            import plotly.graph_objects as go
        except Exception:
            toast("Для интерактивного графика нужен пакет plotly. Установите зависимости.", True)
            return

        chart_kind = dashboard_chart_type.value or "profit"
        title, legend_label, note_label, line_color = chart_definition(chart_kind)
        if chart_kind == "revenue":
            series = revenue_values
            hover_template = "%{x}<br>Выручка: %{y:,.0f} ₽<extra></extra>"
        elif chart_kind == "margin":
            series = margin_values
            hover_template = "%{x}<br>Рентабельность: %{y:.2f}%<extra></extra>"
        else:
            series = profit_values
            hover_template = "%{x}<br>Чистая прибыль: %{y:,.0f} ₽<extra></extra>"

        paper_bg = "#1A1F23"
        plot_bg = "#232A2F"
        text_color = "#F3F4F6"
        grid_color = "#3D474B"

        fig = go.Figure()
        fig.add_trace(
            go.Scatter(
                x=labels,
                y=series,
                mode="lines+markers",
                name=legend_label,
                line={"color": line_color, "width": 3},
                marker={"size": 10, "color": line_color},
                hovertemplate=hover_template,
            )
        )
        fig.update_layout(
            title={"text": title, "x": 0.01, "xanchor": "left"},
            paper_bgcolor=paper_bg,
            plot_bgcolor=plot_bg,
            font={"color": text_color},
            hovermode="x",
            margin={"l": 50, "r": 24, "t": 60, "b": 80},
            legend={"orientation": "h", "yanchor": "top", "y": -0.18, "xanchor": "center", "x": 0.5},
        )
        fig.update_xaxes(showgrid=False, color=text_color)
        fig.update_yaxes(gridcolor=grid_color, zerolinecolor=grid_color, color=text_color)

        chart_dir = ROOT / ".cache"
        chart_dir.mkdir(exist_ok=True)
        cache_stamp = int(latest_source_mtime) if latest_source_mtime else int(datetime.now().timestamp())
        html_path = chart_dir / f"dashboard_interactive_{chart_kind}_{cache_stamp}.html"
        fig.write_html(str(html_path), include_plotlyjs="cdn", full_html=True)
        open_path(html_path)
        dashboard_chart_note.value = f"Открыт интерактивный график: {note_label}."
        set_dashboard_chart_feedback(
            "ready",
            title,
            f"{MONTHS[from_month - 1]} {from_year} - {MONTHS[to_month - 1]} {to_year}",
            f"Интерактивный график открыт. Метрика: {note_label}.",
        )
        set_status("График открыт", f"{MONTHS[from_month - 1]} {from_year} - {MONTHS[to_month - 1]} {to_year}", PRIMARY, busy=False)
        if page.controls:
            page.update(dashboard_chart_panel)

    def refresh_dashboard() -> None:
        selected_month = int(dashboard_m.value or today.month)
        selected_year = int(dashboard_y.value or today.year)
        report_path = report_path_for_period(selected_month, selected_year)
        set_status("Чтение отчёта", f"{report_path.name}", ACCENT, busy=True)
        try:
            metrics, built_at = read_business_metrics(report_path)
            promotion_summary, promotion_campaigns = read_campaign_metrics(report_path)
        except Exception as exc:
            set_status("Ошибка чтения отчёта", str(exc), DANGER, busy=False)
            toast(f"Не удалось прочитать {report_path.name}: {exc}", True)
            return
        period_title = f"{MONTHS[selected_month - 1]} {selected_year}"
        # Always build dashboard cards from the light base palette.
        # Dark mode is applied afterward via apply_control_theme(), which keeps
        # theme switching instant and prevents stale dark surfaces in light mode.
        dark_ui = False
        business_metric_help = {
            "Выручка": "Общая выручка за период по месячному отчёту. Основной показатель оборота.",
            "Чистая прибыль": "Прибыль после вычета себестоимости, комиссий, логистики, маркетинга и прочих расходов.",
            "Себестоимость": "Итоговая товарная себестоимость реализованных заказов за период.",
            "Net Margin": "Рентабельность по чистой прибыли. Формула: чистая прибыль / выручка × 100%.",
            "Средний чек": "Формула: выручка / количество заказов.",
            "Всего заказов": "Общее количество заказов за период.",
            "Доставляются": "Количество заказов в статусе доставки на момент формирования отчёта.",
            "Доставлено": "Количество успешно доставленных заказов.",
            "Отменено": "Количество отменённых заказов. Рост показателя требует проверки причин отмен.",
            "Возвраты": "Количество возвращённых заказов.",
            "Продвижение Ozon": "Расходы на внутреннее продвижение на Ozon за период.",
            "Внешний маркетинг": "Затраты на внешний маркетинг вне Ozon, введённые при сборке отчёта.",
            "Звёздные товары": "Расходы по программе продвижения 'Звёздные товары' за период.",
            "Хранение FBO": "Затраты на хранение товаров на складах Ozon по схеме FBO.",
            "Комиссии Ozon": "Комиссии маркетплейса. В карточке показаны доля в выручке и абсолютная сумма.",
            "Логистика": "Логистические расходы Ozon. В карточке показаны доля в выручке и абсолютная сумма.",
            "COGS": "Валовая прибыль после вычета себестоимости товара, до прочих операционных расходов.",
            "Gross Margin": "Рентабельность по валовой прибыли. Формула: валовая прибыль / выручка × 100%.",
            "Опер. расходы": "Операционные расходы периода, влияющие на чистую прибыль.",
            "В доставке": "Количество заказов, которые ещё находятся в логистической цепочке.",
            "Выплачено Ozon": "Реальный перевод от Ozon за период — из отчёта о балансе (Финансы → Баланс), а не расчётная величина.",
            "Начислено Ozon": "Сумма начислений Ozon за период по календарным дням (отчёт о балансе). Сравните со строкой «Наш расчёт» — методика группировки разная, поэтому небольшое расхождение нормально.",
            "Ранний вывод": "Комиссия Ozon за досрочный вывод средств (если вы им пользовались в этом периоде).",
        }

        if not metrics:
            dashboard_freshness.value = "Актуальность данных: нет отчёта за выбранный период"
            set_status("Нет отчёта", period_title, DANGER, busy=False)
            dashboard_summary.controls = [
                ft.ResponsiveRow(
                    run_spacing=12,
                    spacing=12,
                    controls=[
                        card(
                            "Нет отчёта",
                            "",
                            [
                                ft.Text("Отчёт за этот период не найден в папке reports.", size=12, color=MUTED),
                                ft.Button("Сформировать отчёт", bgcolor=PRIMARY, color="white", on_click=lambda _e: launch_dashboard_report()),
                            ],
                            tone=semantic_surface("danger", dark=dark_ui),
                            width=320,
                            col=4,
                        ),
                    ],
                )
            ]
            dashboard_metric_wrap.controls = [
                metric("Выручка", "—", "", PRIMARY_SOFT, width=155, col=2),
                metric("Чистая прибыль", "—", "", "#F6E7DA", width=155, col=2),
                metric("Net Margin", "—", "", "#E8EEF8", width=155, col=2),
                metric("Средний чек", "—", "", "#EFE6F5", width=155, col=2),
                metric("Доставляются", "—", "", "#EEF3E9", width=155, col=2),
            ]
            dashboard_details.controls = [
                card("Что дальше", "", [ft.Text("Выберите период и обновите данные.", size=12, color=MUTED)], tone=semantic_surface("neutral_soft", dark=dark_ui))
            ]
            dashboard_promotion.controls = []
            if current_dark["value"]:
                apply_control_theme(dashboard_summary, True)
                apply_control_theme(dashboard_metric_wrap, True)
                apply_control_theme(dashboard_details, True)
                apply_control_theme(dashboard_promotion, True)
            if page.controls:
                page.update(dashboard_summary, dashboard_metric_wrap, dashboard_details, dashboard_promotion, dashboard_freshness, dashboard_chart_panel)
            return

        dashboard_freshness.value = f"Актуальность данных: {built_at or '—'}"
        set_status("Данные загружены", period_title, PRIMARY, busy=False)
        dashboard_summary.controls = []
        net_margin_value = as_float(metrics.get("Рентабельность по чистой прибыли (Net Profit Margin) %")) or 0.0
        profit_value = as_float(metrics.get("Чистая прибыль")) or 0.0
        margin_tone = (
            semantic_surface("positive", dark=dark_ui)
            if net_margin_value >= 20
            else semantic_surface("warning", dark=dark_ui)
            if net_margin_value >= 5
            else semantic_surface("danger", dark=dark_ui)
        )
        margin_accent = PRIMARY if net_margin_value >= 20 else ACCENT if net_margin_value >= 5 else DANGER
        profit_tone = semantic_surface("positive", dark=dark_ui) if profit_value >= 0 else semantic_surface("danger", dark=dark_ui)
        profit_accent = PRIMARY if profit_value >= 0 else DANGER
        dashboard_metric_wrap.controls = [
            hero_metric("Выручка", _format_currency(metrics.get("Общая выручка")), "", semantic_surface("neutral", dark=dark_ui), PRIMARY, width=330, col=4, tooltip=business_metric_help["Выручка"]),
            hero_metric("Чистая прибыль", _format_currency(metrics.get("Чистая прибыль")), "", profit_tone, profit_accent, width=330, col=4, tooltip=business_metric_help["Чистая прибыль"]),
            hero_metric("Себестоимость", _format_currency(abs(as_float(metrics.get("Итоговая себестоимость")) or 0.0)), "", semantic_surface("neutral_soft", dark=dark_ui), ACCENT, width=330, col=4, tooltip=business_metric_help["Себестоимость"]),
            metric("Net Margin", _format_percent(metrics.get("Рентабельность по чистой прибыли (Net Profit Margin) %")), "", margin_tone, width=150, col=2, tooltip=business_metric_help["Net Margin"]),
            metric("Средний чек", _format_currency(metrics.get("Средний чек")), "", semantic_surface("neutral_soft", dark=dark_ui), width=150, col=2, tooltip=business_metric_help["Средний чек"]),
            metric("Заказы", _format_number(metrics.get("Общее количество заказов")), "", semantic_surface("neutral_alt", dark=dark_ui), width=150, col=2, tooltip=business_metric_help["Всего заказов"]),
            metric("В пути", _format_number(metrics.get("Количество заказов в доставке")), "", semantic_surface("neutral_alt", dark=dark_ui), width=150, col=2, tooltip=business_metric_help["Доставляются"]),
            metric("Доставлено", _format_number(metrics.get("Количество доставленных заказов")), "", semantic_surface("neutral_soft", dark=dark_ui), width=150, col=2, tooltip=business_metric_help["Доставлено"]),
            metric("Отменено", _format_number(metrics.get("Количество отменённых заказов")), "", semantic_surface("danger", dark=dark_ui), width=150, col=2, tooltip=business_metric_help["Отменено"]),
        ]
        dashboard_details.controls = [
            ft.ResponsiveRow(
                run_spacing=12,
                spacing=12,
                controls=[
                    card(
                        "Продажи",
                        "",
                        [
                            kv_row_compact("Заказы", _format_number(metrics.get("Общее количество заказов")), tooltip=business_metric_help["Всего заказов"]),
                            kv_row_compact("Доставлено", _format_number(metrics.get("Количество доставленных заказов")), tooltip=business_metric_help["Доставлено"]),
                            kv_row_compact("Отменено", _format_number(metrics.get("Количество отменённых заказов")), tooltip=business_metric_help["Отменено"]),
                            kv_row_compact("Возвраты", _format_number(metrics.get("Количество возвращённых заказов")), tooltip=business_metric_help["Возвраты"]),
                        ],
                        tone=semantic_surface("neutral", dark=dark_ui),
                        width=290,
                        col=4,
                    ),
                    card(
                        "Прибыль",
                        "",
                        [
                            kv_row_compact("Чистая прибыль", _format_currency(metrics.get("Чистая прибыль")), tooltip=business_metric_help["Чистая прибыль"]),
                            kv_row_compact("Net Margin", _format_percent(metrics.get("Рентабельность по чистой прибыли (Net Profit Margin) %")), tooltip=business_metric_help["Net Margin"]),
                            kv_row_compact("Себестоимость", _format_currency(abs(as_float(metrics.get("Итоговая себестоимость")) or 0.0)), tooltip=business_metric_help["Себестоимость"]),
                            kv_row_compact("COGS", _format_currency(metrics.get("COGS (валовая прибыль)")), tooltip=business_metric_help["COGS"]),
                            kv_row_compact("Опер. расходы", _format_currency(metrics.get("Операционные расходы")), tooltip=business_metric_help["Опер. расходы"]),
                        ],
                        tone=semantic_surface("neutral", dark=dark_ui),
                        width=290,
                        col=4,
                    ),
                    card(
                        "Расходы",
                        "",
                        [
                            kv_row_compact("Продвижение Ozon", _format_currency(metrics.get("Продвижение Ozon")), tooltip=business_metric_help["Продвижение Ozon"]),
                            kv_row_compact("Комиссии Ozon", f"{_format_percent(metrics.get('Комиссии Ozon %'))} • {_format_currency(metrics.get('Комиссии Ozon сумма'))}", tooltip=business_metric_help["Комиссии Ozon"]),
                            kv_row_compact("Логистика", f"{_format_percent(metrics.get('Логистика %'))} • {_format_currency(metrics.get('Логистика сумма'))}", tooltip=business_metric_help["Логистика"]),
                            kv_row_compact("Хранение FBO", _format_currency(metrics.get("Расход хранения FBO")), tooltip=business_metric_help["Хранение FBO"]),
                            kv_row_compact("Внешний маркетинг", _format_currency(metrics.get("Внешний маркетинг")), tooltip=business_metric_help["Внешний маркетинг"]),
                        ],
                        tone=semantic_surface("neutral_alt", dark=dark_ui),
                        width=290,
                        col=4,
                    ),
                    card(
                        "Выплаты Ozon",
                        "Из отчёта о балансе — реальные переводы, а не наш расчёт.",
                        [
                            kv_row_compact("Выплачено", _format_currency(metrics.get("Выплачено Ozon за период (реальный перевод)")), tooltip=business_metric_help["Выплачено Ozon"]),
                            kv_row_compact("Начислено (Ozon)", _format_currency(metrics.get("Начислено Ozon за период (по балансу, календарные дни)")), tooltip=business_metric_help["Начислено Ozon"]),
                            kv_row_compact("Наш расчёт", _format_currency(metrics.get("Сумма начисления по нашему расчёту (по датам отгрузки)"))),
                            kv_row_compact("Расхождение", _format_currency(metrics.get("Расхождение: баланс Ozon минус наш расчёт"))),
                            kv_row_compact("Комиссия за ранний вывод", _format_currency(metrics.get("Комиссия за ранний вывод средств")), tooltip=business_metric_help["Ранний вывод"]),
                            kv_row_compact("Баланс на начало → конец", f"{_format_currency(metrics.get('Входящий баланс на начало периода'))} → {_format_currency(metrics.get('Исходящий баланс на конец периода'))}"),
                        ],
                        tone=semantic_surface("neutral", dark=dark_ui),
                        width=290,
                        col=4,
                    ),
                ],
            )
        ]
        if not promotion_campaigns:
            dashboard_promotion.controls = [
                ft.Text("Продвижение", size=24, weight=ft.FontWeight.W_800, color=TEXT),
                card(
                    "Продвижение",
                    "В листе `Кампании` нет кампаний с расходом за период, поэтому блок анализа не заполнен.",
                    [
                        ft.Text(
                            "Как только в месячном отчёте появятся кампании с ненулевым расходом, здесь автоматически появятся KPI и аналитика по эффективности.",
                            size=12,
                            color=MUTED,
                        )
                    ],
                    tone=semantic_surface("neutral_soft", dark=dark_ui),
                ),
            ]
        else:
            best_roas_campaign = max(promotion_campaigns, key=lambda item: item.get("__roas") or 0.0)
            top_revenue_campaign = max(promotion_campaigns, key=lambda item: item.get("__revenue") or 0.0)
            campaign_section_tone = semantic_surface("neutral_soft", dark=dark_ui)
            campaign_card_tone = semantic_surface("neutral", dark=dark_ui)
            campaign_context_tone = semantic_surface("neutral_soft", dark=dark_ui)
            campaign_secondary_tone = semantic_surface("neutral_alt", dark=dark_ui)
            roas_tone = choose_metric_tone(as_float(promotion_summary.get("roas")), good_at_least=5, warn_at_least=3, dark=dark_ui)
            drr_value = as_float(promotion_summary.get("drr"))
            orders_value = as_float(promotion_summary.get("orders")) or 0.0
            revenue_value = as_float(promotion_summary.get("revenue")) or 0.0
            drr_tone = (
                semantic_surface("danger", dark=dark_ui)
                if orders_value <= 0 or revenue_value <= 0
                else choose_metric_tone(drr_value, good_at_most=20, warn_at_most=35, dark=dark_ui)
            )
            ctr_tone = semantic_surface("neutral_alt", dark=dark_ui)
            cr_tone = choose_metric_tone(as_float(promotion_summary.get("cr")), good_at_least=5, warn_at_least=2, dark=dark_ui)
            cpa_limit = (as_float(promotion_summary.get("avg_order_value")) or 0.0) * 0.25
            cpa_warn_limit = (as_float(promotion_summary.get("avg_order_value")) or 0.0) * 0.45
            cpa_value = as_float(promotion_summary.get("cpa"))
            cpa_tone = (
                choose_metric_tone(cpa_value, good_at_most=cpa_limit, warn_at_most=cpa_warn_limit, dark=dark_ui)
                if cpa_limit > 0 and cpa_warn_limit > 0
                else semantic_surface("info", dark=dark_ui)
            )
            orders_tone = semantic_surface("neutral_soft", dark=dark_ui)
            campaigns_count_tone = semantic_surface("neutral_alt", dark=dark_ui)
            promotion_metric_help = {
                "Рекламный расход": "Расход всего по кампаниям с ненулевым расходом за период. Берется как сумма поля 'Расход за период (руб.)'.",
                "Выручка с рекламы": "Выручка, атрибутированная рекламным кампаниям. Сумма поля 'Заказы (руб.)' по кампаниям с расходом.",
                "ROAS": "Return on Ad Spend. Формула: выручка с рекламы / рекламный расход.",
                "ДРР": "Доля рекламных расходов. Формула: рекламный расход / выручка с рекламы × 100%. Чем ниже, тем эффективнее реклама.",
                "CTR": "Click-Through Rate. Формула: клики / показы × 100%. Показывает, насколько объявление привлекает клик.",
                "CR": "Conversion Rate в заказ. Формула: заказы / клики × 100%. Показывает, как трафик конвертируется в покупку.",
                "CPC": "Cost Per Click. Формула: рекламный расход / клики.",
                "CPA": "Cost Per Action или стоимость заказа. Формула: рекламный расход / заказы.",
                "CPM": "Cost Per Mille. Формула: рекламный расход / показы × 1000.",
                "Заказы": "Количество заказов, атрибутированных рекламным кампаниям. Сумма поля 'Заказы (шт.)'.",
                "Средний чек": "Формула: выручка с рекламы / количество заказов.",
                "Кампаний": "Количество кампаний, у которых за период был фактический расход.",
                "Расход": "Фактический расход кампании за период. Поле 'Расход за период (руб.)'.",
                "Выручка": "Выручка, атрибутированная конкретной кампании. Поле 'Заказы (руб.)'.",
                "Клики": "Количество кликов по кампании за период.",
                "Показы": "Количество показов рекламных объявлений кампании.",
                "CPC кампании": "Средняя цена клика кампании. Поле 'Средняя цена клика (руб.)' или расход / клики.",
                "CPA кампании": "Стоимость одного заказа по кампании. Формула: расход / заказы.",
                "CPM кампании": "Стоимость 1000 показов по кампании. Формула: расход / показы × 1000.",
                "CR кампании": "Конверсия клика в заказ по кампании. Формула: заказы / клики × 100%.",
                "CTR кампании": "Кликабельность кампании. Формула: клики / показы × 100%.",
                "ДРР кампании": "Доля рекламных расходов кампании в выручке. Чем ниже, тем лучше при сопоставимой марже.",
                "Средний чек кампании": "Формула: выручка кампании / заказы кампании.",
                "Бюджет": "Установленный бюджет кампании из отчёта. Используется как контекст, а не как KPI эффективности.",
            }

            campaign_cards: list[ft.Control] = []
            campaign_card_col = 12 if len(promotion_campaigns) == 1 else 6 if len(promotion_campaigns) == 2 else 4
            summary_fields = [
                ("Расход", lambda item: _format_currency(item.get("Расход за период (руб.)"))),
                ("Выручка", lambda item: _format_currency(item.get("Заказы (руб.)"))),
                ("ROAS", lambda item: f"{(item.get('__roas') or 0.0):.2f}x".replace(".", ",")),
                ("ДРР", lambda item: _format_percent(item.get("ДРР (%)"))),
            ]
            detail_fields = [
                ("Заказы", lambda item: _format_number(item.get("Заказы (шт.)"))),
                ("CPA", lambda item: _format_currency(item.get("__cpa"))),
                ("CTR", lambda item: _format_percent(item.get("CTR (%)"))),
                ("CR", lambda item: _format_percent(item.get("__cr"))),
            ]
            context_fields = [
                ("Клики", lambda item: _format_number(item.get("Клики"))),
                ("Показы", lambda item: _format_number(item.get("Показы"))),
            ]
            for campaign in promotion_campaigns:
                campaign_name = str(campaign.get("Название кампании") or f"Кампания {campaign.get('ID кампании') or '—'}")
                campaign_id = str(campaign.get("ID кампании") or "—")
                details_wrap = ft.Container(visible=False)
                details_button = ft.TextButton(
                    "Детали",
                    style=ft.ButtonStyle(
                        color=MUTED,
                        padding=ft.Padding.symmetric(horizontal=10, vertical=6),
                        shape=ft.RoundedRectangleBorder(radius=999),
                    ),
                )
                detail_grid = ft.ResponsiveRow(
                    run_spacing=8,
                    spacing=8,
                    controls=[
                        fact_chip(
                            label,
                            formatter(campaign),
                            tone=campaign_context_tone,
                            col=3,
                            tooltip=promotion_metric_help.get(
                                {
                                    "CPA": "CPA кампании",
                                    "CTR": "CTR кампании",
                                    "CR": "CR кампании",
                                    "Заказы": "Заказы",
                                }.get(label, ""),
                                None,
                            ),
                            variant="secondary",
                        )
                        for label, formatter in detail_fields
                    ]
                    + [
                        fact_chip(
                            label,
                            formatter(campaign),
                            tone=campaign_secondary_tone,
                            col=6,
                            tooltip=promotion_metric_help.get(
                                {
                                    "Показы": "Показы",
                                    "Клики": "Клики",
                                }.get(label, ""),
                                None,
                            ),
                            variant="secondary",
                        )
                        for label, formatter in context_fields
                    ],
                )
                details_wrap.content = ft.Container(
                    padding=ft.Padding.only(top=2),
                    content=detail_grid,
                )

                def toggle_campaign_details(_e=None, target=details_wrap, button=details_button):
                    target.visible = not target.visible
                    button.text = "Скрыть детали" if target.visible else "Детали"
                    if page.controls:
                        page.update(target, button)

                details_button.on_click = toggle_campaign_details
                body = ft.Column(
                    [
                        ft.Row(
                            [
                                ft.Text(campaign_name, expand=True, size=18, weight=ft.FontWeight.W_700, color=TEXT),
                                ft.Text(f"ID {campaign_id}", size=10, color=MUTED, weight=ft.FontWeight.W_600),
                            ],
                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                            vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        ),
                        ft.Container(
                            padding=ft.Padding.symmetric(horizontal=10, vertical=6),
                            border_radius=12,
                            bgcolor=str(campaign.get("__efficiency_tone") or campaign_context_tone),
                            tooltip=str(campaign.get("__efficiency_reason") or ""),
                            content=ft.Row(
                                [
                                    ft.Text(str(campaign.get("__efficiency_label") or "Без оценки"), size=11, weight=ft.FontWeight.W_700, color=TEXT),
                                    ft.Text(
                                        f"ROAS {(campaign.get('__roas') or 0.0):.2f}x".replace(".", ","),
                                        size=11,
                                        color=MUTED,
                                        weight=ft.FontWeight.W_600,
                                    ),
                                ],
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                            ),
                        ),
                        ft.ResponsiveRow(
                            run_spacing=8,
                            spacing=8,
                            controls=[
                                fact_chip(
                                    label,
                                    formatter(campaign),
                                    tone=campaign_card_tone,
                                    col=3,
                                    tooltip=promotion_metric_help.get(
                                        {
                                            "Расход": "Расход",
                                            "Выручка": "Выручка",
                                            "ROAS": "ROAS",
                                            "ДРР": "ДРР кампании",
                                            "CPA": "CPA кампании",
                                            "CTR": "CTR кампании",
                                            "CR": "CR кампании",
                                            "Заказы": "Заказы",
                                        }.get(label, ""),
                                        None,
                                    ),
                                )
                                for label, formatter in summary_fields
                            ],
                        ),
                        ft.Row(
                            [
                                details_button,
                                ft.Text(
                                    "Заказы, CPA, CTR, CR, клики и показы",
                                    size=10,
                                    color=MUTED,
                                ),
                            ],
                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                            vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        ),
                        details_wrap,
                    ],
                    spacing=10,
                )
                campaign_card = ft.Container(
                    col=campaign_card_col,
                    padding=18,
                    bgcolor=campaign_card_tone,
                    border_radius=24,
                    border=ft.Border.all(1, BORDER),
                    shadow=ft.BoxShadow(blur_radius=16, spread_radius=0, color="#0D000000", offset=ft.Offset(0, 6)),
                    content=body,
                )

                def hover_campaign_card(e: ft.ControlEvent, target=campaign_card):
                    hovered = e.data == "true"
                    if hovered:
                        target.border = ft.Border.all(1, "#D8DEE8" if not current_dark["value"] else "#3E474D")
                        target.shadow = ft.BoxShadow(
                            blur_radius=22,
                            spread_radius=0,
                            color="#12000000" if not current_dark["value"] else "#18000000",
                            offset=ft.Offset(0, 8),
                        )
                    else:
                        target.border = ft.Border.all(1, "#D2D2D7" if not current_dark["value"] else "#3A3A3C")
                        target.shadow = ft.BoxShadow(blur_radius=16, spread_radius=0, color="#0D000000", offset=ft.Offset(0, 6))
                    if page.controls:
                        page.update(target)

                campaign_card.on_hover = hover_campaign_card
                campaign_cards.append(campaign_card)

            summary_controls = [
                hero_metric("Рекламный расход", _format_currency(promotion_summary.get("spend")), "суммарно по кампаниям с расходом", semantic_surface("neutral", dark=dark_ui), ACCENT, width=330, col=4, tooltip=promotion_metric_help["Рекламный расход"]),
                hero_metric("Выручка с рекламы", _format_currency(promotion_summary.get("revenue")), "атрибутированная выручка кампаний", semantic_surface("neutral", dark=dark_ui), PRIMARY, width=330, col=4, tooltip=promotion_metric_help["Выручка с рекламы"]),
                hero_metric("ROAS", f"{(promotion_summary.get('roas') or 0.0):.2f}x".replace(".", ","), "возврат на рекламные инвестиции", roas_tone, "#4A6FA5", width=330, col=4, tooltip=promotion_metric_help["ROAS"]),
                metric("ДРР", _format_percent(promotion_summary.get("drr")), "доля рекламы в выручке", drr_tone, width=150, col=2, tooltip=promotion_metric_help["ДРР"]),
                metric("CTR", _format_percent(promotion_summary.get("ctr")), "кликабельность объявлений", semantic_surface("neutral_alt", dark=dark_ui), width=150, col=2, tooltip=promotion_metric_help["CTR"]),
                metric("CR", _format_percent(promotion_summary.get("cr")), "конверсия клика в заказ", cr_tone, width=150, col=2, tooltip=promotion_metric_help["CR"]),
                metric("CPC", _format_currency(promotion_summary.get("cpc")), "цена клика", semantic_surface("neutral", dark=dark_ui), width=150, col=2, tooltip=promotion_metric_help["CPC"]),
                metric("CPA", _format_currency(promotion_summary.get("cpa")), "стоимость заказа", cpa_tone, width=150, col=2, tooltip=promotion_metric_help["CPA"]),
                metric("CPM", _format_currency(promotion_summary.get("cpm")), "цена 1000 показов", semantic_surface("neutral_alt", dark=dark_ui), width=150, col=2, tooltip=promotion_metric_help["CPM"]),
                metric("Заказы", _format_number(promotion_summary.get("orders")), "с рекламы", orders_tone, width=150, col=2, tooltip=promotion_metric_help["Заказы"]),
                metric("Средний чек", _format_currency(promotion_summary.get("avg_order_value")), "по рекламным заказам", semantic_surface("neutral", dark=dark_ui), width=150, col=2, tooltip=promotion_metric_help["Средний чек"]),
                metric("Кампаний", _format_number(promotion_summary.get("campaigns_count")), "с фактическим расходом", campaigns_count_tone, width=150, col=2, tooltip=promotion_metric_help["Кампаний"]),
            ]
            insight_controls = [
                card(
                    "Лучшая по ROAS",
                    "",
                    [
                        ft.Text(str(best_roas_campaign.get("Название кампании") or best_roas_campaign.get("ID кампании") or "—"), size=18, weight=ft.FontWeight.W_700, color=TEXT),
                        kv_row_compact("ROAS", f"{(best_roas_campaign.get('__roas') or 0.0):.2f}x".replace(".", ",")),
                    ],
                    tone=semantic_surface("neutral_alt", dark=dark_ui),
                    col=4,
                ),
                card(
                    "Лидер по выручке",
                    "",
                    [
                        ft.Text(str(top_revenue_campaign.get("Название кампании") or top_revenue_campaign.get("ID кампании") or "—"), size=18, weight=ft.FontWeight.W_700, color=TEXT),
                        kv_row_compact("Выручка", _format_currency(top_revenue_campaign.get("__revenue"))),
                    ],
                    tone=semantic_surface("neutral", dark=dark_ui),
                    col=4,
                ),
                card(
                    "Итог по каналу",
                    "",
                    [
                        ft.ResponsiveRow(
                            run_spacing=8,
                            spacing=8,
                            controls=[
                                fact_chip("Расход всего", _format_currency(promotion_summary.get("spend")), tone=campaign_context_tone, col=6),
                                fact_chip("Кликов", _format_number(promotion_summary.get("clicks")), tone=campaign_secondary_tone, col=6, variant="secondary"),
                                fact_chip("Показов", _format_number(promotion_summary.get("impressions")), tone=campaign_secondary_tone, col=6, variant="secondary"),
                                fact_chip("Кампаний", _format_number(promotion_summary.get("campaigns_count")), tone=campaign_secondary_tone, col=6, variant="secondary"),
                            ],
                        )
                    ],
                    tone=semantic_surface("neutral_soft", dark=dark_ui),
                    col=4,
                ),
            ]
            dashboard_promotion.controls = [
                ft.Text("Продвижение", size=24, weight=ft.FontWeight.W_800, color=TEXT),
                ft.ResponsiveRow(run_spacing=14, spacing=14, controls=summary_controls),
                ft.ResponsiveRow(run_spacing=12, spacing=12, controls=insight_controls),
                card(
                    "Кампании с фактическим расходом",
                    "",
                    [
                        ft.ResponsiveRow(run_spacing=12, spacing=12, controls=campaign_cards),
                    ],
                    tone=campaign_section_tone,
                    col=12,
                ),
            ]
        if current_dark["value"]:
            apply_control_theme(dashboard_summary, True)
            apply_control_theme(dashboard_metric_wrap, True)
            apply_control_theme(dashboard_details, True)
            apply_control_theme(dashboard_promotion, True)
        if page.controls:
            page.update(dashboard_summary, dashboard_metric_wrap, dashboard_details, dashboard_promotion, dashboard_freshness, dashboard_chart_panel)

    def valid_dashboard_chart_period() -> str | None:
        error = valid_month_year(dashboard_chart_from_m.value or "", dashboard_chart_from_y.value or "") or valid_month_year(
            dashboard_chart_to_m.value or "", dashboard_chart_to_y.value or ""
        )
        if error:
            return error
        if not (dashboard_chart_type.value or ""):
            return "Сначала выберите тип графика."
        if (int(dashboard_chart_from_y.value or 0), int(dashboard_chart_from_m.value or 0)) > (
            int(dashboard_chart_to_y.value or 0),
            int(dashboard_chart_to_m.value or 0),
        ):
            return "Начало периода графика должно быть не позже конца."
        return None

    def get_missing_chart_reports() -> list[tuple[int, int]]:
        from_month = int(dashboard_chart_from_m.value or prev.month)
        from_year = int(dashboard_chart_from_y.value or prev.year)
        to_month = int(dashboard_chart_to_m.value or prev.month)
        to_year = int(dashboard_chart_to_y.value or prev.year)
        missing: list[tuple[int, int]] = []
        for month_value, year_value in iter_period_range(from_month, from_year, to_month, to_year):
            if not report_path_for_period(month_value, year_value).exists():
                missing.append((month_value, year_value))
        return missing

    def get_missing_abc_reports() -> list[tuple[int, int]]:
        from_month = int(abc_fm.value or prev.month)
        from_year = int(abc_fy.value or prev.year)
        to_month = int(abc_tm.value or today.month)
        to_year = int(abc_ty.value or today.year)
        missing: list[tuple[int, int]] = []
        for month_value, year_value in iter_period_range(from_month, from_year, to_month, to_year):
            if not report_path_for_period(month_value, year_value).exists():
                missing.append((month_value, year_value))
        return missing

    def build_missing_chart_reports(missing_periods: list[tuple[int, int]], action_type: str) -> None:
        if not missing_periods:
            return
        if not job_lock.acquire(blocking=False):
            toast("Дождитесь завершения текущей операции.", True)
            return
        job_cancel_event.clear()
        chart_pending_action["type"] = action_type
        period_labels = ", ".join(f"{MONTHS[m - 1]} {y}" for m, y in missing_periods)
        if action_type == "abc":
            title = "Подготовка отчётов для ABC/XYZ"
            lead = "Готовим monthly-отчёты для ABC/XYZ"
        else:
            title = "Подготовка данных для графика"
            lead = "Подготавливаем недостающие отчёты для графика"
        cancel_state["requested"] = False
        cancel_state["title"] = title
        set_status("Готовим отчёты", period_labels, ACCENT, busy=True)
        if action_type != "abc":
            set_dashboard_chart_feedback("loading", "Подготавливаем данные", f"Собираем недостающие отчёты: {period_labels}.")
        set_log(f"{lead}: {period_labels}")
        page.update()

        def worker() -> None:
            log_lines = deque([f"{lead}: {period_labels}\n"], maxlen=500)

            def on_line(line: str) -> None:
                log_lines.append(line)
                ui_queue.put(("log_append", line))

            last_code = 0
            for month_value, year_value in missing_periods:
                if job_cancel_event.is_set():
                    ui_queue.put(("finish", title, 130, "Остановлено пользователем."))
                    return
                cmd = [
                    str(PYTHON),
                    str(SCRIPTS / "Monthly_sales_report.py"),
                    "--month",
                    str(month_value),
                    "--year",
                    str(year_value),
                ]
                log_lines.append(f"\n$ {' '.join(cmd)}\n")
                ui_queue.put(("log_append", f"\n$ {' '.join(cmd)}\n"))
                try:
                    last_code, output = stream_process(
                        cmd,
                        on_line,
                        on_prompt=request_prompt,
                        extra_env={"OZONREPORTX_UI_PROMPTS": "1"},
                        proc_holder=active_proc,
                        cancel_event=job_cancel_event,
                    )
                except Exception as exc:
                    last_code, output = 1, f"Не удалось выполнить задачу: {exc}"
                if last_code != 0:
                    ui_queue.put(("finish", title, last_code, output))
                    return
            ui_queue.put(("finish", title, 0, "".join(log_lines).rstrip()))

        try:
            page.run_thread(worker)
        except Exception as exc:
            ui_queue.put(("finish", title, 1, f"Не удалось запустить задачу: {exc}"))

    def build_dashboard_chart(_e=None) -> None:
        chart_error = valid_dashboard_chart_period()
        if chart_error:
            toast(chart_error, True)
            return
        missing_periods = get_missing_chart_reports()
        if missing_periods:
            build_missing_chart_reports(missing_periods, "build")
            return
        from_month = int(dashboard_chart_from_m.value or prev.month)
        from_year = int(dashboard_chart_from_y.value or prev.year)
        to_month = int(dashboard_chart_to_m.value or prev.month)
        to_year = int(dashboard_chart_to_y.value or prev.year)
        set_status("Строим график", f"{MONTHS[from_month - 1]} {from_year} - {MONTHS[to_month - 1]} {to_year}", ACCENT, busy=True)
        set_dashboard_chart_feedback("loading", "Строим график", f"{MONTHS[from_month - 1]} {from_year} - {MONTHS[to_month - 1]} {to_year}")
        if page.controls:
            page.update(dashboard_chart_panel)
        render_dashboard_chart(from_month, from_year, to_month, to_year)
        if dashboard_chart_state["built"]:
            set_status("График построен", dashboard_chart_note.value, PRIMARY, busy=False)
        else:
            set_status("Нет данных", "Для графика нет отчётов.", DANGER, busy=False)
        if page.controls:
            page.update(dashboard_chart_panel)

    def launch_abc_analysis() -> None:
        abc_error = valid_abc(abc_fm.value or "", abc_fy.value or "", abc_tm.value or "", abc_ty.value or "")
        if abc_error:
            toast(abc_error, True)
            return
        missing_periods = get_missing_abc_reports()
        if missing_periods:
            build_missing_chart_reports(missing_periods, "abc")
            return
        job(
            "ABC/XYZ",
            [
                str(PYTHON),
                str(SCRIPTS / "ABC_XYZ_analytics_report.py"),
                "-i",
                "reports",
                "--output_dir",
                "ABC&XYZ reports",
                "--from-month",
                abc_fm.value or "",
                "--from-year",
                abc_fy.value or "",
                "--to-month",
                abc_tm.value or "",
                "--to-year",
                abc_ty.value or "",
            ],
        )

    def dashboard_period_changed(_e=None) -> None:
        selected_month = int(dashboard_m.value or today.month)
        selected_year = int(dashboard_y.value or today.year)
        report_path = report_path_for_period(selected_month, selected_year)
        set_log(f"Открываем отчёт за {MONTHS[selected_month - 1]} {selected_year}: {report_path.name}")
        set_status("Открытие периода", f"{MONTHS[selected_month - 1]} {selected_year}", ACCENT, busy=True)
        refresh_dashboard()
        if not report_path.exists():
            set_log(f"Нет данных за {MONTHS[selected_month - 1]} {selected_year}. Нажмите «Сформировать отчёт» для получения данных.")
        elif page.controls:
            page.update(log)

    def apply_control_theme(control, dark: bool) -> None:
        if control is None:
            return
        if isinstance(control, ft.Text):
            if control.data == "preserve_excel_text":
                return
            current = (control.color or "").upper() if isinstance(control.color, str) else control.color
            if current in PRIMARY_TEXT_COLORS:
                control.color = "#F5F5F7" if dark else "#1D1D1F"
            elif current in MUTED_TEXT_COLORS:
                control.color = "#A1A1A6" if dark else "#6E6E73"
            else:
                control.color = themed_color(control.color, dark)
        if isinstance(control, ft.Dropdown):
            control.color = "#F5F5F7" if dark else "#1D1D1F"
            control.bgcolor = themed_color(control.bgcolor, dark)
            control.border_color = themed_color(control.border_color, dark)
            if control.focused_border_color:
                control.focused_border_color = PRIMARY
            if getattr(control, "label_style", None):
                control.label_style = ft.TextStyle(
                    size=getattr(control.label_style, "size", 11) or 11,
                    color="#A1A1A6" if dark else "#6E6E73",
                )
        if isinstance(control, ft.Container):
            if control.data != "preserve_excel_fill":
                control.bgcolor = themed_color(control.bgcolor, dark)
                if control.border is not None and hasattr(control.border, "top") and control.border.top is not None:
                    border_color = themed_color(control.border.top.color, dark)
                    control.border = ft.Border.all(control.border.top.width, border_color)
        for attr in ("content", "title", "subtitle", "leading", "trailing"):
            if hasattr(control, attr):
                child = getattr(control, attr)
                apply_control_theme(child, dark)
        for attr in ("controls", "actions", "destinations"):
            if hasattr(control, attr):
                children = getattr(control, attr) or []
                for child in children:
                    apply_control_theme(child, dark)

    def refresh_files() -> None:
        files_col.controls = []
        monthly_reports_list.controls = []
        abc_reports_list.controls = []
        for spec in folder_specs:
            rows: list[ft.Control] = [ft.TextButton("Открыть папку", on_click=lambda _e, x=spec.path: reveal_path(x, x.name))]
            batch = files_by_relevance(spec.path)[:12]
            if not batch:
                rows.append(ft.Text("Файлы пока не найдены.", color=MUTED))
            else:
                for p in batch:
                    rows.append(
                        ft.Row(
                            [
                                ft.Text(p.name, expand=True),
                                ft.Text(datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M"), color=MUTED),
                                ft.IconButton(ft.Icons.OPEN_IN_NEW, on_click=lambda _e, x=p: reveal_path(x, x.name)),
                            ],
                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        )
                    )
            files_col.controls.append(card(spec.title, f"Папка: {spec.path.name}", rows))

            target_list = None
            if spec.path == ROOT / "reports":
                target_list = monthly_reports_list
            elif spec.path == ROOT / "ABC&XYZ reports":
                target_list = abc_reports_list
            if target_list is not None:
                if not batch:
                    target_list.controls.append(
                        ft.Container(
                            padding=14,
                            border_radius=18,
                            bgcolor="#F7F7FA",
                            border=ft.Border.all(1, "#E5E5EA"),
                            content=ft.Text("Файлы пока не найдены.", color=MUTED, size=12),
                        )
                    )
                else:
                    for p in batch:
                        target_list.controls.append(
                            report_list_item(
                                p.name,
                                datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M"),
                                lambda _e, x=p: reveal_path(x, x.name),
                            )
                        )

        if current_dark["value"]:
            apply_control_theme(files_col, True)
            apply_control_theme(monthly_reports_list, True)
            apply_control_theme(abc_reports_list, True)

    def set_pricing_costs_dirty(flag: bool, push_update: bool = False) -> None:
        pricing_costs_dirty["value"] = flag
        pricing_costs_save_button.disabled = not flag
        pricing_costs_save_button.text = "Сохранить изменения *" if flag else "Сохранить изменения"
        if push_update and page.controls:
            page.update(pricing_costs_save_button)

    def mark_pricing_costs_dirty(_e=None) -> None:
        pricing_costs_dirty["revision"] += 1
        if not pricing_costs_dirty["value"]:
            set_pricing_costs_dirty(True, push_update=True)

    def build_costs_edit_field(row_number: int, column: str, cell_payload: dict[str, object] | None, dark: bool, expand: int) -> ft.Control:
        payload = cell_payload or {}
        fill_color = payload.get("fill_color") if isinstance(payload.get("fill_color"), str) else None
        font_color = payload.get("font_color") if isinstance(payload.get("font_color"), str) else None
        resolved_text_color = font_color or _excel_text_color(fill_color, dark)
        field = ft.TextField(
            value=_format_costs_editor_value(column, payload.get("value")),
            expand=True,
            dense=True,
            text_size=12,
            color=resolved_text_color,
            bgcolor="transparent" if fill_color else ("#232A2F" if dark else "#FFFFFF"),
            border_color="transparent" if fill_color else ("#D6D0C4" if not dark else "#465056"),
            focused_border_color=PRIMARY,
            border_radius=10,
            content_padding=ft.Padding.symmetric(horizontal=10, vertical=8),
            on_change=mark_pricing_costs_dirty,
        )
        pricing_costs_editors.setdefault(row_number, {})[column] = field
        return ft.Container(
            expand=expand,
            bgcolor=fill_color,
            border_radius=10,
            padding=2 if not fill_color else 0,
            data="preserve_excel_fill" if fill_color else None,
            content=field,
        )

    def refresh_pricing_costs_preview(push_update: bool = False) -> str | None:
        if pricing_costs_dirty["value"]:
            return None
        saved_min_margin, saved_desired_margin = margin_defaults()
        active_min_margin = margin_min.value or saved_min_margin
        active_desired_margin = margin_des.value or saved_desired_margin
        if valid_margin(active_min_margin, active_desired_margin):
            active_min_margin, active_desired_margin = saved_min_margin, saved_desired_margin
        min_margin_pct = float(active_min_margin.replace(",", ".")) * 100
        desired_margin_pct = float(active_desired_margin.replace(",", ".")) * 100

        try:
            signature = file_signature(ROOT / "costs.xlsx")
        except OSError:
            signature = None
        rows, error = read_costs_main_rows(
            ROOT / "costs.xlsx",
            min_margin_pct=min_margin_pct,
            desired_margin_pct=desired_margin_pct,
        )
        if signature is not None:
            try:
                if file_signature(ROOT / "costs.xlsx") != signature:
                    return "Файл изменился во время чтения. Обновите таблицу ещё раз."
            except OSError as exc:
                return f"Не удалось проверить costs.xlsx: {exc}"
        # An edit event may have arrived while Excel was being read.
        if pricing_costs_dirty["value"]:
            return None
        pricing_costs_snapshot["signature"] = signature
        dark = current_dark["value"]
        header_border = "#D8DEE8" if not dark else "#3E474E"
        row_border = "#E5E7EF" if not dark else "#384046"
        pricing_costs_editors.clear()
        set_pricing_costs_dirty(False)

        if error:
            pricing_costs_status.value = error
            pricing_costs_status.color = "#E6BBB5" if dark else DANGER
            pricing_costs_status_shell.bgcolor = semantic_surface("danger", dark=dark)
            pricing_costs_status_shell.border = ft.Border.all(1, "#D9A49D" if not dark else "#6E464C")
            pricing_costs_list.controls = [
                ft.Container(
                    padding=18,
                    border_radius=18,
                    bgcolor=semantic_surface("danger", dark=dark),
                    border=ft.Border.all(1, "#D9A49D" if not dark else "#6E464C"),
                    content=ft.Text(error, size=12, color="#F3F4F6" if dark else DANGER),
                )
            ]
        else:
            refreshed_at = datetime.fromtimestamp((ROOT / "costs.xlsx").stat().st_mtime).strftime("%Y-%m-%d %H:%M")
            pricing_costs_status.value = f"Строк: {len(rows)} • лист «Основной» • обновлено {refreshed_at}"
            pricing_costs_status.color = "#B6C0BC" if dark else MUTED
            pricing_costs_status_shell.bgcolor = semantic_surface("neutral_alt", dark=dark)
            pricing_costs_status_shell.border = ft.Border.all(1, header_border)

            lead_expand = COSTS_MAIN_COLUMN_EXPANDS["Артикул"] + COSTS_MAIN_COLUMN_EXPANDS["Себестоимость"]
            calculated_expand = sum(COSTS_MAIN_COLUMN_EXPANDS[column] for column in PRICING_CALCULATED_COLUMNS)
            sync_expand = sum(COSTS_MAIN_COLUMN_EXPANDS[column] for column in PRICING_SYNC_COLUMNS)
            header = ft.Container(
                padding=ft.Padding.symmetric(horizontal=12, vertical=10),
                bgcolor=semantic_surface("neutral_alt", dark=dark),
                border_radius=14,
                border=ft.Border.all(1, header_border),
                content=ft.Column(
                    [
                        ft.Row(
                            [
                                ft.Container(expand=lead_expand),
                                ft.Container(
                                    expand=calculated_expand,
                                    padding=ft.Padding.symmetric(horizontal=10, vertical=8),
                                    border_radius=12,
                                    bgcolor=semantic_surface("warning", dark=dark),
                                    border=ft.Border.all(1, "#CFBE8A" if not dark else "#655C43"),
                                    content=ft.Column(
                                        [
                                            ft.Button(
                                                "Рассчитать цены",
                                                bgcolor=ACCENT,
                                                color="white",
                                                height=34,
                                                on_click=lambda _e: launch_recommended_prices(),
                                            ),
                                            ft.Text(
                                                "Обновит минимальную и желательную цену продажи",
                                                size=10,
                                                color="#7A5A20" if not dark else "#F6E3B0",
                                                text_align=ft.TextAlign.CENTER,
                                            ),
                                        ],
                                        spacing=4,
                                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                                        tight=True,
                                    ),
                                ),
                                ft.Container(
                                    expand=sync_expand,
                                    padding=ft.Padding.symmetric(horizontal=10, vertical=8),
                                    border_radius=12,
                                    bgcolor=semantic_surface("info", dark=dark),
                                    border=ft.Border.all(1, "#C7D7EF" if not dark else "#425264"),
                                    content=ft.Column(
                                        [
                                            ft.Button(
                                                "Синхронизировать",
                                                bgcolor=PRIMARY,
                                                color="white",
                                                height=34,
                                                on_click=lambda _e: job("Текущие цены и акции", [str(PYTHON), str(SCRIPTS / "ui_actions.py"), "refresh-pricing"]),
                                            ),
                                            ft.Text(
                                                "Обновит цену Ozon, цену со скидкой и рентабельность",
                                                size=10,
                                                color="#3C5D85" if not dark else "#B9D4FF",
                                                text_align=ft.TextAlign.CENTER,
                                            ),
                                        ],
                                        spacing=4,
                                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                                        tight=True,
                                    ),
                                ),
                            ],
                            spacing=12,
                            vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        ),
                        ft.Row(
                            [ft.Text(column, expand=COSTS_MAIN_COLUMN_EXPANDS[column], size=11, weight=ft.FontWeight.W_800, color="#A1A1A6" if dark else MUTED) for column in COSTS_MAIN_COLUMNS],
                            spacing=12,
                            vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        ),
                    ],
                    spacing=10,
                ),
            )

            row_controls: list[ft.Control] = [header]
            for idx, row in enumerate(rows):
                row_number = int(row.get("__row_number") or 0)
                row_controls.append(
                    ft.Container(
                        padding=ft.Padding.symmetric(horizontal=12, vertical=10),
                        bgcolor=semantic_surface("neutral", dark=dark) if idx % 2 == 0 else semantic_surface("neutral_soft", dark=dark),
                        border_radius=14,
                        border=ft.Border.all(1, row_border),
                        content=ft.Row(
                            [
                                build_costs_edit_field(row_number, column, row.get(column), dark, COSTS_MAIN_COLUMN_EXPANDS[column])
                                if column in EDITABLE_COSTS_COLUMNS and row_number >= 2
                                else _build_costs_cell(column, row.get(column), dark, COSTS_MAIN_COLUMN_EXPANDS[column])
                                for column in COSTS_MAIN_COLUMNS
                            ],
                            spacing=12,
                            vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        ),
                    )
                )
            pricing_costs_list.controls = row_controls

        if push_update and page.controls:
            page.update(pricing_costs_status, pricing_costs_status_shell, pricing_costs_list, pricing_costs_save_button)
        return error

    def _save_pricing_costs(_e=None) -> None:
        revision = pricing_costs_dirty["revision"]
        updates: list[dict[str, object]] = []
        for row_number, editors in pricing_costs_editors.items():
            article_value = (editors.get("Артикул").value or "").strip() if editors.get("Артикул") else ""
            cost_raw = (editors.get("Себестоимость").value or "").strip() if editors.get("Себестоимость") else ""
            if cost_raw:
                try:
                    cost_value = float(cost_raw.replace(" ", "").replace(",", "."))
                except ValueError:
                    set_status("Ошибка сохранения", f"Строка {row_number}: себестоимость должна быть числом.", DANGER, busy=False)
                    return
            else:
                cost_value = None
            updates.append(
                {
                    "row_number": row_number,
                    "values": {
                        "Артикул": article_value or None,
                        "Себестоимость": cost_value,
                    },
                }
            )

        saved_signature = {}
        error = write_costs_main_rows(ROOT / "costs.xlsx", updates, expected_signature=pricing_costs_snapshot["signature"], saved_signature=saved_signature)
        if error:
            set_status("Ошибка сохранения", error, DANGER, busy=False)
            return
        pricing_costs_snapshot["signature"] = saved_signature["signature"]
        if pricing_costs_dirty["revision"] != revision:
            set_status("Изменения сохранены", "В редакторе есть более новые несохранённые правки.", ACCENT, busy=False)
            return
        set_pricing_costs_dirty(False)
        refresh_error = refresh_pricing_costs_preview(push_update=True)
        if refresh_error:
            set_status("Изменения сохранены", "Значения записаны, но превью не удалось перечитать.", ACCENT, busy=False)
        else:
            set_status("Изменения сохранены", "Артикул и себестоимость обновлены в costs.xlsx.", PRIMARY, busy=False)

    def handle_pricing_costs_save(_e=None) -> None:
        if not job_lock.acquire(blocking=False):
            toast("Дождитесь завершения текущей операции перед сохранением.", True)
            return
        try:
            _save_pricing_costs(_e)
        finally:
            job_lock.release()

    def handle_pricing_costs_refresh(_e=None) -> None:
        if pricing_costs_dirty["value"]:
            discard_costs_dialog.open = True
            page.update(discard_costs_dialog)
            return
        error = refresh_pricing_costs_preview(push_update=True)
        if error:
            set_status("Ошибка таблицы", error, DANGER, busy=False)
        else:
            set_status("Таблица обновлена", "Лист «Основной» перечитан из costs.xlsx.", PRIMARY, busy=False)

    def close_discard_costs_dialog(_e=None) -> None:
        discard_costs_dialog.open = False
        page.update(discard_costs_dialog)

    def discard_costs_edits(_e=None) -> None:
        close_discard_costs_dialog()
        set_pricing_costs_dirty(False)
        handle_pricing_costs_refresh()

    discard_costs_dialog = ft.AlertDialog(
        modal=True,
        title=ft.Text("Есть несохранённые изменения"),
        content=ft.Text("Обновление таблицы удалит введённые правки. Оставьте их в редакторе или явно отмените."),
        actions=[ft.TextButton("Оставить правки", on_click=close_discard_costs_dialog),
                 ft.TextButton("Отменить правки и обновить", on_click=discard_costs_edits)],
    )
    page.overlay.append(discard_costs_dialog)
    pricing_costs_save_button.on_click = handle_pricing_costs_save

    def refresh_report_lists(_e=None) -> None:
        refresh_files()
        set_status("Списки обновлены", "Отчёты перечитаны.", PRIMARY, busy=False)
        if page.controls:
            page.update(files_col, monthly_reports_list, abc_reports_list)

    def finish(title: str, code: int, output: str) -> None:
        was_cancelled = cancel_state["requested"]
        active_proc["proc"] = None
        cancel_state["requested"] = False
        cancel_state["title"] = ""
        set_log(output, persist=False)
        append_log(f"\nЗадача «{title}» завершена, код {code}.")
        if code != 0:
            append_log(output)
        try:
            refresh_dashboard()
            refresh_pricing_costs_preview()
        except Exception as exc:
            append_log(f"\nНе удалось обновить интерфейс: {exc}")
        finally:
            if job_lock.locked():
                job_lock.release()
        if was_cancelled:
            set_status(f"Отменено: {title}", "Остановлено пользователем.", DANGER, busy=False)
        else:
            set_status(f"Готово: {title}" if code == 0 else f"Ошибка ({code}): {title}", "Задача завершена.", PRIMARY if code == 0 else DANGER, busy=False)
        pending_action = chart_pending_action["type"]
        chart_pending_action["type"] = None
        if was_cancelled:
            pending_action = None
        if code == 0 and pending_action == "build":
            build_dashboard_chart()
        elif code == 0 and pending_action == "interactive":
            open_interactive_dashboard_chart()
        elif code == 0 and pending_action == "abc":
            launch_abc_analysis()
        elif code == 0 and dashboard_chart_state["built"] and not was_cancelled:
            build_dashboard_chart()
        refresh_files()
        if page.controls:
            page.update(state, spin, cancel_button, files_col, monthly_reports_list, abc_reports_list, pricing_costs_list, pricing_costs_status, pricing_costs_status_shell, pricing_costs_save_button, log)

    def job(title: str, cmd: list[str], prompt_for_costs: bool = False) -> None:
        if not job_lock.acquire(blocking=False):
            toast("Дождитесь завершения текущей операции.", True)
            return
        job_cancel_event.clear()
        cancel_state["requested"] = False
        cancel_state["title"] = title
        set_status(f"Выполняется: {title}", "Фоновый процесс запущен.", ACCENT, busy=True)
        set_log(f"$ {' '.join(cmd)}")
        page.update()

        def worker() -> None:
            def on_line(line: str) -> None:
                ui_queue.put(("log_append", line))

            try:
                code, output = stream_process(
                    cmd,
                    on_line,
                    on_prompt=request_prompt if prompt_for_costs else None,
                    extra_env={"OZONREPORTX_UI_PROMPTS": "1"} if prompt_for_costs else None,
                    proc_holder=active_proc,
                    cancel_event=job_cancel_event,
                )
            except Exception as exc:
                code, output = 1, f"Не удалось выполнить задачу: {exc}"
            ui_queue.put(("finish", title, code, output))

        try:
            page.run_thread(worker)
        except Exception as exc:
            ui_queue.put(("finish", title, 1, f"Не удалось запустить задачу: {exc}"))

    def launch_recommended_prices() -> None:
        margin_error = valid_margin(margin_min.value or "", margin_des.value or "")
        if margin_error:
            toast(margin_error, True)
            return
        job(
            "Рекомендованные цены",
            [
                str(PYTHON),
                str(SCRIPTS / "ui_actions.py"),
                "recommended-prices",
                "--min-margin",
                (margin_min.value or "").replace(",", "."),
                "--desired-margin",
                (margin_des.value or "").replace(",", "."),
            ],
        )

    def launch_monthly_report() -> None:
        month_error = valid_month_year(monthly_m.value or "", monthly_y.value or "")
        if month_error:
            toast(month_error, True)
            return
        job(
            "Месячный отчёт",
            [
                str(PYTHON),
                str(SCRIPTS / "Monthly_sales_report.py"),
                "--month",
                monthly_m.value or "",
                "--year",
                monthly_y.value or "",
            ],
            prompt_for_costs=True,
        )

    def open_selected_monthly_report() -> None:
        month_error = valid_month_year(monthly_m.value or "", monthly_y.value or "")
        if month_error:
            toast(month_error, True)
            return
        report_path = report_path_for_period(int(monthly_m.value or today.month), int(monthly_y.value or today.year))
        if not report_path.exists():
            toast("Отчёт за выбранный период пока не найден.", True)
            return
        reveal_path(report_path, report_path.name)

    def open_selected_dashboard_report() -> None:
        month_error = valid_month_year(dashboard_m.value or "", dashboard_y.value or "")
        if month_error:
            toast(month_error, True)
            return
        report_path = report_path_for_period(int(dashboard_m.value or today.month), int(dashboard_y.value or today.year))
        if not report_path.exists():
            toast("Отчёт за выбранный период пока не найден.", True)
            return
        reveal_path(report_path, report_path.name)

    def launch_dashboard_report() -> None:
        month_error = valid_month_year(dashboard_m.value or "", dashboard_y.value or "")
        if month_error:
            toast(month_error, True)
            return
        monthly_m.value = dashboard_m.value
        monthly_y.value = dashboard_y.value
        job(
            "Обновление бизнес-сводки",
            [
                str(PYTHON),
                str(SCRIPTS / "Monthly_sales_report.py"),
                "--month",
                dashboard_m.value or "",
                "--year",
                dashboard_y.value or "",
            ],
            prompt_for_costs=True,
        )

    dashboard_m.on_select = dashboard_period_changed
    dashboard_y.on_select = dashboard_period_changed

    prompt_dialog = ft.AlertDialog(
        modal=True,
        title=ft.Text("Нужны расходы на маркетинг"),
        content=ft.Container(
            width=520,
            padding=ft.Padding.symmetric(horizontal=4, vertical=2),
            content=ft.Column([prompt_label, prompt_hint, prompt_context, prompt_input], tight=True, spacing=10),
        ),
        actions=[
            ft.TextButton("Использовать 0", on_click=lambda _e: None),
            ft.Button("Подтвердить", on_click=lambda _e: None),
        ],
        actions_alignment=ft.MainAxisAlignment.END,
        open=False,
    )

    def prompt_submit(_e=None) -> None:
        holder = prompt_state["holder"]
        event = prompt_state["event"]
        if isinstance(holder, dict):
            holder["value"] = (prompt_input.value or "0").replace(",", ".")
        prompt_dialog.open = False
        page.update(prompt_dialog)
        if isinstance(event, threading.Event):
            event.set()

    prompt_dialog.actions = [
        ft.TextButton("Использовать 0", on_click=prompt_submit),
        ft.Button("Подтвердить", on_click=prompt_submit),
    ]
    page.overlay.append(prompt_dialog)

    discount_dialog_title = ft.Text("Заявки на скидку", color=TEXT, size=18, weight=ft.FontWeight.W_800)
    discount_dialog_badge = ft.Text("", color=MUTED, size=12, weight=ft.FontWeight.W_600)
    discount_dialog_status = ft.Text("", color=MUTED, size=12)
    discount_dialog_badge_shell = ft.Container(
        padding=ft.Padding.symmetric(horizontal=12, vertical=8),
        border_radius=999,
        bgcolor=semantic_surface("neutral_alt", dark=False),
        border=ft.Border.all(1, "#E5E5EA"),
        content=discount_dialog_badge,
    )
    discount_overview_row = ft.ResponsiveRow(run_spacing=8, spacing=8)
    discount_dialog_list = ft.Column(
        controls=[ft.Text("Нажмите «Обновить заявки», чтобы загрузить список.", size=12, color=MUTED)],
        spacing=8,
        scroll=ft.ScrollMode.AUTO,
    )
    discount_dialog_state: dict[str, object] = {"plan": None, "busy": False, "processing_ids": set()}
    discount_dialog: ft.AlertDialog | None = None
    discount_requests_refresh_button = ft.IconButton(
        icon=ft.Icons.SYNC,
        tooltip="Обновить заявки",
        icon_size=18,
        style=ft.ButtonStyle(shape=ft.CircleBorder(), padding=ft.Padding.all(10)),
    )
    discount_requests_auto_button = ft.Button("Обработать автоматически", bgcolor=ACCENT, color="white", disabled=True)

    def pricing_state_panel(
        title: str,
        detail: str,
        *,
        icon: str = ft.Icons.INFO_OUTLINE,
        kind: str = "neutral_soft",
        loading: bool = False,
    ) -> ft.Control:
        dark = current_dark["value"]
        border_map = {
            "neutral": "#D8DEE8" if not dark else "#3E474E",
            "neutral_soft": "#E5E7EF" if not dark else "#384046",
            "neutral_alt": "#E0E5EF" if not dark else "#3D454B",
            "positive": "#9DC9BA" if not dark else "#2E5A50",
            "warning": "#CFBE8A" if not dark else "#655C43",
            "danger": "#D9A49D" if not dark else "#6E464C",
            "info": "#C7D7EF" if not dark else "#425264",
        }
        icon_color_map = {
            "positive": PRIMARY if not dark else "#70D4BC",
            "warning": ACCENT if not dark else "#E0A15F",
            "danger": DANGER if not dark else "#FFD2CB",
            "info": "#5B84C4" if not dark else "#9FC1F7",
        }
        leading: ft.Control = (
            ft.ProgressRing(width=22, height=22, stroke_width=2.4, color=PRIMARY)
            if loading
            else ft.Container(
                width=40,
                height=40,
                border_radius=999,
                bgcolor=semantic_surface("neutral_alt", dark=dark),
                alignment=ft.Alignment(0, 0),
                content=ft.Icon(icon, size=20, color=icon_color_map.get(kind, "#6E6E73" if not dark else "#A1A1A6")),
            )
        )
        return ft.Container(
            expand=True,
            alignment=ft.Alignment(0, 0),
            padding=ft.Padding.symmetric(horizontal=20, vertical=26),
            border_radius=18,
            bgcolor=semantic_surface(kind, dark=dark),
            border=ft.Border.all(1, border_map.get(kind, border_map["neutral_soft"])),
            content=ft.Column(
                [
                    leading,
                    ft.Text(title, size=15, weight=ft.FontWeight.W_700, color="#F3F4F6" if dark else TEXT, text_align=ft.TextAlign.CENTER),
                    ft.Text(detail, size=11, color="#B6C0BC" if dark else MUTED, text_align=ft.TextAlign.CENTER),
                ],
                spacing=10,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                tight=True,
            ),
        )

    def close_discount_dialog(_e=None) -> None:
        discount_dialog_state["plan"] = None
        discount_dialog_state["busy"] = False
        discount_dialog_state["processing_ids"] = set()
        render_discount_dialog()
        if page.controls:
            page.update(discount_dialog_badge, discount_dialog_badge_shell, discount_dialog_status, discount_overview_row, discount_dialog_list, discount_requests_auto_button)

    def current_discount_items() -> list[dict[str, object]]:
        plan = discount_dialog_state.get("plan")
        if not isinstance(plan, dict):
            return []
        items = plan.get("items")
        if not isinstance(items, list):
            return []
        return [item for item in items if isinstance(item, dict)]

    def discount_row_tokens(item: dict[str, object]) -> dict[str, str]:
        dark = current_dark["value"]
        price_state = str(item.get("price_state") or "")
        price_delta = as_float(item.get("price_delta"))

        if price_state == "below_min":
            delta_label = "Ниже минимума"
            if price_delta is not None:
                delta_label = f"Ниже на {_format_currency(abs(price_delta))}"
            return {
                "row_bg": semantic_surface("danger", dark=dark),
                "row_border": "#D9A49D" if not dark else "#6E464C",
                "title_color": "#F3F4F6" if dark else TEXT,
                "value_color": "#FFD2CB" if dark else DANGER,
                "detail_color": "#E6BBB5" if dark else "#7D443E",
                "badge_bg": "#F6DFDB" if not dark else "#5A383D",
                "badge_border": "#D9A49D" if not dark else "#81555C",
                "badge_text": "#8F342B" if not dark else "#FFD2CB",
                "badge_label": "Ниже минимума",
                "delta_label": delta_label,
                "approve_color": ACCENT if not dark else "#E0A15F",
                "approve_tooltip": "Одобрить вручную, несмотря на цену ниже минимума",
            }

        if price_state == "ok" or item.get("recommended_action") == "approve":
            if price_delta is None:
                delta_label = "Можно одобрять"
            elif abs(price_delta) < 0.01:
                delta_label = "Ровно по минимуму"
            else:
                delta_label = f"Запас {_format_currency(price_delta)}"
            return {
                "row_bg": semantic_surface("positive", dark=dark),
                "row_border": "#9DC9BA" if not dark else "#2E5A50",
                "title_color": "#F3F4F6" if dark else TEXT,
                "value_color": "#D8FFF3" if dark else PRIMARY,
                "detail_color": "#C5E0D7" if dark else "#355D54",
                "badge_bg": "#DCEFE5" if not dark else "#295046",
                "badge_border": "#9DC9BA" if not dark else "#3A6E62",
                "badge_text": "#1F5E51" if not dark else "#D8FFF3",
                "badge_label": "Цена нормальная",
                "delta_label": delta_label,
                "approve_color": PRIMARY if not dark else "#70D4BC",
                "approve_tooltip": "Одобрить заявку",
            }

        return {
            "row_bg": semantic_surface("warning", dark=dark),
            "row_border": "#CFBE8A" if not dark else "#655C43",
            "title_color": "#F3F4F6" if dark else TEXT,
            "value_color": "#F6E3B0" if dark else ACCENT,
            "detail_color": "#D8C698" if dark else "#72552A",
            "badge_bg": "#F4EDDB" if not dark else "#5A5037",
            "badge_border": "#CFBE8A" if not dark else "#7B6F4E",
            "badge_text": "#7A5A20" if not dark else "#F6E3B0",
            "badge_label": "Проверить вручную",
            "delta_label": "Нет полной привязки к минимальной цене",
            "approve_color": ACCENT if not dark else "#E0A15F",
            "approve_tooltip": "Одобрить заявку вручную",
        }

    def build_discount_auto_plan() -> dict[str, object]:
        items = current_discount_items()
        approve_tasks = [item.get("approve_payload") for item in items if item.get("recommended_action") == "approve" and item.get("approve_payload")]
        decline_tasks = [item.get("decline_payload") for item in items if item.get("recommended_action") != "approve" and item.get("decline_payload")]
        return {
            "items": items,
            "approve_tasks": approve_tasks,
            "decline_tasks": decline_tasks,
        }

    def remove_discount_item(task_id: object) -> None:
        plan = discount_dialog_state.get("plan")
        if not isinstance(plan, dict):
            return
        items = current_discount_items()
        plan["items"] = [item for item in items if str(item.get("id")) != str(task_id)]

    def restore_discount_item(item: dict[str, object]) -> None:
        plan = discount_dialog_state.get("plan")
        if not isinstance(plan, dict):
            return
        items = current_discount_items()
        items.append(item)
        items.sort(key=lambda row: str(row.get("id") or ""))
        plan["items"] = items

    def render_discount_dialog() -> None:
        items = current_discount_items()
        processing_ids = discount_dialog_state.get("processing_ids")
        processing_ids = processing_ids if isinstance(processing_ids, set) else set()
        dark = current_dark["value"]
        busy = discount_dialog_state.get("busy") is True
        requires_refresh = discount_dialog_state.get("requires_refresh") is True
        neutral_border = "#D8DEE8" if not dark else "#3E474E"
        discount_overview_row.controls = []
        if not isinstance(discount_dialog_state.get("plan"), dict):
            discount_dialog_badge.value = "Собираю список" if busy else "Заявки не загружены"
            discount_dialog_status.value = (
                "Проверяю новые скидочные заявки и готовлю рекомендации."
                if busy
                else "Загрузите список, чтобы увидеть рекомендации и обработать заявки прямо на странице."
            )
            discount_dialog_badge_shell.bgcolor = semantic_surface("info" if busy else "neutral_alt", dark=dark)
            discount_dialog_badge_shell.border = ft.Border.all(
                1,
                ("#C7D7EF" if not dark else "#425264") if busy else neutral_border,
            )
            discount_dialog_list.controls = [
                pricing_state_panel(
                    "Собираю рекомендации" if busy else "Список пока не загружен",
                    "Запрашиваю заявки у Ozon и сопоставляю их с минимальной ценой."
                    if busy
                    else "Нажмите «Обновить заявки», и система покажет решения для каждой заявки.",
                    icon=ft.Icons.SYNC if busy else ft.Icons.LOCAL_OFFER_OUTLINED,
                    kind="info" if busy else "neutral_soft",
                    loading=busy,
                )
            ]
            discount_requests_auto_button.disabled = True
            return
        total = len(items)
        discount_dialog_badge.value = f"Новых заявок: {total}"
        auto_ready_items = [
            item
            for item in items
            if item.get("recommended_action") == "approve" and item.get("approve_payload") is not None
        ]
        attention_items = [item for item in items if item not in auto_ready_items]
        discount_overview_row.controls = [
            fact_chip("Новых", str(total), tone=semantic_surface("neutral_alt", dark=dark), col=4, variant="secondary"),
            fact_chip(
                "Можно авто",
                str(len(auto_ready_items)),
                tone=semantic_surface("positive", dark=dark) if auto_ready_items else semantic_surface("neutral_alt", dark=dark),
                col=4,
                variant="secondary",
            ),
            fact_chip(
                "Требуют внимания",
                str(len(attention_items)),
                tone=semantic_surface("warning", dark=dark) if attention_items else semantic_surface("neutral_alt", dark=dark),
                col=4,
                variant="secondary",
            ),
        ]
        if not items:
            discount_dialog_status.value = "Все заявки обработаны или новых заявок нет."
            discount_dialog_badge_shell.bgcolor = semantic_surface("neutral_alt", dark=dark)
            discount_dialog_badge_shell.border = ft.Border.all(1, neutral_border)
            discount_dialog_list.controls = [
                pricing_state_panel(
                    "Новых заявок нет",
                    "Когда появятся новые запросы на скидку, они появятся здесь с рекомендацией по решению.",
                    icon=ft.Icons.TASK_ALT,
                    kind="neutral_soft",
                ),
            ]
            discount_requests_auto_button.disabled = True
            return

        discount_dialog_badge_shell.bgcolor = (
            semantic_surface("warning", dark=dark)
            if attention_items
            else semantic_surface("positive", dark=dark)
        )
        discount_dialog_badge_shell.border = ft.Border.all(
            1,
            "#CFBE8A" if (attention_items and not dark) else "#655C43" if attention_items else "#9DC9BA" if not dark else "#2E5A50",
        )
        discount_dialog_status.value = (
            "Идёт обработка заявок. По завершении список обновится."
            if busy
            else "Зелёные строки проходят автоматически. Рискованные решения остаются доступны для ручного подтверждения."
        )
        header = ft.Container(
            padding=10,
            bgcolor=semantic_surface("neutral_alt", dark=dark),
            border_radius=12,
            border=ft.Border.all(1, "#D8DEE8" if not dark else "#3E474E"),
            content=ft.Row(
                [
                    ft.Text("Артикул", expand=3, size=11, weight=ft.FontWeight.W_700, color="#A1A1A6" if dark else MUTED),
                    ft.Text("Мин. цена продажи", expand=2, size=11, weight=ft.FontWeight.W_700, color="#A1A1A6" if dark else MUTED),
                    ft.Text("Цена по скидке", expand=2, size=11, weight=ft.FontWeight.W_700, color="#A1A1A6" if dark else MUTED),
                    ft.Text("Действие", width=96, size=11, weight=ft.FontWeight.W_700, color="#A1A1A6" if dark else MUTED, text_align=ft.TextAlign.CENTER),
                ],
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
            ),
        )
        rows: list[ft.Control] = [header]
        for item in items:
            task_id = item.get("id")
            item_id = str(task_id or "")
            reason = str(item.get("recommendation_reason") or "")
            is_processing = item_id in processing_ids or busy or requires_refresh
            row_tokens = discount_row_tokens(item)
            approve_disabled = is_processing or item.get("approve_payload") is None
            approve_tooltip = row_tokens["approve_tooltip"] if not approve_disabled else (reason or "Одобрение недоступно для этой заявки")
            rows.append(
                ft.Container(
                    padding=10,
                    bgcolor=row_tokens["row_bg"],
                    border_radius=12,
                    border=ft.Border.all(1, row_tokens["row_border"]),
                    tooltip=reason,
                    content=ft.Row(
                        [
                            ft.Container(
                                expand=3,
                                content=ft.Column(
                                    [
                                        ft.Text(
                                            str(item.get("offer_id") or item.get("sku") or "—"),
                                            size=12,
                                            color=row_tokens["title_color"],
                                            weight=ft.FontWeight.W_700,
                                        ),
                                        ft.Container(
                                            padding=ft.Padding.symmetric(horizontal=8, vertical=3),
                                            bgcolor=row_tokens["badge_bg"],
                                            border=ft.Border.all(1, row_tokens["badge_border"]),
                                            border_radius=999,
                                            content=ft.Text(
                                                row_tokens["badge_label"],
                                                size=10,
                                                color=row_tokens["badge_text"],
                                                weight=ft.FontWeight.W_700,
                                            ),
                                        ),
                                    ],
                                    spacing=4,
                                    tight=True,
                                ),
                            ),
                            ft.Container(
                                expand=2,
                                content=ft.Column(
                                    [
                                        ft.Text(_format_currency(item.get("min_price")), size=12, color=row_tokens["title_color"], weight=ft.FontWeight.W_600),
                                        ft.Text("Порог Ozon", size=10, color=row_tokens["detail_color"]),
                                    ],
                                    spacing=3,
                                    tight=True,
                                ),
                            ),
                            ft.Container(
                                expand=2,
                                content=ft.Column(
                                    [
                                        ft.Text(
                                            _format_currency(item.get("requested_price")),
                                            size=12,
                                            color=row_tokens["value_color"],
                                            weight=ft.FontWeight.W_700,
                                        ),
                                        ft.Text(row_tokens["delta_label"], size=10, color=row_tokens["detail_color"]),
                                    ],
                                    spacing=3,
                                    tight=True,
                                ),
                            ),
                            ft.Row(
                                [
                                    ft.IconButton(
                                        icon=ft.Icons.CHECK_CIRCLE_OUTLINE,
                                        icon_color=row_tokens["approve_color"],
                                        tooltip=approve_tooltip,
                                        disabled=approve_disabled,
                                        on_click=lambda _e, x=item: submit_discount_decision(x, "approve"),
                                    ),
                                    ft.IconButton(
                                        icon=ft.Icons.CANCEL_OUTLINED,
                                        icon_color=DANGER,
                                        tooltip="Отклонить заявку",
                                        disabled=is_processing,
                                        on_click=lambda _e, x=item: submit_discount_decision(x, "decline"),
                                    ),
                                ],
                                width=96,
                                spacing=0,
                                alignment=ft.MainAxisAlignment.CENTER,
                            ),
                        ],
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    ),
                )
        )
        discount_dialog_list.controls = rows
        discount_requests_auto_button.disabled = busy or requires_refresh

    def submit_discount_decision(item: dict[str, object], action: str) -> None:
        if discount_dialog_state.get("requires_refresh") or discount_dialog_state["busy"]:
            return
        if process_discount_request_item is None:
            toast(f"Модуль обработки скидок недоступен: {PRICING_IMPORT_ERROR or 'ошибка импорта'}", True)
            return
        processing_ids = discount_dialog_state.get("processing_ids")
        if not isinstance(processing_ids, set):
            processing_ids = set()
            discount_dialog_state["processing_ids"] = processing_ids
        item_id = str(item.get("id") or "")
        processing_ids.add(item_id)
        remove_discount_item(item.get("id"))
        render_discount_dialog()
        if page.controls:
            page.update(discount_dialog_badge, discount_dialog_badge_shell, discount_dialog_status, discount_overview_row, discount_dialog_list, discount_requests_auto_button)

        def worker() -> None:
            result = process_discount_request_item(item, action)
            processing_ids.discard(item_id)
            if result.get("ok"):
                toast(str(result.get("message") or "Готово."))
                if not current_discount_items():
                    toast("Все заявки обработаны.")
                    render_discount_dialog()
                else:
                    render_discount_dialog()
                    if page.controls:
                        page.update(discount_dialog_badge, discount_dialog_badge_shell, discount_dialog_status, discount_overview_row, discount_dialog_list, discount_requests_auto_button)
            else:
                restore_discount_item(item)
                render_discount_dialog()
                toast(str(result.get("message") or "Не удалось обработать заявку."), True)
                if page.controls:
                    page.update(discount_dialog_badge, discount_dialog_badge_shell, discount_dialog_status, discount_overview_row, discount_dialog_list, discount_requests_auto_button)

        page.run_thread(worker)

    def handle_discount_auto(_e=None) -> None:
        if discount_dialog_state["busy"] or discount_dialog_state.get("requires_refresh"):
            return
        plan = build_discount_auto_plan()
        if not plan.get("items"):
            render_discount_dialog()
            return
        if process_discount_request_plan is None:
            toast(f"Модуль обработки скидок недоступен: {PRICING_IMPORT_ERROR or 'ошибка импорта'}", True)
            return
        discount_dialog_state["busy"] = True
        render_discount_dialog()
        if page.controls:
            page.update(discount_dialog_badge, discount_dialog_badge_shell, discount_dialog_status, discount_overview_row, discount_dialog_list, discount_requests_auto_button)

        def worker() -> None:
            try:
                result = process_discount_request_plan(plan)
            except Exception as exc:
                result = {"ok": False, "message": f"Ошибка обработки заявок: {exc}"}
            discount_dialog_state["busy"] = False
            if result.get("ok"):
                toast(str(result.get("message") or "Автоматическая обработка завершена."))
                discount_dialog_state["plan"] = {"items": [], "ok": True, "message": str(result.get("message") or "")}
                render_discount_dialog()
                if page.controls:
                    page.update(discount_dialog_badge, discount_dialog_badge_shell, discount_dialog_status, discount_overview_row, discount_dialog_list, discount_requests_auto_button)
            else:
                discount_dialog_state["requires_refresh"] = True
                render_discount_dialog()
                toast(str(result.get("message") or "Не удалось автоматически обработать заявки."), True)
                # Aggregate API counts do not identify successful items reliably.
                # Reload before allowing a retry to avoid sending successful items twice.
                if page.controls:
                    page.update(discount_dialog_badge, discount_dialog_badge_shell, discount_dialog_status, discount_overview_row, discount_dialog_list, discount_requests_auto_button)

        page.run_thread(worker)

    def open_discount_requests_dialog(_e=None) -> None:
        if discount_dialog_state["busy"]:
            return
        if build_discount_request_plan is None:
            toast(f"Модуль обработки скидок недоступен: {PRICING_IMPORT_ERROR or 'ошибка импорта'}", True)
            return
        set_status("Заявки на скидку", "Собираю рекомендации по заявкам...", ACCENT, busy=True)
        discount_dialog_state["busy"] = True
        discount_dialog_state["plan"] = None
        discount_dialog_state["processing_ids"] = set()
        render_discount_dialog()
        if page.controls:
            page.update(discount_dialog_badge, discount_dialog_badge_shell, discount_dialog_status, discount_overview_row, discount_dialog_list, discount_requests_auto_button)

        def worker() -> None:
            try:
                plan = build_discount_request_plan(ROOT)
            except Exception as exc:
                plan = {"ok": False, "message": f"Не удалось загрузить заявки: {exc}", "items": []}
            set_status("Заявки на скидку", str(plan.get("message") or "Готово"), PRIMARY if plan.get("ok") else DANGER, busy=False)
            if not plan.get("ok"):
                discount_dialog_state["plan"] = {"items": [], "ok": False, "message": str(plan.get("message") or "")}
                discount_dialog_state["busy"] = False
                discount_dialog_state["processing_ids"] = set()
                render_discount_dialog()
                if page.controls:
                    page.update(discount_dialog_badge, discount_dialog_badge_shell, discount_dialog_status, discount_overview_row, discount_dialog_list, discount_requests_auto_button)
                toast(str(plan.get("message") or "Не удалось подготовить заявки."), True)
                return
            discount_dialog_state["plan"] = plan
            discount_dialog_state["requires_refresh"] = False
            discount_dialog_state["busy"] = False
            discount_dialog_state["processing_ids"] = set()
            render_discount_dialog()
            apply_discount_dialog_theme()
            if page.controls:
                page.update(discount_dialog_badge, discount_dialog_badge_shell, discount_dialog_status, discount_overview_row, discount_dialog_list, discount_requests_auto_button)
            if not plan.get("items"):
                toast(str(plan.get("message") or "Новых заявок нет."))

        page.run_thread(worker)
    discount_requests_refresh_button.on_click = open_discount_requests_dialog
    discount_requests_auto_button.on_click = handle_discount_auto

    clipboard_service = ft.Clipboard()
    page.services.append(clipboard_service)

    async def copy_log_to_clipboard_async() -> None:
        try:
            await clipboard_service.set(get_full_log_text())
            toast("Полный лог скопирован в буфер обмена.")
        except Exception as exc:
            toast(f"Не удалось скопировать лог: {exc}", True)

    def copy_log_to_clipboard(_e=None) -> None:
        page.run_task(copy_log_to_clipboard_async)

    def clear_log(_e=None) -> None:
        set_log("Отображение журнала очищено. Полная история сохранена в файле.")
        if page.controls:
            page.update(log, log_selectable)

    log_dialog = ft.AlertDialog(
        modal=False,
        title=ft.Text("Операционный лог"),
        content=ft.Container(content=log_selectable, width=920, height=620),
        actions=[
            ft.TextButton("Открыть файл журнала", on_click=lambda _e: reveal_path(log_store.path)),
            ft.TextButton("Очистить лог", on_click=clear_log),
            ft.TextButton("Копировать лог", on_click=copy_log_to_clipboard),
            ft.TextButton("Закрыть", on_click=lambda _e: close_log_dialog()),
        ],
        actions_alignment=ft.MainAxisAlignment.END,
        open=False,
    )

    def open_log_dialog() -> None:
        log_dialog.open = True
        page.update(log_dialog, log_selectable)

    def close_log_dialog() -> None:
        log_dialog.open = False
        page.update(log_dialog)

    page.overlay.append(log_dialog)

    def handle_keyboard(e: ft.KeyboardEvent) -> None:
        if e.key == "F2":
            if not ai_state["debug_enabled"]:
                toast("Включите Debug режим через «Инструменты», чтобы открыть лог.", True)
                return
            if log_dialog.open:
                close_log_dialog()
            else:
                open_log_dialog()

    page.on_keyboard_event = handle_keyboard

    async def ui_event_loop() -> None:
        while True:
            processed = False
            while True:
                try:
                    event = ui_queue.get_nowait()
                except queue.Empty:
                    break
                processed = True
                kind = event[0]
                if kind == "log":
                    current = event[1]
                    set_log(current)
                    page.update(log, log_selectable, state, spin)
                elif kind == "log_append":
                    append_log(event[1])
                    page.update(log, log_selectable, state, spin)
                elif kind == "ai_step":
                    _, title, summary, phase_payload = event
                    if phase_payload:
                        phase, phase_detail = phase_payload
                        ai_set_phase(phase, phase_detail)
                    ai_add_trace(title, summary)
                    page.update(ai_trace, ai_phase_label, ai_phase_detail)
                elif kind == "ai_stream_start":
                    _, request_id, note = event
                    if request_id in ai_state["cancelled_request_ids"]:
                        continue
                    existing_payload = ai_state["stream_controls"].get(request_id)
                    if isinstance(existing_payload, dict):
                        note_control = existing_payload.get("note")
                        if note_control is not None:
                            note_control.value = note
                    else:
                        ai_state["stream_controls"][request_id] = ai_append_streaming_placeholder(note)
                    ai_sync_layout()
                    page.update(ai_body, ai_messages, ai_messages_shell, ai_welcome_shell, ai_composer_shell)
                elif kind == "ai_thinking_chunk":
                    _, request_id, chunk = event
                    if request_id in ai_state["cancelled_request_ids"]:
                        continue
                    stream_payload = ai_state["stream_controls"].get(request_id)
                    thinking_list = stream_payload.get("thinking_list") if isinstance(stream_payload, dict) else None
                    thinking_current = stream_payload.get("thinking_current") if isinstance(stream_payload, dict) else None
                    thinking_anchor = stream_payload.get("thinking_anchor") if isinstance(stream_payload, dict) else None
                    thinking_toggle = stream_payload.get("thinking_toggle") if isinstance(stream_payload, dict) else None
                    thinking_wrap = stream_payload.get("thinking_wrap") if isinstance(stream_payload, dict) else None
                    if thinking_list is not None and thinking_current is not None and thinking_wrap is not None and thinking_anchor is not None:
                        parts = chunk.split("\n")
                        for idx, part in enumerate(parts):
                            if idx == 0:
                                thinking_current.value = (thinking_current.value or "") + part
                            else:
                                thinking_current = ft.Text(part, size=12, color=ai_surface_tokens()["muted"], selectable=True)
                                thinking_list.controls.insert(max(len(thinking_list.controls) - 1, 0), thinking_current)
                        stream_payload["thinking_current"] = thinking_current
                        if thinking_anchor in thinking_list.controls:
                            thinking_list.controls.remove(thinking_anchor)
                        thinking_list.controls.append(thinking_anchor)
                        thinking_wrap.visible = True
                        if thinking_toggle is not None:
                            thinking_toggle.visible = True
                        page.update(ai_messages, thinking_wrap, thinking_list)
                elif kind == "ai_stream_chunk":
                    _, request_id, chunk = event
                    if request_id in ai_state["cancelled_request_ids"]:
                        continue
                    stream_payload = ai_state["stream_controls"].get(request_id)
                    text_control = stream_payload.get("text") if isinstance(stream_payload, dict) else None
                    if text_control is not None:
                        text_control.value = (text_control.value or "") + chunk
                        page.update(ai_messages)
                elif kind == "prompt":
                    _, label, source_line, holder, waiter = event
                    prompt_state["holder"] = holder
                    prompt_state["event"] = waiter
                    period_label = infer_prompt_period(source_line)
                    prompt_label.value = (
                        f"Введите сумму расходов на маркетинг для отчёта за {period_label}."
                        if period_label
                        else "Введите сумму расходов на маркетинг для отчёта, который сейчас строится."
                    )
                    prompt_hint.value = "Укажите общую сумму в рублях. Если расходов не было, оставьте 0."
                    prompt_context.value = f"Источник запроса: {source_line}" if source_line else ""
                    prompt_input.label = f"{label} (руб.)" if label else "Сумма маркетинга (руб.)"
                    prompt_input.value = "0"
                    apply_prompt_dialog_theme()
                    set_status("Ожидается ввод", period_label or label, ACCENT, busy=False)
                    prompt_dialog.open = True
                    page.update(prompt_label, prompt_hint, prompt_context, prompt_input, prompt_dialog, log)
                elif kind == "finish":
                    _, title, code, output = event
                    finish(title, code, output)
                elif kind == "ai_finish":
                    _, request_id, prompt, answer, meta = event
                    if request_id in ai_state["cancelled_request_ids"]:
                        ai_state["stream_controls"].pop(request_id, None)
                        ai_state["request_started_at"].pop(request_id, None)
                        ai_state["cancelled_request_ids"].discard(request_id)
                        page.update(log, log_selectable)
                        continue
                    if ai_state["active_request_id"] not in (None, request_id):
                        continue
                    ai_state["active_request_id"] = None
                    elapsed = None
                    started_at = ai_state["request_started_at"].pop(request_id, None)
                    if started_at is not None:
                        elapsed = max(0.0, time.perf_counter() - started_at)
                    elapsed_note = ai_format_elapsed(elapsed)
                    streamed_payload = ai_state["stream_controls"].pop(request_id, None)
                    streamed_control = streamed_payload.get("text") if isinstance(streamed_payload, dict) else None
                    streamed_host = streamed_payload.get("content_host") if isinstance(streamed_payload, dict) else None
                    streamed_note = streamed_payload.get("note") if isinstance(streamed_payload, dict) else None
                    streamed_thinking_list = streamed_payload.get("thinking_list") if isinstance(streamed_payload, dict) else None
                    streamed_toggle = streamed_payload.get("thinking_toggle") if isinstance(streamed_payload, dict) else None
                    streamed_wrap = streamed_payload.get("thinking_wrap") if isinstance(streamed_payload, dict) else None
                    if answer:
                        if streamed_control is not None:
                            streamed_control.value = answer
                            if streamed_host is not None:
                                streamed_host.content = ai_markdown(answer)
                            if streamed_note is not None:
                                streamed_note.value = elapsed_note or ""
                            if streamed_thinking_list is not None and streamed_wrap is not None:
                                has_thinking = any(isinstance(ctrl, ft.Text) and (ctrl.value or "").strip() for ctrl in streamed_thinking_list.controls)
                                if not has_thinking:
                                    streamed_wrap.visible = False
                                elif streamed_toggle is not None:
                                    streamed_toggle.visible = True
                        else:
                            ai_append_message("assistant", answer, elapsed_note)
                        ai_state["conversation_history"].append({"role": "user", "content": prompt})
                        ai_state["conversation_history"].append({"role": "assistant", "content": answer})
                        if len(ai_state["conversation_history"]) > 6:
                            ai_state["conversation_history"] = ai_state["conversation_history"][-6:]
                        ai_set_busy(False, "Ответ готов. Можно продолжать разбор месяца.")
                        set_status("AI ответил", "Локальный анализ завершён.", PRIMARY, busy=False)
                    else:
                        ai_append_message("assistant", str(meta), elapsed_note or "Диагностика runtime")
                        ai_set_busy(False, str(meta))
                        set_status("AI недоступен", str(meta), DANGER, busy=False)
                    ai_sync_layout()
                    page.update(ai_body, ai_messages, ai_messages_shell, ai_welcome_shell, ai_composer_shell, ai_busy_label, state, status_detail, spin, cancel_button)
            await asyncio.sleep(0.01 if processed else 0.03)

    ai_refresh_status(push_update=False)
    ai_refresh_tools_menu()
    ai_apply_chat_theme()

    dash = ft.Column(
        [
            ft.Text("Бизнес-сводка", size=26, weight=ft.FontWeight.W_800, color=TEXT),
            ft.Container(
                padding=18,
                border_radius=26,
                bgcolor=semantic_surface("neutral_soft", dark=False),
                border=ft.Border.all(1, BORDER),
                content=ft.Column(
                    [
                        ft.Text("Период", size=16, weight=ft.FontWeight.W_700, color=TEXT),
                        ft.Row([dashboard_m, dashboard_y], spacing=10),
                        dashboard_freshness,
                        ft.Row(
                            [
                                ft.Button("Обновить данные", bgcolor=PRIMARY, color="white", on_click=lambda _e: launch_dashboard_report()),
                                ft.TextButton("Открыть отчёт", on_click=lambda _e: open_selected_dashboard_report()),
                                ft.TextButton("Открыть папку с отчётами", on_click=lambda _e: reveal_path(ROOT / "reports", "reports")),
                            ],
                            spacing=8,
                        ),
                    ],
                    spacing=10,
                ),
            ),
            dashboard_summary,
            dashboard_metric_wrap,
            dashboard_details,
            dashboard_promotion,
            ft.Container(
                padding=18,
                border_radius=26,
                bgcolor=semantic_surface("neutral", dark=False),
                border=ft.Border.all(1, BORDER),
                content=ft.Column(
                    [
                        ft.Text("Динамика", size=16, weight=ft.FontWeight.W_700, color=TEXT),
                        ft.Container(
                            padding=14,
                            border_radius=22,
                            bgcolor=semantic_surface("neutral_soft", dark=False),
                            border=ft.Border.all(1, BORDER),
                            content=ft.ResponsiveRow(
                                run_spacing=10,
                                spacing=10,
                                controls=[
                                    ft.Container(content=dashboard_chart_type, col=3),
                                    ft.Container(content=dashboard_chart_from_m, col=1.85),
                                    ft.Container(content=dashboard_chart_from_y, col=1.15),
                                    ft.Container(content=dashboard_chart_to_m, col=1.85),
                                    ft.Container(content=dashboard_chart_to_y, col=1.15),
                                    ft.Container(
                                        col=3,
                                        alignment=ft.Alignment(1, 0),
                                        content=ft.Row(
                                            [
                                                ft.Button("Построить график", bgcolor=PRIMARY, color="white", on_click=build_dashboard_chart),
                                                ft.TextButton("Интерактивный", on_click=open_interactive_dashboard_chart),
                                            ],
                                            spacing=6,
                                            alignment=ft.MainAxisAlignment.END,
                                            vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                            wrap=True,
                                        ),
                                    ),
                                ],
                            ),
                        ),
                        dashboard_chart_panel,
                    ],
                    spacing=12,
                ),
            ),
        ],
        spacing=14,
    )

    reports = ft.Column(
        [
            ft.Text("Отчёты", size=26, weight=ft.FontWeight.W_800, color=TEXT),
            ft.ResponsiveRow(
                [
                    card(
                        "Месячный отчёт",
                        "Сборка monthly Excel за выбранный период.",
                        [
                            ft.Container(
                                height=152,
                                content=ft.Column(
                                    [
                                        ft.Row([monthly_m, monthly_y], spacing=12),
                                        ft.Row(
                                            [
                                                ft.Button("Сформировать отчёт", bgcolor=PRIMARY, color="white", on_click=lambda _e: launch_monthly_report()),
                                                ft.TextButton("Папка reports", on_click=lambda _e: reveal_path(ROOT / "reports", "reports")),
                                            ],
                                            spacing=10,
                                        ),
                                    ],
                                    spacing=12,
                                    alignment=ft.MainAxisAlignment.START,
                                ),
                            ),
                            ft.Container(height=1, bgcolor="#E5E5EA", border_radius=10),
                            ft.Row(
                                [
                                    ft.Text("Готовые файлы", size=12, weight=ft.FontWeight.W_700, color=MUTED, expand=True),
                                    ft.IconButton(ft.Icons.REFRESH, tooltip="Обновить список", on_click=refresh_report_lists, icon_color=MUTED),
                                ],
                                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                            ),
                            monthly_reports_list,
                        ],
                        tone=semantic_surface("neutral", dark=False),
                        col=6,
                    ),
                    card(
                        "ABC/XYZ аналитика",
                        "Анализ диапазона на основе monthly reports.",
                        [
                            ft.Container(
                                height=152,
                                content=ft.Column(
                                    [
                                        ft.Row([abc_fm, abc_fy], spacing=12),
                                        ft.Row([abc_tm, abc_ty], spacing=12),
                                        ft.Row(
                                            [
                                                ft.Button(
                                                    "Провести ABC/XYZ анализ",
                                                    bgcolor=PRIMARY,
                                                    color="white",
                                                    on_click=lambda _e: launch_abc_analysis(),
                                                ),
                                                ft.TextButton("Папка ABC&XYZ", on_click=lambda _e: reveal_path(ROOT / "ABC&XYZ reports", "ABC&XYZ reports")),
                                            ],
                                            spacing=10,
                                        ),
                                    ],
                                    spacing=12,
                                    alignment=ft.MainAxisAlignment.START,
                                ),
                            ),
                            ft.Container(height=1, bgcolor="#E5E5EA", border_radius=10),
                            ft.Row(
                                [
                                    ft.Text("Готовые файлы", size=12, weight=ft.FontWeight.W_700, color=MUTED, expand=True),
                                    ft.IconButton(ft.Icons.REFRESH, tooltip="Обновить список", on_click=refresh_report_lists, icon_color=MUTED),
                                ],
                                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                            ),
                            abc_reports_list,
                        ],
                        tone=semantic_surface("neutral_soft", dark=False),
                        col=6,
                    ),
                ],
                run_spacing=16,
                spacing=16,
            ),
        ],
        spacing=16,
    )

    def pricing_round_action(icon: str, tooltip: str, on_click, *, tone: str = "neutral_alt") -> ft.Control:
        return ft.Container(
            width=42,
            height=42,
            border_radius=999,
            bgcolor=semantic_surface(tone, dark=False),
            border=ft.Border.all(1, "#E5E5EA"),
            alignment=ft.Alignment(0, 0),
            content=ft.IconButton(
                icon=icon,
                tooltip=tooltip,
                on_click=on_click,
                icon_size=18,
                style=ft.ButtonStyle(shape=ft.CircleBorder(), padding=ft.Padding.all(10)),
            ),
        )

    discount_requests_surface = ft.Container(
        height=540,
        padding=4,
        border_radius=22,
        bgcolor=semantic_surface("neutral_soft", dark=False),
        border=ft.Border.all(1, BORDER),
        content=discount_dialog_list,
    )
    pricing_costs_surface = ft.Container(
        height=760,
        padding=4,
        border_radius=22,
        bgcolor=semantic_surface("neutral_soft", dark=False),
        border=ft.Border.all(1, BORDER),
        content=pricing_costs_list,
    )

    pricing = ft.Column(
        [
            ft.Text("Цены и скидки", size=26, weight=ft.FontWeight.W_800, color=TEXT),
            card(
                "",
                "",
                [
                    ft.ResponsiveRow(
                        run_spacing=12,
                        spacing=12,
                        controls=[
                            ft.Container(
                                col=8,
                                content=ft.Column(
                                    [
                                        ft.Text("Диапазон рентабельности", size=15, weight=ft.FontWeight.W_700, color=TEXT),
                                        ft.Text("Маржа задаёт базовый контур расчёта. Пересчёт цен выполняется прямо над целевыми столбцами таблицы.", size=11, color=MUTED),
                                        ft.Row(
                                            [
                                                ft.Container(content=margin_min, width=190),
                                                ft.Container(content=margin_des, width=190),
                                                ft.Button(
                                                    "Сохранить маржу",
                                                    bgcolor=PRIMARY,
                                                    color="white",
                                                    height=42,
                                                    on_click=lambda _e: toast(valid_margin(margin_min.value or "", margin_des.value or ""), True)
                                                    if valid_margin(margin_min.value or "", margin_des.value or "")
                                                    else job(
                                                        "Сохранение маржи",
                                                        [
                                                            str(PYTHON),
                                                            str(SCRIPTS / "ui_actions.py"),
                                                            "save-margin",
                                                            "--min-margin",
                                                            (margin_min.value or "").replace(",", "."),
                                                            "--desired-margin",
                                                            (margin_des.value or "").replace(",", "."),
                                                        ],
                                                    ),
                                                ),
                                            ],
                                            spacing=12,
                                            wrap=True,
                                            vertical_alignment=ft.CrossAxisAlignment.END,
                                        ),
                                    ],
                                    spacing=6,
                                    tight=True,
                                ),
                            ),
                            ft.Container(
                                col=4,
                                content=ft.Column(
                                    [
                                        ft.Text("Сервисы", size=15, weight=ft.FontWeight.W_700, color=TEXT),
                                        ft.Text("Файл, корень проекта и обновление минимальных цен в кабинете.", size=11, color=MUTED),
                                        ft.Row(
                                            [
                                                pricing_round_action(
                                                    ft.Icons.OPEN_IN_NEW,
                                                    "Открыть costs.xlsx",
                                                    lambda _e: reveal_path(ROOT / "costs.xlsx", "costs.xlsx"),
                                                ),
                                                pricing_round_action(
                                                    ft.Icons.FOLDER_OPEN,
                                                    "Открыть корень проекта",
                                                    lambda _e: reveal_path(ROOT, "корень проекта"),
                                                ),
                                                pricing_round_action(
                                                    ft.Icons.SYNC,
                                                    "Обновить минимальные цены в Ozon",
                                                    lambda _e: job("Обновление минимальных цен", [str(PYTHON), str(SCRIPTS / "ui_actions.py"), "update-min-prices"]),
                                                    tone="warning",
                                                ),
                                            ],
                                            spacing=8,
                                        ),
                                        ft.Text(
                                            "Кнопки, которые меняют значения в столбцах, находятся прямо над этими столбцами в рабочей таблице.",
                                            size=10,
                                            color=MUTED,
                                        ),
                                    ],
                                    spacing=6,
                                    tight=True,
                                ),
                            ),
                        ],
                    ),
                ],
                tone=semantic_surface("neutral_soft", dark=False),
                col=12,
            ),
            ft.ResponsiveRow(
                run_spacing=14,
                spacing=14,
                controls=[
                    card(
                        "",
                        "",
                        [
                            ft.ResponsiveRow(
                                run_spacing=10,
                                spacing=10,
                                controls=[
                                    ft.Container(
                                        col=7,
                                        content=ft.Column(
                                            [
                                                ft.Text("Заявки на скидку", size=20, weight=ft.FontWeight.W_700, color=TEXT),
                                                ft.Text("Рабочая очередь решений по скидочным заявкам.", size=11, color=MUTED),
                                            ],
                                            spacing=4,
                                            tight=True,
                                        ),
                                    ),
                                    ft.Container(
                                        col=5,
                                        alignment=ft.Alignment(1, 0),
                                        content=ft.Row(
                                            [
                                                discount_dialog_badge_shell,
                                                ft.Container(
                                                    width=42,
                                                    height=42,
                                                    border_radius=999,
                                                    bgcolor=semantic_surface("neutral_alt", dark=False),
                                                    border=ft.Border.all(1, "#E5E5EA"),
                                                    alignment=ft.Alignment(0, 0),
                                                    content=discount_requests_refresh_button,
                                                ),
                                            ],
                                            spacing=8,
                                            wrap=True,
                                            alignment=ft.MainAxisAlignment.END,
                                        ),
                                    ),
                                ],
                            ),
                            ft.Row([discount_requests_auto_button], spacing=8, wrap=True),
                            discount_dialog_status,
                            discount_overview_row,
                            discount_requests_surface,
                        ],
                        tone=semantic_surface("neutral", dark=False),
                        col={"md": 12, "lg": 3, "xl": 3},
                    ),
                    card(
                        "",
                        "",
                        [
                            ft.ResponsiveRow(
                                run_spacing=10,
                                spacing=10,
                                controls=[
                                    ft.Container(
                                        col=6,
                                        content=ft.Column(
                                            [
                                                ft.Text("Рабочая таблица", size=20, weight=ft.FontWeight.W_700, color=TEXT),
                                                ft.Text("Правка себестоимости и контроль цены в одном рабочем поле.", size=11, color=MUTED),
                                            ],
                                            spacing=4,
                                            tight=True,
                                        ),
                                    ),
                                    ft.Container(
                                        col=6,
                                        alignment=ft.Alignment(1, 0),
                                        content=ft.Row(
                                            [
                                                pricing_costs_status_shell,
                                                pricing_round_action(
                                                    ft.Icons.SYNC,
                                                    "Обновить таблицу",
                                                    handle_pricing_costs_refresh,
                                                ),
                                                pricing_round_action(
                                                    ft.Icons.OPEN_IN_NEW,
                                                    "Открыть costs.xlsx",
                                                    lambda _e: reveal_path(ROOT / "costs.xlsx", "costs.xlsx"),
                                                ),
                                            ],
                                            spacing=8,
                                            wrap=True,
                                            alignment=ft.MainAxisAlignment.END,
                                        ),
                                    ),
                                ],
                            ),
                            ft.Row(
                                [
                                    pricing_costs_save_button,
                                    ft.Text("Кнопки над колонками влияют только на соответствующие столбцы.", size=11, color=MUTED),
                                ],
                                spacing=8,
                                wrap=True,
                                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                            ),
                            pricing_costs_surface,
                        ],
                        tone=semantic_surface("neutral", dark=False),
                        col={"md": 12, "lg": 9, "xl": 9},
                    ),
                ],
            ),
        ],
        spacing=18,
    )

    actions = ft.Column(
        [
            ft.Text("Акции", size=26, weight=ft.FontWeight.W_800, color=TEXT),
            ft.Text("Продвинутые промо-операции Ozon: удаление невыгодных и добавление выгодных товаров в акции.", size=13, color=MUTED),
            ft.ResponsiveRow(
                run_spacing=14,
                spacing=14,
                controls=[
                    card(
                        "Акции Ozon",
                        "Продвинутые промо-операции: удалить невыгодные или добавить выгодные товары.",
                        [
                            ft.Column(
                                [
                                    ft.Button("Удалить невыгодные", bgcolor=DANGER, color="white", on_click=lambda _e: job("Удаление невыгодных акций", [str(PYTHON), str(SCRIPTS / "ui_actions.py"), "remove-unprofitable-actions"])),
                                    ft.Button("Добавить в акции", bgcolor=PRIMARY, color="white", on_click=lambda _e: job("Добавление в акции", [str(PYTHON), str(SCRIPTS / "ui_actions.py"), "add-to-actions"])),
                                ],
                                spacing=8,
                            )
                        ],
                        width=420,
                    ),
                ],
            ),
        ],
        spacing=18,
    )

    supply = ft.Column(
        [
            ft.Text("Поставки", size=26, weight=ft.FontWeight.W_800, color=TEXT),
            ft.Text("FBO supply planning на базе ABC/XYZ и продаж за 90 дней.", size=13, color=MUTED),
            card(
                "Расчёт поставки FBO",
                "Автоматически использует последние 3 полных месяца и формирует Excel в stocks reports.",
                [
                    ft.Row(
                        [
                            ft.Button("Рассчитать поставку", bgcolor=PRIMARY, color="white", on_click=lambda _e: job("Расчёт поставки FBO", [str(PYTHON), str(SCRIPTS / "fbo_supply_report.py")])),
                            ft.TextButton("Открыть stocks reports", on_click=lambda _e: reveal_path(ROOT / "stocks reports", "stocks reports")),
                        ]
                    )
                ],
                tone="#EAF0F5",
            ),
        ],
        spacing=18,
    )

    finance = ft.Column(
        [
            ft.Text("Финансы", size=26, weight=ft.FontWeight.W_800, color=TEXT),
            ft.Text("Выгрузка баланса Ozon с ограничением до 30 дней.", size=13, color=MUTED),
            card(
                "Отчёт по балансу",
                "Сохраняет JSON и Excel в balance reports.",
                [
                    ft.Row([bal_from, bal_to]),
                    ft.Row(
                        [
                            ft.Button(
                                "Скачать баланс",
                                bgcolor=PRIMARY,
                                color="white",
                                on_click=lambda _e: toast(valid_balance(bal_from.value or "", bal_to.value or ""), True)
                                if valid_balance(bal_from.value or "", bal_to.value or "")
                                else job("Баланс", [str(PYTHON), str(SCRIPTS / "balance_report.py"), "--date_from", bal_from.value or "", "--date_to", bal_to.value or ""]),
                            ),
                            ft.TextButton("Открыть balance reports", on_click=lambda _e: reveal_path(ROOT / "balance reports", "balance reports")),
                        ]
                    ),
                ],
                tone="#EEF0F8",
            ),
        ],
        spacing=18,
    )

    def ai_quick_prompt(prompt_text: str):
        return lambda _e: ai_submit(prompt_text)

    ai_input.on_submit = lambda _e: ai_submit(ai_input.value or "")
    ai_send_button.on_click = lambda _e: ai_submit(ai_input.value or "")
    ai_reset_button.on_click = ai_reset
    ai_refresh_button.on_click = lambda _e: ai_refresh_status()
    ai_new_chat_button.on_click = ai_reset
    ai_model_select.on_change = ai_model_changed
    ai_quick_month_button.on_click = ai_quick_prompt("Что произошло в марте 2026?")
    ai_quick_pressure_button.on_click = ai_quick_prompt("Что сильнее всего давит на прибыль в марте 2026?")
    ai_quick_brief_button.on_click = ai_quick_prompt("Сделай краткий управленческий разбор марта 2026.")
    ai_quick_risk_button.on_click = ai_quick_prompt("Какие риски ты видишь в месячной экономике марта 2026?")

    ai_messages_shell = ft.Container(
        expand=True,
        padding=ft.Padding.only(left=28, right=28, top=18, bottom=12),
        content=ai_messages,
        visible=False,
    )
    ai_composer_card = ft.Container(
        padding=ft.Padding.only(left=18, right=14, top=14, bottom=12),
        border_radius=30,
        bgcolor=ai_surface_tokens()["composer"],
        border=ft.Border.all(1, ai_surface_tokens()["composer_border"]),
        shadow=ft.BoxShadow(
            blur_radius=28,
            spread_radius=0,
            color="#1A000000" if not current_dark["value"] else "#22000000",
            offset=ft.Offset(0, 10),
        ),
        content=ft.Column(
            [
                ft.Row(
                    [
                        ai_input,
                        ai_send_button,
                    ],
                    vertical_alignment=ft.CrossAxisAlignment.END,
                    spacing=10,
                ),
                ft.Row(
                    [
                        ft.Row([ai_busy_ring, ai_busy_label], spacing=8),
                        ft.Row(
                            [
                                ft.Container(content=ai_model_select, width=180),
                                ft.Row([ai_status_dot, ai_status_label], spacing=8),
                            ],
                            spacing=14,
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                ),
            ],
            spacing=8,
        ),
    )
    ai_composer_shell = ft.Container(
        padding=ft.Padding.only(left=28, right=28, top=10, bottom=24),
        content=ai_centered(ai_composer_card),
    )
    ai_welcome_shell = ft.Container(
        expand=True,
        alignment=ft.Alignment(0, 0),
        content=ft.Column(
            [
                ai_empty_state,
            ],
            spacing=28,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            alignment=ft.MainAxisAlignment.CENTER,
        ),
    )
    ai_body = ft.Column(expand=True, spacing=0)

    def ai_sync_layout() -> None:
        active = ai_has_conversation()
        ai_welcome_shell.visible = not active
        ai_messages_shell.visible = active
        ai_messages_shell.expand = active
        ai_composer_shell.padding = ft.Padding.only(left=24, right=24, top=6, bottom=20 if active else 0)
        ai_body.controls = [ai_messages_shell, ai_composer_shell] if active else [ai_welcome_shell, ai_composer_shell]

    ai_sync_layout()

    ai_header_card = ft.Container(
        padding=ft.Padding.symmetric(horizontal=20, vertical=16),
        border_radius=26,
        bgcolor=ai_surface_tokens()["assistant_bg"],
        border=ft.Border.all(1, ai_surface_tokens()["assistant_border"]),
        content=ft.Row(
            [
                ft.Column(
                    [
                        ai_top_title,
                        ai_top_subtitle,
                    ],
                    spacing=2,
                ),
                ft.Row(
                    [
                        ai_new_chat_button,
                    ],
                    spacing=4,
                ),
            ],
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        ),
    )
    ai_view = ft.Container(
        expand=True,
        bgcolor=ai_surface_tokens()["canvas"],
        padding=0,
        content=ft.Column(
            [
                ft.Container(
                    padding=ft.Padding.only(left=28, right=28, top=20, bottom=6),
                    content=ai_centered(ai_header_card),
                ),
                ai_body,
            ],
            spacing=0,
            expand=True,
        ),
    )

    def open_marketplace_settings(_e=None):
        current = read_settings(ROOT)
        fields = {
            key: ft.TextField(
                label=label, value=current.get(key) or "", password=secret,
                can_reveal_password=secret,
            ) for key, label, secret in FIELDS
        }
        feedback = ft.Text(color="#C62828")

        def close_setup(_e=None):
            setup_dialog.open = False
            page.update()
            page.overlay.remove(setup_dialog)

        def save_setup(_e):
            if job_lock.locked() or discount_dialog_state["busy"] or ai_state["busy"]:
                feedback.value = "Дождитесь завершения текущей операции перед сменой ключей."
                page.update()
                return
            try:
                save_settings(ROOT, {key: field.value or "" for key, field in fields.items()})
            except ValueError as exc:
                feedback.value = str(exc)
                page.update()
                return
            except Exception:
                feedback.value = "Не удалось сохранить настройки. Проверьте доступ к .env и повторите."
                page.update()
                return
            for module_name in ("recommended_prices", "scripts.recommended_prices"):
                module = sys.modules.get(module_name)
                if module is not None:
                    module.refresh_api_credentials()
            close_setup()
            set_status("Настройки сохранены", "Ключи Ozon и WB сохранены. Проверка подключения не выполнялась.", PRIMARY, busy=False)

        setup_dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("Подключение Ozon и Wildberries"),
            content=ft.Column(
                [ft.Text(HELP, size=13),
                 ft.Text("Заполните данные нужных маркетплейсов. Остальные поля можно оставить пустыми. Ключи сохраняются локально в .env.", size=13),
                 *fields.values(), feedback],
                width=600, height=490, scroll=ft.ScrollMode.AUTO, spacing=12,
            ),
            actions=[ft.TextButton("Позже", on_click=close_setup),
                     ft.Button("Сохранить", on_click=save_setup)],
        )
        page.overlay.append(setup_dialog)
        setup_dialog.open = True
        page.update()

    def open_balance_check(_e=None):
        today = date.today()
        month_field = ft.Dropdown(label="Месяц", value=str(today.month), width=200,
                                  options=[ft.dropdown.Option(str(i), title) for i, title in enumerate(MONTHS, 1)])
        year_field = ft.Dropdown(label="Год", value=str(today.year), width=140,
                                 options=[ft.dropdown.Option(str(y)) for y in range(today.year, 2023, -1)])
        result_column = ft.Column(spacing=6)
        status_text = ft.Text("Выберите месяц и нажмите «Проверить».", size=12, color=MUTED)
        checking = {"value": False}
        recent_column = ft.Column(spacing=10)
        recent_status = ft.Text("Загружаем последние 3 месяца…", size=12, color=MUTED)
        recent_checking = {"value": False}

        def render_result(summary):
            if not summary:
                result_column.controls = [ft.Text("Ozon не вернул данных о балансе за этот период.", size=13, color=MUTED)]
                return
            def row(label, value):
                return ft.Text(f"{label}: {_format_currency(value)}", size=13)
            result_column.controls = [
                row("Входящий баланс на начало периода", summary.get("opening_balance")),
                row("Начислено Ozon за период", summary.get("accrued")),
                row("Выплачено Ozon (реальный перевод)", summary.get("payments")),
                row("Исходящий баланс на конец периода", summary.get("closing_balance")),
                row("Комиссия за ранний вывод средств", summary.get("early_payment_fee")),
            ]

        async def check_balance():
            if checking["value"]:
                return
            checking["value"] = True
            check_button.disabled = True
            status_text.value = "Запрашиваем баланс Ozon…"
            result_column.controls = []
            page.update()
            try:
                month, year = int(month_field.value), int(year_field.value)
                summary = await asyncio.to_thread(balance_report.summarize_month, month, year)
                render_result(summary)
                status_text.value = f"{MONTHS[month - 1]} {year} · из раздела Ozon «Финансы → Баланс»."
            except Exception as exc:
                status_text.value = f"Не удалось получить баланс: {exc}"
                result_column.controls = []
            finally:
                checking["value"] = False
                check_button.disabled = False
            page.update()

        def load_one_period(month: int, year: int) -> dict:
            try:
                summary = balance_report.summarize_month(month, year) or {}
            except Exception as exc:
                summary = {"error": str(exc)}
            path = report_path_for_period(month, year)
            metrics, built_at = read_business_metrics(path) if path.exists() else ({}, None)
            return {"month": month, "year": year, "summary": summary, "metrics": metrics, "built_at": built_at}

        async def load_recent_periods():
            if recent_checking["value"]:
                return
            recent_checking["value"] = True
            recent_status.value = "Загружаем последние 3 месяца…"
            recent_column.controls = []
            page.update()
            try:
                periods = iter_periods(today.month, today.year, 3)
                results = await asyncio.to_thread(lambda: [load_one_period(m, y) for m, y in periods])
                cards = []
                for entry in reversed(results):
                    summary = entry["summary"]
                    metrics = entry["metrics"]
                    title = f"{MONTHS[entry['month'] - 1]} {entry['year']}"
                    if summary.get("error"):
                        cards.append(ft.Text(f"{title}: не удалось получить баланс ({summary['error']})", size=13, color=DANGER))
                        continue
                    received = summary.get("payments")
                    cost = metrics.get("Итоговая себестоимость")
                    profit = metrics.get("Чистая прибыль")
                    lines = [
                        ft.Text(title, size=15, weight=ft.FontWeight.W_700, color=TEXT),
                        ft.Text(f"Получено от Ozon (реальный перевод): {_format_currency(received)}", size=13),
                    ]
                    if metrics:
                        lines.append(ft.Text(f"Себестоимость к отправке (по отчёту): {_format_currency(abs(as_float(cost) or 0.0) if cost is not None else None)}", size=13))
                        lines.append(ft.Text(f"Прибыль оставить себе (по отчёту): {_format_currency(profit)}", size=13))
                        pending = metrics.get("Заказы, ожидающие расчёта Ozon (не учтены в прибыли/себестоимости)")
                        if pending:
                            lines.append(ft.Text(f"⚠ Ещё {pending} заказ(ов) не досчитаны Ozon — цифры за месяц могут подрасти позже.", size=12, color=ACCENT))
                    else:
                        lines.append(ft.Text("Отчёт за этот месяц не сформирован — себестоимость и прибыль неизвестны. Сформируйте отчёт на вкладке «Бизнес-сводка».", size=12, color=ACCENT))
                    cards.append(ft.Container(
                        content=ft.Column(lines, spacing=4),
                        padding=12, border_radius=10, bgcolor=semantic_surface("neutral", dark=current_dark["value"]),
                    ))
                recent_column.controls = cards
                recent_status.value = "Готово. «Получено от Ozon» — реальные деньги; себестоимость/прибыль — из уже сформированных месячных отчётов."
            except Exception as exc:
                recent_status.value = f"Не удалось загрузить последние месяцы: {exc}"
            finally:
                recent_checking["value"] = False
            page.update()

        def close_balance_dialog(_e=None):
            balance_dialog.open = False
            page.update()
            page.overlay.remove(balance_dialog)

        check_button = ft.Button("Проверить", on_click=lambda _e: page.run_task(check_balance))
        balance_dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("Баланс Ozon и сверка по месяцам"),
            content=ft.Column(
                [
                    ft.Text("Последние 3 месяца: сколько реально пришло от Ozon и сколько из этого — себестоимость к отправке, а сколько — ваша прибыль.", size=12, color=MUTED),
                    recent_status, recent_column,
                    ft.Divider(),
                    ft.Text("Проверить конкретный месяц отдельно (только баланс Ozon, без отчёта):", size=12, color=MUTED),
                    ft.Row([month_field, year_field, check_button], wrap=True),
                    status_text, result_column,
                ],
                width=560, height=560, scroll=ft.ScrollMode.AUTO, spacing=12,
            ),
            actions=[ft.TextButton("Закрыть", on_click=close_balance_dialog)],
        )
        page.overlay.append(balance_dialog)
        balance_dialog.open = True
        page.update()
        page.run_task(load_recent_periods)

    settings = ft.Column(
        [
            ft.Text("Настройки", size=26, weight=ft.FontWeight.W_800, color=TEXT),
            ft.Text("Системные действия и рабочие артефакты проекта.", size=13, color=MUTED),
            ft.ResponsiveRow(
                run_spacing=14,
                spacing=14,
                controls=[
                    card(
                        "Ozon и Wildberries",
                        "Подключение маркетплейсов и настройка ключей API.",
                        [
                            ft.Button("Открыть корень проекта", on_click=lambda _e: reveal_path(ROOT, "корень проекта")),
                            ft.Button("Настроить подключения", on_click=open_marketplace_settings),
                            ft.TextButton("Проверить баланс Ozon", on_click=open_balance_check),
                            ft.Text("Для бизнес-сводки WB нужен доступ токена к категории «Финансы».", size=12, color=MUTED),
                        ],
                        width=320,
                    ),
                    card(
                        "Рабочие файлы",
                        "Быстрый доступ к ключевым папкам и costs.xlsx.",
                        [
                            ft.Button("Открыть costs.xlsx", on_click=lambda _e: reveal_path(ROOT / "costs.xlsx", "costs.xlsx")),
                            ft.TextButton("reports", on_click=lambda _e: reveal_path(ROOT / "reports", "reports")),
                            ft.TextButton("ABC&XYZ reports", on_click=lambda _e: reveal_path(ROOT / "ABC&XYZ reports", "ABC&XYZ reports")),
                        ],
                        width=320,
                    ),
                    card(
                        "Проверка обновлений",
                        "Запускает встроенный модуль автообновления.",
                        [ft.Button("Проверить обновления", bgcolor=ACCENT, color="white", on_click=lambda _e: job("Проверка обновлений", [str(PYTHON), str(SCRIPTS / "_auto_update.py")]))],
                        width=320,
                    ),
                ],
            ),
        ],
        spacing=18,
    )

    file_view = ft.Column(
        [
            ft.Text("Файлы", size=26, weight=ft.FontWeight.W_800, color=TEXT),
            ft.Text("Быстрый доступ ко всем результатам работы приложения.", size=13, color=MUTED),
            files_col,
        ],
        spacing=18,
    )

    views = [dash, ai_view, reports, pricing, actions, supply, finance, settings, file_view]
    active_store = {"value": "ozon"}
    store_sections = {"ozon": 0, "wb": 0}
    section_names = ["Бизнес-сводка", "AI-ассистент", "Отчёты", "Цены и скидки", "Акции", "Поставки", "Финансы", "Настройки", "Файлы"]

    wb_dashboard = build_wb_dashboard(
        months=MONTHS, metric=metric, hero_metric=hero_metric, card=card,
        text_color=TEXT, muted_color=MUTED,
        primary=PRIMARY, accent=ACCENT, surface=semantic_surface("neutral", dark=False),
        page=page, root=ROOT, reveal=reveal_path,
        log=lambda message: ui_queue.put(("log_append", message)),
        theme=lambda control: apply_control_theme(control, current_dark["value"]),
    )

    def wb_section(index):
        if index == 0:
            return wb_dashboard
        if index == 7:
            return settings
        return ft.Column(
            [
                ft.Text("Wildberries", size=13, weight=ft.FontWeight.W_600, color=MUTED),
                ft.Text(section_names[index], size=26, weight=ft.FontWeight.W_800, color=TEXT),
                card(
                    "Раздел Wildberries готовится",
                    "Здесь появятся данные вашего магазина Wildberries. Сейчас можно сохранить API-токен в настройках.",
                    [ft.Button("Настроить Wildberries", on_click=open_marketplace_settings)],
                ),
            ], spacing=18,
        )

    host = ft.Column([views[0]], expand=True, spacing=0, scroll=ft.ScrollMode.AUTO)

    def show_store_section(selected_index: int) -> None:
        store_sections[active_store["value"]] = selected_index
        is_ozon = active_store["value"] == "ozon"
        host.controls = [views[selected_index] if is_ozon else wb_section(selected_index)]
        is_ai_view = is_ozon and selected_index == 1
        header.visible = is_ozon and not is_ai_view
        host.scroll = ft.ScrollMode.HIDDEN if is_ai_view else ft.ScrollMode.AUTO
        host_shell.padding = 0 if is_ai_view else ft.Padding.only(left=8, right=10, top=20, bottom=20)
        host_shell.bgcolor = ai_surface_tokens()["canvas"] if is_ai_view else ("#FBFCFF" if not current_dark["value"] else "#1E2328")
        host_shell.gradient = ft.LinearGradient(
            begin=ft.Alignment(0, -1),
            end=ft.Alignment(0, 1),
            colors=[ai_surface_tokens()["canvas"], ai_surface_tokens()["canvas"]]
            if is_ai_view
            else (["#FBFCFF", "#F3F6FB"] if not current_dark["value"] else ["#1F252A", "#191D21"]),
        )
        apply_control_theme(host.controls[0], current_dark["value"])
        if is_ai_view:
            ai_apply_chat_theme()
        store_name = "Ozon" if is_ozon else "Wildberries"
        page.title = f"{store_name} — OzonReportX"
        set_status("Раздел открыт", f"{store_name} · {section_names[selected_index]}", PRIMARY, busy=False)
        page.update()

    def nav_change(e: ft.ControlEvent) -> None:
        show_store_section(e.control.selected_index)

    theme_button = ft.IconButton(
        icon=ft.Icons.DARK_MODE_OUTLINED,
        selected_icon=ft.Icons.LIGHT_MODE_OUTLINED,
        tooltip="Переключить тему",
        selected=True,
    )

    store_buttons = {}

    def style_store_tabs():
        dark = current_dark["value"]
        rail_brand_card.bgcolor = "#252A2F" if dark else "#F7F7FA"
        rail_brand_card.border = ft.Border.all(1, "#394247" if dark else BORDER)
        for key, button in store_buttons.items():
            selected = active_store["value"] == key
            button.bgcolor = ("#0D6B5B" if dark else "#DCEFE9") if selected else "transparent"
            button.content.controls[0].color = ("#FFFFFF" if dark else PRIMARY) if selected else ("#A1A1A6" if dark else MUTED)
            button.content.controls[1].color = ("#FFFFFF" if dark else TEXT) if selected else ("#A1A1A6" if dark else MUTED)
            button.content.controls[2].visible = selected
            button.content.controls[2].color = "#FFFFFF" if dark else PRIMARY

    def select_store(key):
        if active_store["value"] == key:
            return
        active_store["value"] = key
        rail.selected_index = store_sections[key]
        show_store_section(rail.selected_index)
        style_store_tabs()
        page.update()

    for key, title in (("ozon", "Ozon"), ("wb", "Wildberries")):
        store_buttons[key] = ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.STOREFRONT_OUTLINED, size=20),
                ft.Text(title, size=16, weight=ft.FontWeight.W_700, expand=True),
                ft.Icon(ft.Icons.CHECK_ROUNDED, size=16, color="white", visible=key == "ozon"),
            ], spacing=10),
            padding=ft.Padding.symmetric(horizontal=12, vertical=11),
            border_radius=16,
            on_click=lambda _e, store=key: select_store(store),
            tooltip=f"Открыть магазин {title}",
        )
    rail_brand_card = ft.Container(
        width=240,
        padding=7,
        border_radius=22,
        bgcolor="#F7F7FA",
        border=ft.Border.all(1, BORDER),
        content=ft.Column(list(store_buttons.values()), spacing=4),
    )

    rail = ft.NavigationRail(
        selected_index=0,
        extended=True,
        min_width=92,
        min_extended_width=228,
        group_alignment=-0.88,
        label_type=ft.NavigationRailLabelType.ALL,
        bgcolor=SURFACE,
        indicator_color=PRIMARY_SOFT,
        on_change=nav_change,
        leading=ft.Container(
            content=rail_brand_card,
            padding=ft.Padding.only(bottom=18),
        ),
        destinations=[
            ft.NavigationRailDestination(icon=ft.Icons.SPACE_DASHBOARD_OUTLINED, selected_icon=ft.Icons.SPACE_DASHBOARD, label="Бизнес-сводка"),
            ft.NavigationRailDestination(icon=ft.Icons.AUTO_AWESOME_OUTLINED, selected_icon=ft.Icons.AUTO_AWESOME, label="AI-ассистент"),
            ft.NavigationRailDestination(icon=ft.Icons.INSERT_CHART_OUTLINED, selected_icon=ft.Icons.INSERT_CHART, label="Отчёты"),
            ft.NavigationRailDestination(icon=ft.Icons.LOCAL_OFFER_OUTLINED, selected_icon=ft.Icons.LOCAL_OFFER, label="Цены и скидки"),
            ft.NavigationRailDestination(icon=ft.Icons.LOCAL_ACTIVITY_OUTLINED, selected_icon=ft.Icons.LOCAL_ACTIVITY, label="Акции"),
            ft.NavigationRailDestination(icon=ft.Icons.INVENTORY_2_OUTLINED, selected_icon=ft.Icons.INVENTORY_2, label="Supply"),
            ft.NavigationRailDestination(icon=ft.Icons.ACCOUNT_BALANCE_WALLET_OUTLINED, selected_icon=ft.Icons.ACCOUNT_BALANCE_WALLET, label="Finance"),
            ft.NavigationRailDestination(icon=ft.Icons.SETTINGS_OUTLINED, selected_icon=ft.Icons.SETTINGS, label="Settings"),
            ft.NavigationRailDestination(icon=ft.Icons.FOLDER_OPEN_OUTLINED, selected_icon=ft.Icons.FOLDER_OPEN, label="Files"),
        ],
    )

    rail_shell = ft.Container(
        content=rail,
        width=278,
        height=780,
        padding=18,
        bgcolor=SURFACE,
        border_radius=34,
        border=ft.Border.all(1, BORDER),
        shadow=ft.BoxShadow(blur_radius=34, spread_radius=0, color="#12000000", offset=ft.Offset(0, 14)),
    )
    host_shell = ft.Container(
        content=host,
        expand=True,
        padding=ft.Padding.only(left=8, right=10, top=20, bottom=20),
        bgcolor="#FBFCFF",
        gradient=ft.LinearGradient(
            begin=ft.Alignment(0, -1),
            end=ft.Alignment(0, 1),
            colors=["#1F252A", "#191D21"],
        ),
        border_radius=30,
        border=ft.Border.all(1, "#ECECF1"),
        shadow=ft.BoxShadow(blur_radius=34, spread_radius=0, color="#12000000", offset=ft.Offset(0, 14)),
    )
    status_shell = ft.Container(
        content=ft.Column(
            [
                ft.Row(
                    [
                        status_indicator,
                        state,
                        ft.Container(expand=True),
                        spin,
                    ],
                    spacing=10,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                ),
                status_detail,
                ft.Row(
                    [cancel_button],
                    alignment=ft.MainAxisAlignment.END,
                ),
            ],
            spacing=6,
        ),
        padding=ft.Padding.symmetric(horizontal=11, vertical=8),
        bgcolor=SURFACE,
        border_radius=16,
        border=ft.Border.all(1, BORDER),
        shadow=ft.BoxShadow(blur_radius=12, spread_radius=0, color="#0D000000", offset=ft.Offset(0, 4)),
        width=264,
    )
    cancel_button.on_click = cancel_current_action
    def apply_theme(dark: bool) -> None:
        current_dark["value"] = dark
        page.theme_mode = ft.ThemeMode.DARK if dark else ft.ThemeMode.LIGHT
        page.bgcolor = "#111214" if dark else BG
        rail.bgcolor = "#20262B" if dark else SURFACE
        rail.indicator_color = "#2F4A45" if dark else PRIMARY_SOFT
        rail_shell.bgcolor = "#20262B" if dark else SURFACE
        rail_shell.border = ft.Border.all(1, "#394247" if dark else "#E1E5EE")
        rail_shell.shadow = ft.BoxShadow(
            blur_radius=28,
            spread_radius=0,
            color="#22000000" if dark else "#14000000",
            offset=ft.Offset(0, 12),
        )
        host_shell.bgcolor = ai_surface_tokens()["canvas"] if active_store["value"] == "ozon" and rail.selected_index == 1 else ("#1E2328" if dark else "#FBFCFF")
        host_shell.gradient = ft.LinearGradient(
            begin=ft.Alignment(0, -1),
            end=ft.Alignment(0, 1),
            colors=[ai_surface_tokens()["canvas"], ai_surface_tokens()["canvas"]]
            if active_store["value"] == "ozon" and rail.selected_index == 1
            else (["#1F252A", "#191D21"] if dark else ["#FBFCFF", "#F3F6FB"]),
        )
        host_shell.border = ft.Border.all(1, "#2C2C2E" if dark else "#ECECF1")
        host_shell.shadow = ft.BoxShadow(
            blur_radius=30,
            spread_radius=0,
            color="#18000000" if dark else "#12000000",
            offset=ft.Offset(0, 12),
        )
        status_shell.bgcolor = "#1F252A" if dark else "#FBFCFF"
        status_shell.border = ft.Border.all(1, "#353D42" if dark else "#E1E5EE")
        status_shell.shadow = ft.BoxShadow(
            blur_radius=12,
            spread_radius=0,
            color="#14000000" if dark else "#0D000000",
            offset=ft.Offset(0, 4),
        )
        status_detail.color = "#A1A1A6" if dark else "#6E6E73"
        top_tools_shell.bgcolor = "#D020252A" if dark else "#FAFBFE"
        top_tools_shell.border = ft.Border.all(1, "#3A4348" if dark else "#E1E5EE")
        splash_overlay.bgcolor = "#CC111214" if dark else "#D8F5F5F7"
        splash_card.bgcolor = "#20262B" if dark else "#FFFFFF"
        splash_card.border = ft.Border.all(1, "#354148" if dark else BORDER)
        splash_card.shadow = ft.BoxShadow(
            blur_radius=40,
            spread_radius=0,
            color="#28000000" if dark else "#16000000",
            offset=ft.Offset(0, 16),
        )
        splash_title.color = "#F5F5F7" if dark else "#1D1D1F"
        splash_tagline.color = "#A1A1A6" if dark else "#6E6E73"
        splash_note.color = "#A1A1A6" if dark else "#6E6E73"
        theme_button.selected = dark
        theme_button.icon_color = "#B6C0BC" if dark else "#4F5B66"
        theme_button.selected_icon_color = "#B6C0BC" if dark else "#4F5B66"
        menu_bar_inner.bgcolor = "#23282D" if dark else "#FFFFFF"
        menu_bar_inner.border = ft.Border.all(1, "#384046" if dark else "#E5E7EF")
        apply_control_theme(rail_brand_card, dark)
        apply_control_theme(rail_shell, dark)
        _render_log_lines(log_state["lines"])
        apply_prompt_dialog_theme()
        apply_discount_dialog_theme()
        refresh_pricing_costs_preview()
        for root_control in page.controls:
            apply_control_theme(root_control, dark)
        ai_apply_chat_theme()
        style_store_tabs()
        set_status(state.value, status_detail.value, state.color if isinstance(state.color, str) else PRIMARY, busy=spin.visible)
        page.update()

    theme_button.on_click = lambda _e: apply_theme(not theme_button.selected)

    header = ft.Container(
        content=ft.Row(
            [
                ft.Container(expand=True),
            ],
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        ),
        padding=ft.Padding.only(bottom=18),
    )
    menu_bar_inner = ft.Container(
        content=ft.Row(
            [
                ai_tools_button,
            ],
            spacing=4,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
            tight=True,
        ),
        padding=ft.Padding.symmetric(horizontal=10, vertical=6),
        border_radius=14,
        bgcolor="#23282D",
        border=ft.Border.all(1, "#384046"),
    )
    top_tools_shell = ft.Container(
        content=ft.Row(
            [
                menu_bar_inner,
                theme_button,
            ],
            spacing=10,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        ),
        padding=ft.Padding.symmetric(horizontal=8, vertical=8),
        bgcolor="#D020252A" if current_dark["value"] else "#FAFBFE",
        border_radius=18,
        border=ft.Border.all(1, "#3A4348" if current_dark["value"] else "#E1E5EE"),
        shadow=ft.BoxShadow(blur_radius=16, spread_radius=0, color="#12000000", offset=ft.Offset(0, 6)),
    )
    splash_progress = ft.ProgressRing(width=26, height=26, stroke_width=2.6, color=PRIMARY)
    splash_title = ft.Text("OzonReportX", size=22, weight=ft.FontWeight.W_800, color="#F5F5F7")
    splash_tagline = ft.Text("Контроль. Аналитика. Автоматизация.", size=12, color="#A1A1A6")
    splash_note = ft.Text("Подготавливаем рабочее пространство", size=12, color="#A1A1A6")
    splash_card = ft.Container(
        width=360,
        padding=28,
        border_radius=32,
        bgcolor="#20262B",
        border=ft.Border.all(1, "#354148"),
        shadow=ft.BoxShadow(blur_radius=40, spread_radius=0, color="#28000000", offset=ft.Offset(0, 16)),
        content=ft.Column(
            [
                ft.Row(
                    [
                        ft.Container(
                            content=ft.Image(
                                src=str((ROOT / "img" / "ozonreportx-mark.svg").resolve()),
                                width=42,
                                height=42,
                            ),
                            width=64,
                            height=64,
                            bgcolor="#0D6B5B",
                            border_radius=22,
                            alignment=ft.Alignment(0, 0),
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                ),
                ft.Column(
                    [
                        splash_title,
                        splash_tagline,
                    ],
                    spacing=4,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                ),
                ft.Container(height=1, bgcolor="#354148", border_radius=10),
                ft.Row(
                    [
                        splash_progress,
                        splash_note,
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                    spacing=12,
                ),
            ],
            spacing=18,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        ),
    )
    splash_overlay = ft.Container(
        expand=True,
        bgcolor="#D8F5F5F7",
        alignment=ft.Alignment(0, 0),
        content=splash_card,
        visible=True,
    )
    main_layout = ft.Column(
        [
            ft.Container(height=76),
            header,
            ft.Row(
                [
                    ft.Container(width=280),
                    host_shell,
                ],
                expand=True,
                vertical_alignment=ft.CrossAxisAlignment.START,
            ),
        ],
        expand=True,
    )
    page.add(main_layout)
    page.overlay.append(
        ft.Container(
            content=status_shell,
            left=18,
            bottom=24,
            animate_position=200,
        )
    )
    page.overlay.append(
        ft.Container(
            content=rail_shell,
            left=24,
            top=98,
        )
    )
    page.overlay.append(
        ft.Container(
            content=top_tools_shell,
            left=24,
            top=18,
            animate_position=200,
        )
    )
    page.overlay.append(splash_overlay)
    status_ui_ready["value"] = True
    set_status("Запуск", "Подготавливаем интерфейс и данные.", ACCENT, busy=True)
    apply_theme(True)

    async def initialize_app() -> None:
        page.update()
        await asyncio.sleep(0.12)
        refresh_dashboard()
        refresh_files()
        refresh_pricing_costs_preview()
        splash_overlay.visible = False
        set_status("Готово", "Интерфейс готов к работе.", PRIMARY, busy=False)
        page.update()
        if needs_setup(ROOT):
            open_marketplace_settings()

    page.run_task(ui_event_loop)
    page.run_task(initialize_app)


if __name__ == "__main__":
    ft.run(main)
