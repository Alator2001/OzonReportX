# -*- coding: utf-8 -*-
"""
Модуль для чата с ИИ на тему бизнеса Ozon.
"""

import os
import sys
import re
import json
import time
import requests
import pandas as pd
from pathlib import Path
from typing import Optional, Tuple, Dict, Any, List
from datetime import datetime, timedelta
from dotenv import load_dotenv
from openpyxl import load_workbook

# Загружаем переменные окружения
load_dotenv()
GROQ_API_KEY = os.getenv('GROQ_API_KEY')

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
except ImportError:
    # Если не удалось импортировать, создаём заглушки
    REPORTS_DIR_NAME = "reports"
    ORDER_SHEET = "Заказы"
    MONTHS_RU = [
        "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
    ]
    
    def get_report_path(repo_root: Path, year: int, month: int) -> Path:
        name = f"{MONTHS_RU[month - 1]} {year}.xlsx"
        return repo_root / REPORTS_DIR_NAME / name

GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"
# Рекомендуемая модель для чата: groq/compound-mini (70K TPM, без лимита TPD)
# llama-3.3-70b-versatile имеет лимит 12K TPM — при больших запросах даёт 429
DEFAULT_MODEL = "groq/compound-mini"
# Модель можно переопределить в .env: GROQ_MODEL=groq/compound-mini
GROQ_MODEL = (os.getenv("GROQ_MODEL") or "").strip() or DEFAULT_MODEL

# Лимиты для моделей Groq API (на основе данных из консоли)
# Обновляйте эти значения, если они изменятся в вашем аккаунте
MODEL_LIMITS = {
    "llama-3.3-70b-versatile": {
        "rpm": 30,  # Requests per Minute
        "rpd": 1000,  # Requests per Day
        "tpm": 12000,  # Tokens per Minute
        "tpd": 100000,  # Tokens per Day
    },
    "groq/compound": {
        "rpm": 30,
        "rpd": 250,
        "tpm": 70000,
        "tpd": None,  # No limit
    },
    "groq/compound-mini": {
        "rpm": 30,
        "rpd": 250,
        "tpm": 70000,
        "tpd": None,  # No limit
    },
    "llama-3.1-8b-instant": {
        "rpm": 30,
        "rpd": 14400,
        "tpm": 6000,
        "tpd": 500000,
    },
}

# Системный промпт - только инструкции и возможности, БЕЗ данных
SYSTEM_PROMPT = r"""
Ты — AI-консультант по бизнесу на маркетплейсе Ozon и работе программы OzonReportX.

# 0) Главная цель
Давай практичные рекомендации и анализ по бизнесу Ozon, используя данные, которые программа может предоставить через инструменты (TOOLS).

# 0.1) ПРИМЕР ПРАВИЛЬНОГО ПОВЕДЕНИЯ
Когда пользователь пишет: "Проанализируй последние 3 месяца"
Ты ДОЛЖЕН вернуть ТОЛЬКО этот JSON (без текста до/после):
{"type": "DATA_REQUEST", "needs": [{"tool": "list_reports", "args": {}}, {"tool": "get_month_summary", "args": {"period": "2025-12"}}, {"tool": "get_month_summary", "args": {"period": "2025-11"}}, {"tool": "get_month_summary", "args": {"period": "2025-10"}}], "reason": "Анализ за последние 3 месяца"}

НЕ возвращай FINAL_ANSWER с текстом "нужны данные" или "запросите данные" - это ОШИБКА!

# 1) Критически важные правила (обязательно)
1) НЕЛЬЗЯ выдумывать цифры, периоды, значения метрик, результаты отчётов и содержимое файлов.
   - Если данных нет в сообщениях, запроси их через DATA_REQUEST.
2) Всегда указывай период, к которому относятся любые цифры и выводы (например: "2025-12" или "Декабрь 2025").
3) Если в полученных данных нет нужного поля или есть ошибка — так и скажи в FINAL_ANSWER: "данные недоступны" / "инструмент вернул ошибку".
4) Экономь токены:
   - Запрашивай минимальный набор инструментов.
   - Не запрашивай большие списки без необходимости.
5) После того как пользователь/система прислали блок "✅ ДАННЫЕ УСПЕШНО ПОЛУЧЕНЫ" (или любой блок данных от инструментов),
   тебе ЗАПРЕЩЕНО возвращать DATA_REQUEST. Ты обязан вернуть FINAL_ANSWER на основе полученных данных.
6) Ты НЕ должен объяснять внутреннюю кухню (про инструменты, токены, архитектуру) пользователю, если он не спрашивает напрямую.

# 1.1) СТРОЖАЙШЕЕ ПРАВИЛО - НИКОГДА НЕ ПРОСИ ПОЛЬЗОВАТЕЛЯ ЗАПРОСИТЬ ДАННЫЕ
⚠️ КРИТИЧЕСКИ ВАЖНО: Если пользователь просит "проанализируй", "сравни", "покажи", "дай", "сколько" + упоминание периода/месяца/артикула:
   - Ты ДОЛЖЕН вернуть DATA_REQUEST с нужными инструментами
   - ЗАПРЕЩЕНО возвращать FINAL_ANSWER с текстом типа "Пожалуйста, запросите данные" или "нужны данные из отчётов"
   - ЗАПРЕЩЕНО объяснять пользователю, что нужны данные - ты САМ их запрашиваешь через DATA_REQUEST
   
Примеры ПРАВИЛЬНОГО поведения:
   - Пользователь: "Проанализируй последние 3 месяца"
     → Ты возвращаешь: {"type": "DATA_REQUEST", "needs": [{"tool": "list_reports", "args": {}}, {"tool": "get_month_summary", "args": {"period": "2025-12"}}, {"tool": "get_month_summary", "args": {"period": "2025-11"}}, {"tool": "get_month_summary", "args": {"period": "2025-10"}}], "reason": "Анализ за последние 3 месяца"}
   
   - Пользователь: "Сравни прибыль за декабрь и ноябрь"
     → Ты возвращаешь: {"type": "DATA_REQUEST", "needs": [{"tool": "get_month_summary", "args": {"period": "2025-12"}}, {"tool": "get_month_summary", "args": {"period": "2025-11"}}], "reason": "Сравнение прибыли за два месяца"}

Примеры НЕПРАВИЛЬНОГО поведения (ЗАПРЕЩЕНО):
   - Пользователь: "Проанализируй последние 3 месяца"
     → ❌ НЕПРАВИЛЬНО: {"type": "FINAL_ANSWER", "answer": "Для анализа нужны данные. Пожалуйста, запросите месячные сводки."}
     → ❌ НЕПРАВИЛЬНО: {"type": "FINAL_ANSWER", "answer": "Сейчас доступны только списки отчётов, без метрик. Запросите get_month_summary."}

# 2) Разрешённые форматы ответа
Ты ОБЯЗАН возвращать ТОЛЬКО один валидный JSON-объект (без Markdown, без ``` и без текста снаружи JSON).

Разрешено только 2 типа:

A) DATA_REQUEST — если для ответа обязательно нужны данные:
{
  "type": "DATA_REQUEST",
  "needs": [
    {"tool": "<tool_name>", "args": {...}}
  ],
  "reason": "<коротко зачем это нужно>"
}

B) FINAL_ANSWER — если данных достаточно или вопрос общий:
{
  "type": "FINAL_ANSWER",
  "answer": "<ответ пользователю на русском>"
}

Запрещено:
- добавлять любые поля кроме type/needs/reason/answer
- писать текст до/после JSON
- возвращать массив вместо объекта

# 3) Как принимать решение: DATA_REQUEST или FINAL_ANSWER
🚨 КРИТИЧЕСКИ ВАЖНО: Ты НИКОГДА не должен просить пользователя запросить данные. Если нужны данные - ты САМ запрашиваешь их через DATA_REQUEST.

Возвращай FINAL_ANSWER ТОЛЬКО если:
- вопрос общий (что умеет программа, как считать маржу, как интерпретировать метрики и т.п.)
- можно ответить без конкретных чисел И без данных из отчётов
- пользователь не указал период/артикул и можно попросить уточнение в самом FINAL_ANSWER (без инструментов)

🚨 ОБЯЗАТЕЛЬНО возвращай DATA_REQUEST, если пользователь просит:
- "проанализируй", "сравни", "покажи", "дай", "сколько", "какая" + упоминание периода/месяца/артикула
- конкретные цифры (выручка/прибыль/маржа/заказы/комиссии) за период
- анализ конкретного артикула / сравнение периодов
- топы по прибыли/заказам
- себестоимость/минимальную/желательную цену по артикулам
- ABC&XYZ категории, списки категорий, сводки
- ЛЮБОЙ запрос, который требует данных из отчётов или costs.xlsx

🚨 ЗАПРЕЩЕНО (это КРИТИЧЕСКАЯ ОШИБКА):
- Просить пользователя запросить данные (например: "Пожалуйста, запросите месячные сводки")
- Объяснять, что нужны данные, вместо того чтобы их запросить
- Возвращать FINAL_ANSWER с просьбой к пользователю использовать инструменты
- Возвращать FINAL_ANSWER с текстом типа "нужны данные из отчётов" или "доступны только списки отчётов"

🚨 ОБЯЗАТЕЛЬНО: Если пользователь просит "последние N месяцев", "проанализируй N месяцев", "сравни последние N месяцев":
- СНАЧАЛА запроси list_reports {}, чтобы узнать доступные периоды
- ЗАТЕМ в ОДНОМ DATA_REQUEST запроси get_month_summary для КАЖДОГО из последних N месяцев
- НЕ делай отдельные запросы для каждого месяца - запроси все сразу в одном DATA_REQUEST
- Пример для "последние 3 месяца": {"type": "DATA_REQUEST", "needs": [{"tool": "list_reports", "args": {}}, {"tool": "get_month_summary", "args": {"period": "2025-12"}}, {"tool": "get_month_summary", "args": {"period": "2025-11"}}, {"tool": "get_month_summary", "args": {"period": "2025-10"}}], "reason": "Анализ за последние 3 месяца"}

Если период/артикул не указан и без него нельзя выбрать данные:
- НЕ вызывай инструменты
- верни FINAL_ANSWER с просьбой уточнить период/артикул (коротко, 1–2 вопроса)

# 4) Контекст программы OzonReportX (кратко)
Программа работает с:
- costs.xlsx: себестоимость и расчётные цены (минимальная/желательная), и др. поля
- месячные отчёты продаж Excel (лист "Заказы"): сводка метрик + детальные данные заказов
- ABC&XYZ отчёты: прибыльность (A/B/C) и стабильность спроса (X/Y/Z)

Формулы:
- Прибыль по заказу = "Сумма начисления" - "Себестоимость"
- Рентабельность = (прибыль / выручка) * 100%
- Рекомендованная цена = Себестоимость / (1 - (Комиссия% + Логистика%) - маржа)

# 5) Доступные инструменты (TOOLS)
Инструменты вызываются ТОЛЬКО через DATA_REQUEST.

Метаданные:
- list_reports {}
- list_abcxyz_reports {}

Месячные отчёты:
- get_month_summary {"period": "2025-12"}
- get_top_profit {"period": "2025-12", "n": 10}
- get_top_orders {"period": "2025-12", "n": 10}
- get_artikul_stats {"period": "2025-12", "artikul": "12345"}
- search_artikul {"query": "доска круглая"}

Costs:
- get_costs_columns {}
- get_artikul_cost {"artikul": "12345"}
- get_costs_bulk {"artikuls": ["12345","67890"]}

ABC&XYZ:
- get_abcxyz_summary {"period": "Ноябрь 2025-Декабрь 2025"}
- get_artikul_abcxyz {"period": "...", "artikul": "12345"}
- get_category_list {"period": "...", "category": "AX"}

# 6) Нормализация периода
Если пользователь пишет месяц словами ("Декабрь 2025"), можно использовать его как period.
Если пользователь пишет "2025-12" — используй этот формат.

# 7) Стиль ответа в FINAL_ANSWER
- Коротко и по делу на русском
- С эмодзи для акцентов (📊 📈 💰 ⚠️ ✅ ❌ 🔍 💡)
- Если есть ошибки в данных — объясни простыми словами
- Предлагай конкретные шаги (что сделать на Ozon / в программе)

"""


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
        df = pd.read_excel(costs_path)
        
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


def load_monthly_report_detailed_data(report_path: Path) -> Optional[Dict[str, Any]]:
    """
    Загружает детальные данные из месячного отчёта (лист "Заказы").
    Агрегирует данные по артикулам для экономии токенов.
    
    Returns:
        Словарь с агрегированными данными или None в случае ошибки
    """
    if not report_path.exists():
        return None
    
    try:
        df = pd.read_excel(report_path, sheet_name=ORDER_SHEET)
        
        if df.empty:
            return None
        
        report_name = report_path.stem
        detailed_data = {
            "Период отчёта": report_name,
            "Всего заказов": len(df),
            "Колонки": list(df.columns),
        }
        
        # Агрегируем данные по артикулам
        if "Артикул" in df.columns:
            artikul_stats = []
            
            # Группируем по артикулам
            for artikul in df["Артикул"].dropna().unique():
                artikul_df = df[df["Артикул"] == artikul]
                
                stats = {
                    "Артикул": str(artikul),
                    "Количество заказов": len(artikul_df),
                }
                
                # Суммируем числовые колонки
                numeric_cols = ["Количество шт.", "Цена продажи", "Комиссия за продажу Ozon", 
                               "Логистика (Включает операционные ошибки продавца)", 
                               "Сумма начисления", "Себестоимость", "Прибыль"]
                
                for col in numeric_cols:
                    if col in artikul_df.columns:
                        total = artikul_df[col].sum()
                        if pd.notna(total):
                            stats[f"Сумма {col}"] = float(total)
                            stats[f"Средняя {col}"] = float(artikul_df[col].mean())
                
                # Статистика по статусам
                if "Статус" in artikul_df.columns:
                    status_counts = artikul_df["Статус"].value_counts().to_dict()
                    stats["Распределение по статусам"] = {str(k): int(v) for k, v in status_counts.items()}
                
                # Статистика по схемам
                if "Схема" in artikul_df.columns:
                    schema_counts = artikul_df["Схема"].value_counts().to_dict()
                    stats["Распределение по схемам"] = {str(k): int(v) for k, v in schema_counts.items()}
                
                artikul_stats.append(stats)
            
            detailed_data["Статистика по артикулам"] = artikul_stats
        
        # Общая статистика по статусам
        if "Статус" in df.columns:
            status_counts = df["Статус"].value_counts().to_dict()
            detailed_data["Общее распределение по статусам"] = {str(k): int(v) for k, v in status_counts.items()}
        
        # Общая статистика по схемам
        if "Схема" in df.columns:
            schema_counts = df["Схема"].value_counts().to_dict()
            detailed_data["Общее распределение по схемам"] = {str(k): int(v) for k, v in schema_counts.items()}
        
        # Топ-10 артикулов по прибыли
        if "Прибыль" in df.columns and "Артикул" in df.columns:
            top_profit = df.groupby("Артикул")["Прибыль"].sum().nlargest(10)
            detailed_data["Топ-10 артикулов по прибыли"] = {
                str(artikul): float(profit) for artikul, profit in top_profit.items()
            }
        
        # Топ-10 артикулов по количеству заказов
        if "Артикул" in df.columns:
            top_orders = df["Артикул"].value_counts().head(10)
            detailed_data["Топ-10 артикулов по количеству заказов"] = {
                str(artikul): int(count) for artikul, count in top_orders.items()
            }
        
        return detailed_data
        
    except Exception as e:
        return None


def load_report_summary(report_path: Path) -> Optional[Dict[str, Any]]:
    """
    Загружает сводку из месячного отчёта.
    Читает итоговые показатели из колонок P и Q листа "Заказы".
    
    Returns:
        Словарь с метриками отчёта или None в случае ошибки
    """
    if not report_path.exists():
        return None
    
    try:
        wb = load_workbook(report_path, data_only=True)
        if ORDER_SHEET not in wb.sheetnames:
            return None
        
        ws = wb[ORDER_SHEET]
        
        # Читаем итоговые показатели из колонок P и Q
        # Структура: P1-Q1: Общая выручка, P2-Q2: Чистая прибыль, и т.д.
        summary = {}
        
        # Маппинг строк к названиям метрик
        metrics_map = {
            1: "Общая выручка",
            2: "Чистая прибыль",
            3: "Итоговая себестоимость",
            4: "Рентабельность по чистой прибыли (Net Profit Margin) %",
            5: "COGS (валовая прибыль)",
            6: "Gross Profit Margin Рентабельность по валовой прибыли %",
            7: "Операционные расходы",
            8: "Продвижение Ozon",
            9: "Звёздные товары",
            10: "Внешний маркетинг",
            11: "Средний чек",
            12: "Общее количество заказов",
            13: "Количество отменённых заказов",
            14: "Количество доставленных заказов",
            15: "Комиссии Ozon %",
            16: "Логистика %",
        }
        
        for row_num, metric_name in metrics_map.items():
            value = ws[f"Q{row_num}"].value
            if value is not None:
                try:
                    # Пробуем преобразовать в число
                    if isinstance(value, (int, float)):
                        summary[metric_name] = float(value)
                    elif isinstance(value, str):
                        # Пробуем распарсить строку
                        cleaned = value.replace(",", ".").replace(" ", "")
                        summary[metric_name] = float(cleaned)
                    else:
                        summary[metric_name] = value
                except (ValueError, TypeError):
                    summary[metric_name] = value
        
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
        if file_path.is_file():
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
        if file_path.is_file():
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
    if not report_path.exists():
        return None
    
    try:
        wb = load_workbook(report_path, data_only=True)
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
    
    def _get_costs_df(self) -> Optional[pd.DataFrame]:
        """Загружает costs.xlsx с кешированием."""
        if self._costs_df is None:
            costs_path = self.repo_root / "costs.xlsx"
            if costs_path.exists():
                try:
                    self._costs_df = pd.read_excel(costs_path)
                except Exception:
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
    
    # Пробуем распарсить весь текст как JSON (если это чистый JSON)
    try:
        parsed = json.loads(response_text)
        if isinstance(parsed, dict) and "type" in parsed:
            return parsed
    except json.JSONDecodeError:
        pass
    
    # Пробуем найти JSON в code blocks (```json ... ```)
    json_match = re.search(r'```json\s*(\{.*?\})\s*```', response_text, re.DOTALL)
    if json_match:
        try:
            parsed = json.loads(json_match.group(1))
            if isinstance(parsed, dict) and "type" in parsed:
                return parsed
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
                            if isinstance(parsed, dict) and "type" in parsed:
                                return parsed
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


def format_tool_results(results: Dict[str, Any]) -> str:
    """Форматирует результаты выполнения инструментов для передачи ИИ."""
    lines = [
        "=" * 70,
        "✅ ДАННЫЕ УСПЕШНО ПОЛУЧЕНЫ ИЗ ПРОГРАММЫ",
        "=" * 70,
        "",
        "Ты запросил данные через инструменты, и программа их предоставила.",
        "Теперь используй эти данные для анализа и дай финальный ответ пользователю.",
        "",
        "⚠️ КРИТИЧЕСКИ ВАЖНО:",
        "- Верни ТОЛЬКО FINAL_ANSWER (не DATA_REQUEST)",
        "- Проанализируй полученные данные",
        "- Дай конкретный ответ на вопрос пользователя",
        "",
        "=" * 70,
        "РЕЗУЛЬТАТЫ ВЫПОЛНЕНИЯ ИНСТРУМЕНТОВ:",
        "=" * 70,
        ""
    ]
    
    for tool_name, result in results.items():
        lines.append(f"📊 Инструмент: {tool_name}")
        
        # Если результат - список (множественные вызовы одного инструмента)
        if isinstance(result, list):
            for idx, item in enumerate(result):
                if idx > 0:
                    lines.append("")  # Разделитель между результатами
                if isinstance(item, dict) and "error" in item:
                    lines.append(f"   ❌ Ошибка (вызов {idx + 1}): {item['error']}")
                else:
                    result_str = json.dumps(item, ensure_ascii=False, indent=2)
                    lines.append(f"   Результат {idx + 1}:")
                    lines.append(f"   {result_str}")
        elif isinstance(result, dict) and "error" in result:
            lines.append(f"   ❌ Ошибка: {result['error']}")
        else:
            # Форматируем результат в читаемый вид
            result_str = json.dumps(result, ensure_ascii=False, indent=2)
            lines.append(f"   {result_str}")
        lines.append("")
    
    lines.append("=" * 70)
    lines.append("Используй эти данные для финального ответа. Верни JSON: {\"type\": \"FINAL_ANSWER\", \"answer\": \"...\"}")
    lines.append("=" * 70)
    
    return "\n".join(lines)


def chat_with_ai(user_message: str, conversation_history: list = None, data_block: Optional[str] = None, context_state: Optional[str] = None, temperature: float = 0.7, max_tokens: int = 2000) -> Tuple[Optional[str], Optional[Dict[str, Any]], Optional[Dict[str, Any]]]:
    """
    Отправляет сообщение в чат с ИИ и получает ответ.
    
    Args:
        user_message: Сообщение пользователя
        conversation_history: История разговора (список сообщений, последние 3-6 сообщений)
        data_block: Блок данных для передачи ИИ (результаты инструментов)
        context_state: Состояние контекста (selected_period, selected_artikul и т.д.)
        temperature: Температура для генерации (0.2 для режима B, 0.7 для режима A)
        max_tokens: Максимальное количество токенов (400-700 для режима A, 2000 для режима B)
    
    Returns:
        Кортеж (ответ ИИ, информация об использовании токенов, информация о лимитах) или (None, None, None) в случае ошибки
    """
    if not GROQ_API_KEY:
        print("⚠️ GROQ_API_KEY не настроен в .env файле.")
        print("   Добавьте GROQ_API_KEY=ваш_ключ в файл .env")
        return None, None, None
    
    # Формируем сообщения для API
    system_prompt = SYSTEM_PROMPT
    
    # Добавляем состояние контекста, если есть
    if context_state:
        system_prompt += f"\n\n## ТЕКУЩЕЕ СОСТОЯНИЕ:\n{context_state}\n"
    
    messages = [{"role": "system", "content": system_prompt}]
    
    # Добавляем историю разговора (только последние 3-6 сообщений для экономии токенов)
    if conversation_history:
        # Берём последние 6 сообщений (3 пары вопрос-ответ)
        recent_history = conversation_history[-6:] if len(conversation_history) > 6 else conversation_history
        messages.extend(recent_history)
    
    # Добавляем блок данных, если есть
    if data_block:
        # Когда передаём данные, явно указываем, что нужен финальный ответ
        messages.append({
            "role": "user", 
            "content": f"""⚠️ ВАЖНО: Ты получил запрошенные данные. Теперь ты ДОЛЖЕН дать FINAL_ANSWER.

ДАННЫЕ:
{data_block}

ВОПРОС ПОЛЬЗОВАТЕЛЯ: {user_message}

ИНСТРУКЦИЯ: 
- Используй полученные данные для анализа
- Верни ТОЛЬКО JSON с type="FINAL_ANSWER" и полным ответом пользователю
- НЕ возвращай DATA_REQUEST - данные уже получены
- Если данных недостаточно - скажи об этом в FINAL_ANSWER"""
        })
    else:
        # Добавляем текущее сообщение пользователя
        messages.append({"role": "user", "content": user_message})
    
    max_retries = 2  # при 429 повторяем до 2 раз (всего 3 попытки)
    last_error = None
    
    try:
        for attempt in range(max_retries + 1):
            try:
                response = requests.post(
                    GROQ_API_URL,
                    headers={
                        "Authorization": f"Bearer {GROQ_API_KEY}",
                        "Content-Type": "application/json"
                    },
                    json={
                        "model": GROQ_MODEL,
                        "messages": messages,
                        "temperature": temperature,
                        "max_tokens": max_tokens
                    },
                    timeout=60
                )
                
                response.raise_for_status()
                data = response.json()
                
                ai_response = data.get("choices", [{}])[0].get("message", {}).get("content", "")
                usage_info = data.get("usage", {})
                
                # Извлекаем информацию о лимитах из заголовков ответа
                rate_limit_info = {}
                headers = response.headers
                
                rate_limit_info["limit"] = (
                    headers.get("x-ratelimit-limit-tokens") or 
                    headers.get("x-ratelimit-limit") or
                    headers.get("ratelimit-limit-tokens") or
                    headers.get("x-ratelimit-limit-tpm") or
                    headers.get("x-ratelimit-limit-tpd") or
                    None
                )
                rate_limit_info["remaining"] = (
                    headers.get("x-ratelimit-remaining-tokens") or 
                    headers.get("x-ratelimit-remaining") or
                    headers.get("ratelimit-remaining-tokens") or
                    headers.get("x-ratelimit-remaining-tpm") or
                    headers.get("x-ratelimit-remaining-tpd") or
                    None
                )
                rate_limit_info["reset"] = (
                    headers.get("x-ratelimit-reset-tokens") or 
                    headers.get("x-ratelimit-reset") or
                    headers.get("ratelimit-reset-tokens") or
                    headers.get("x-ratelimit-reset-tpm") or
                    headers.get("x-ratelimit-reset-tpd") or
                    None
                )
                
                if rate_limit_info["limit"] is None:
                    rate_limit_info["limit"] = headers.get("x-ratelimit-limit-requests") or headers.get("x-ratelimit-limit-rpm")
                    rate_limit_info["remaining"] = headers.get("x-ratelimit-remaining-requests") or headers.get("x-ratelimit-remaining-rpm")
                    rate_limit_info["reset"] = headers.get("x-ratelimit-reset-requests") or headers.get("x-ratelimit-reset-rpm")
                
                rate_limit_info["model"] = GROQ_MODEL
                
                if ai_response:
                    return ai_response, usage_info, rate_limit_info
                else:
                    print("⚠️ ИИ не вернул ответ.")
                    return None, None, None
                    
            except requests.exceptions.RequestException as e:
                last_error = e
                if hasattr(e, 'response') and e.response is not None and e.response.status_code == 429:
                    try:
                        error_data = e.response.json()
                        error_msg = error_data.get('error', {}).get('message', '')
                        wait_time_match = re.search(r'(\d+\.?\d*)\s*s', error_msg)
                        wait_sec = float(wait_time_match.group(1)) + 2.0 if wait_time_match else 10.0
                        wait_sec = min(wait_sec, 60.0)
                        if attempt < max_retries:
                            print(f"\n   ⚠ 429, повтор через {wait_sec:.0f} с…", end="\r")
                            time.sleep(wait_sec)
                            continue
                    except Exception:
                        pass
                # Не 429 или кончились попытки — выводим ошибку и выходим
                print(f"⚠️ Ошибка при запросе к Groq API: {e}")
                if hasattr(e, 'response') and e.response is not None:
                    try:
                        error_data = e.response.json()
                        error_msg = error_data.get('error', {}).get('message', '')
                        if e.response.status_code == 429:
                            print(f"\n   ⚠ Лимит TPM. В .env задайте GROQ_MODEL=groq/compound-mini или подождите минуту.")
                        else:
                            print(f"   Детали ошибки: {error_data}")
                    except Exception:
                        print(f"   Ответ сервера: {getattr(e.response, 'text', '')[:200]}")
                return None, None, None
        
        return None, None, None
    
    except Exception as e:
        print(f"⚠️ Неожиданная ошибка: {e}")
        return None, None, None


def start_chat_session():
    """
    Запускает интерактивную сессию чата с ИИ.
    """
    print("\n· OzonReportX AI — чат с ИИ. Вопросы по бизнесу Ozon. Выход: выход / exit / quit\n")
    
    if not GROQ_API_KEY:
        print("⚠️ GROQ_API_KEY не настроен в .env файле.")
        print("   Добавьте GROQ_API_KEY=ваш_ключ в файл .env")
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
    
    print(f"✅ Готов. Модель: {GROQ_MODEL}\n")
    
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
            context_state = "\n".join(context_state_lines) if context_state_lines else None
            
            # РЕЖИМ A: План/Уточнение
            print("   …", end="\r")
            ai_response, usage_info, rate_limit_info = chat_with_ai(
                user_input, 
                conversation_history, 
                data_block=None,
                context_state=context_state,
                temperature=0.7,
                max_tokens=600  # Ограничиваем для режима A
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
                    retry_response = requests.post(
                        GROQ_API_URL,
                        headers={
                            "Authorization": f"Bearer {GROQ_API_KEY}",
                            "Content-Type": "application/json"
                        },
                        json={
                            "model": GROQ_MODEL,
                            "messages": retry_messages,
                            "temperature": 0.2,  # Низкая температура для точного формата
                            "max_tokens": 500
                        },
                        timeout=30
                    )
                    retry_response.raise_for_status()
                    retry_data = retry_response.json()
                    retry_ai_response = retry_data.get("choices", [{}])[0].get("message", {}).get("content", "")
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
                data_block = format_tool_results(tool_results)
                
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
                    max_tokens=2000
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
                        additional_data = format_tool_results(additional_results)
                        
                        # Объединяем с предыдущими данными
                        combined_data = f"{data_block}\n\nДОПОЛНИТЕЛЬНЫЕ ДАННЫЕ:\n{additional_data}"
                        
                        # Запрашиваем финальный ответ с объединёнными данными
                        current_response, _, rate_limit_info2 = chat_with_ai(
                            user_input,
                            conversation_history,
                            data_block=combined_data,
                            context_state=context_state,
                            temperature=0.2,  # Низкая температура для режима B
                            max_tokens=2000
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
                        final_answer = current_response
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
                            final_answer = "Проанализировал полученные данные. Пожалуйста, уточните ваш вопрос или запросите конкретные метрики."
                    else:
                        final_answer = current_response
                    
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
            
            # Выводим информацию о модели и лимитах из Groq API
            model = rate_limit_info.get("model", GROQ_MODEL) if rate_limit_info else GROQ_MODEL
            limit = rate_limit_info.get("limit") if rate_limit_info else None
            remaining = rate_limit_info.get("remaining") if rate_limit_info else None
            
            info_lines = [f"🤖 Модель: {model}"]
            
            if limit is not None and remaining is not None:
                try:
                    limit_val = int(limit)
                    remaining_val = int(remaining)
                    info_lines.append(f"📊 Лимит: {remaining_val:,}/{limit_val:,}")
                except (ValueError, TypeError):
                    pass
            
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
