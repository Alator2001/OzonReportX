# -*- coding: utf-8 -*-
"""
Модуль управления ценами - разделение логики на отдельные действия.
"""

import json
import sys
from pathlib import Path
from typing import Dict, Optional, Tuple

# Добавляем путь к scripts для импорта
script_dir = Path(__file__).resolve().parent
if str(script_dir) not in sys.path:
    sys.path.insert(0, str(script_dir))

from recommended_prices import (
    load_margin_settings,
    save_margin_settings,
    MIN_MARGIN_DEFAULT,
    DESIRED_MARGIN_DEFAULT,
    load_costs_df,
    COSTS_FILENAME,
    compute_prices,
    get_product_prices_from_ozon,
    get_actions_for_products,
    COL_MIN_PRICE,
    COL_DESIRED_PRICE,
    COL_CURRENT_PRICE,
    COL_MARKETING_PRICE,
    COL_CURRENT_MARGIN,
    compute_current_margin,
    get_report_path,
    get_prev_month_year,
    load_rates_from_report,
    generate_monthly_report,
    MONTHS_RU,
    collect_deactivation_candidates_from_sheet,
    deactivate_products_in_action,
    get_action_candidates,
    activate_products_in_action,
    _artikul_normalize,
    _normalize_offer_id,
    get_discount_requests,
    approve_discount_requests,
    decline_discount_requests,
    get_sku_to_offer_id_mapping,
)

try:
    from utils import prompt_yes_no, print_step, log_verbose
except ImportError:
    def prompt_yes_no(prompt: str, default_yes: bool = True) -> bool:
        default_str = "Y/n" if default_yes else "y/N"
        response = input(f"{prompt} ({default_str}): ").strip().lower()
        if not response:
            return default_yes
        return response in ("y", "yes", "да", "д")
    
    def print_step(text: str):
        print(f"\n· {text}")
    
    def log_verbose(_msg: str) -> None:
        pass

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter


def action_set_margin_range(repo_root: Path) -> Tuple[float, float]:
    """
    Действие 1: Диапазон рентабельности.
    Пользователь задаёт диапазон минимальной и желательной рентабельности.
    """
    print_step("Диапазон рентабельности")
    
    # Загружаем сохранённые настройки
    saved_min, saved_desired = load_margin_settings(repo_root)
    
    if saved_min is not None and saved_desired is not None:
        print(f"Текущие настройки:")
        print(f"  Минимальная рентабельность: {saved_min*100:.1f}%")
        print(f"  Желательная рентабельность: {saved_desired*100:.1f}%")
        
        if not prompt_yes_no("Изменить настройки?", default_yes=False):
            return saved_min, saved_desired
    
    # Запрашиваем новые значения
    while True:
        try:
            min_input = input(f"Минимальная рентабельность (доля 0-1, по умолчанию {MIN_MARGIN_DEFAULT}): ").strip()
            if not min_input:
                min_margin = MIN_MARGIN_DEFAULT
            else:
                min_margin = float(min_input.replace(",", "."))
                if not (0 < min_margin < 1):
                    print(f"Рентабельность должна быть между 0 и 1. Используется значение по умолчанию {MIN_MARGIN_DEFAULT}.")
                    min_margin = MIN_MARGIN_DEFAULT
            
            desired_input = input(f"Желательная рентабельность (доля 0-1, по умолчанию {DESIRED_MARGIN_DEFAULT}): ").strip()
            if not desired_input:
                desired_margin = DESIRED_MARGIN_DEFAULT
            else:
                desired_margin = float(desired_input.replace(",", "."))
                if not (0 < desired_margin < 1):
                    print(f"Рентабельность должна быть между 0 и 1. Используется значение по умолчанию {DESIRED_MARGIN_DEFAULT}.")
                    desired_margin = DESIRED_MARGIN_DEFAULT
            
            if min_margin >= desired_margin:
                print("⚠️ Минимальная рентабельность должна быть меньше желательной. Попробуйте снова.")
                continue
            
            break
        except (ValueError, KeyboardInterrupt):
            print("⚠️ Некорректный ввод. Используются значения по умолчанию.")
            min_margin = MIN_MARGIN_DEFAULT
            desired_margin = DESIRED_MARGIN_DEFAULT
            break
    
    # Сохраняем настройки
    save_margin_settings(repo_root, min_margin, desired_margin)
    
    return min_margin, desired_margin


def action_calculate_optimal_prices(repo_root: Path) -> bool:
    """
    Действие 2: Рассчитать оптимальную цену.
    Расчёт колонок Минимальная цена продажи и Желательная цена продажи.
    """
    print_step("Рассчитать оптимальную цену")
    
    # Проверяем наличие настроек рентабельности
    min_margin, desired_margin = load_margin_settings(repo_root)
    
    if min_margin is None or desired_margin is None:
        print("⚠️ Диапазон рентабельности не задан.")
        if prompt_yes_no("Задать диапазон рентабельности сейчас?", default_yes=True):
            min_margin, desired_margin = action_set_margin_range(repo_root)
        else:
            print("❌ Невозможно рассчитать цены без диапазона рентабельности.")
            return False
    
    # Получаем отчёт для расчёта комиссии
    prev_year, prev_month = get_prev_month_year()
    report_path = get_report_path(repo_root, prev_year, prev_month)
    costs_path = repo_root / COSTS_FILENAME
    
    print(f"Используется отчёт за предыдущий месяц: {MONTHS_RU[prev_month - 1]} {prev_year}")
    
    if not report_path.exists():
        print(f"⚠ Файл отчёта не найден: {report_path.name}")
        if prompt_yes_no("Сгенерировать отчёт за предыдущий месяц?", default_yes=True):
            try:
                report_path = generate_monthly_report(repo_root, prev_month, prev_year)
            except Exception as e:
                print(f"❌ Не удалось сгенерировать отчёт: {e}")
                return False
        else:
            print("❌ Невозможно рассчитать цены без отчёта.")
            return False
    
    log_verbose(f"Файл отчёта: {report_path}")
    total_rate = load_rates_from_report(report_path)
    log_verbose(f"Комиссия+логистика: {total_rate*100:.2f}%")
    df, key_col, cost_col = load_costs_df(costs_path)
    log_verbose(f"Загружено записей: {len(df)}")
    
    # Рассчитываем цены
    min_prices = []
    desired_prices = []

    def _parse_cost(val):
        """Преобразует значение себестоимости в float (поддержка запятой как разделителя)."""
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        s = str(val).strip().replace(",", ".")
        if not s or s.lower() == "nan":
            return 0.0
        try:
            return float(s)
        except (TypeError, ValueError):
            return 0.0

    for _, row in df.iterrows():
        raw = row.get(cost_col, 0)
        cost_val = _parse_cost(raw)
        min_p, des_p = compute_prices(cost_val, total_rate, min_margin, desired_margin)
        min_prices.append(min_p)
        desired_prices.append(des_p)

    # Если ни одна цена не посчиталась — предупреждаем
    filled_min = sum(1 for p in min_prices if p is not None)
    if filled_min == 0 and len(df) > 0:
        print("⚠️ Не удалось рассчитать ни одной цены. Проверьте:")
        print(f"   — столбец «{cost_col}»: все значения должны быть положительными числами (можно с запятой);")
        print(f"   — комиссия+логистика из отчёта: {total_rate*100:.1f}%. Сумма с маржой не должна превышать 100%.")
        print(f"   (Маржа мин/жел: {min_margin*100:.0f}% / {desired_margin*100:.0f}%).")

    # Обновляем колонки
    if COL_MIN_PRICE in df.columns:
        df = df.drop(columns=[COL_MIN_PRICE])
    if COL_DESIRED_PRICE in df.columns:
        df = df.drop(columns=[COL_DESIRED_PRICE])
    
    df[COL_MIN_PRICE] = min_prices
    df[COL_DESIRED_PRICE] = desired_prices
    
    df.to_excel(costs_path, index=False)
    print(f"✅ Рассчитаны оптимальные цены для {len(df)} товаров (маржа {min_margin*100:.0f}% / {desired_margin*100:.0f}%).")
    return True


def action_get_current_prices(repo_root: Path) -> bool:
    """
    Действие 3: Узнать текущую цену продажи.
    Расчёт колонок Текущая цена на Ozon, Цена с учётом акций и скидок, Ожидаемая рентабельность.
    """
    print_step("Узнать текущую цену продажи")
    
    costs_path = repo_root / COSTS_FILENAME
    
    if not costs_path.exists():
        print(f"❌ Файл {COSTS_FILENAME} не найден.")
        return False
    
    df, key_col, cost_col = load_costs_df(costs_path)
    log_verbose(f"Загружено записей: {len(df)}")
    if COL_MIN_PRICE not in df.columns:
        print(f"⚠️ Колонка «{COL_MIN_PRICE}» не найдена.")
        if prompt_yes_no("Рассчитать оптимальные цены сейчас?", default_yes=True):
            if not action_calculate_optimal_prices(repo_root):
                return False
            # Перезагружаем данные
            df, key_col, cost_col = load_costs_df(costs_path)
        else:
            print("❌ Невозможно рассчитать рентабельность без минимальной цены.")
            return False
    
    # Получаем отчёт для расчёта комиссии
    prev_year, prev_month = get_prev_month_year()
    report_path = get_report_path(repo_root, prev_year, prev_month)
    
    if not report_path.exists():
        print(f"⚠ Файл отчёта не найден: {report_path.name}")
        if prompt_yes_no("Сгенерировать отчёт за предыдущий месяц?", default_yes=True):
            try:
                report_path = generate_monthly_report(repo_root, prev_month, prev_year)
            except Exception as e:
                print(f"❌ Не удалось сгенерировать отчёт: {e}")
                return False
        else:
            print("❌ Невозможно рассчитать рентабельность без отчёта.")
            return False
    
    total_rate = load_rates_from_report(report_path)
    log_verbose(f"Комиссия+логистика: {total_rate*100:.2f}%")
    log_verbose("Получение цен с Ozon...")
    offer_ids_list = []
    for _, row in df.iterrows():
        art = _artikul_normalize(row.get(key_col))
        if art:
            offer_ids_list.append(art)
    
    prices_map, marketing_prices_map = get_product_prices_from_ozon(offer_ids_list)
    if prices_map:
        print(f"✅ Получено цен для {len(prices_map)} артикулов")
    else:
        print("⚠️ Не удалось получить цены с Ozon (возможно, не настроены API ключи).")
    
    # Рассчитываем текущие цены и рентабельность
    current_prices = []
    marketing_prices = []
    current_margins = []
    
    for _, row in df.iterrows():
        try:
            cost_val = float(row.get(cost_col, 0) or 0)
        except (TypeError, ValueError):
            cost_val = 0.0
        
        art = _artikul_normalize(row.get(key_col))
        if art:
            art_normalized = _normalize_offer_id(art)
            current_price = prices_map.get(art) or prices_map.get(art_normalized)
            marketing_price = marketing_prices_map.get(art) or marketing_prices_map.get(art_normalized)
        else:
            current_price = None
            marketing_price = None
        
        current_prices.append(round(current_price) if current_price is not None else None)
        marketing_prices.append(round(marketing_price) if marketing_price is not None else None)
        
        price_for_margin = marketing_price if marketing_price is not None else current_price
        margin = compute_current_margin(price_for_margin, cost_val, total_rate)
        current_margins.append(round(margin * 100, 2) if margin is not None else None)
    
    # Обновляем колонки
    for c in [COL_CURRENT_PRICE, COL_MARKETING_PRICE, COL_CURRENT_MARGIN]:
        if c in df.columns:
            df = df.drop(columns=[c])
    
    df[COL_CURRENT_PRICE] = current_prices
    df[COL_MARKETING_PRICE] = marketing_prices
    df[COL_CURRENT_MARGIN] = current_margins
    
    df.to_excel(costs_path, index=False)
    
    # Применяем условное форматирование
    try:
        wb = load_workbook(costs_path)
        ws = wb.active
        
        # Форматирование для рентабельности
        min_margin, desired_margin = load_margin_settings(repo_root)
        if min_margin is None:
            min_margin = MIN_MARGIN_DEFAULT
        if desired_margin is None:
            desired_margin = DESIRED_MARGIN_DEFAULT
        
        margin_col_idx = None
        for col_idx, cell in enumerate(ws[1], start=1):
            if cell.value == COL_CURRENT_MARGIN:
                margin_col_idx = col_idx
                break
        
        if margin_col_idx:
            min_margin_pct = min_margin * 100
            desired_margin_pct = desired_margin * 100
            
            green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            green_rule = CellIsRule(
                operator="between",
                formula=[min_margin_pct, desired_margin_pct],
                fill=green_fill
            )
            
            red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            red_rule = CellIsRule(
                operator="lessThan",
                formula=[min_margin_pct],
                fill=red_fill
            )
            
            margin_col_letter = ws.cell(row=1, column=margin_col_idx).column_letter
            data_range = f"{margin_col_letter}2:{margin_col_letter}{len(df) + 1}"
            ws.conditional_formatting.add(data_range, green_rule)
            ws.conditional_formatting.add(data_range, red_rule)
        
        # Форматирование для текущей цены: красный < мин, зелёный >= мин, более зелёный >= желательной
        current_price_col_idx = None
        min_price_col_idx = None
        desired_price_col_idx = None
        for col_idx, cell in enumerate(ws[1], start=1):
            if cell.value == COL_CURRENT_PRICE:
                current_price_col_idx = col_idx
            elif cell.value == COL_MIN_PRICE:
                min_price_col_idx = col_idx
            elif cell.value == COL_DESIRED_PRICE:
                desired_price_col_idx = col_idx
        
        if current_price_col_idx and min_price_col_idx:
            current_price_col_letter = ws.cell(row=1, column=current_price_col_idx).column_letter
            min_price_col_letter = ws.cell(row=1, column=min_price_col_idx).column_letter
            data_range = f"{current_price_col_letter}2:{current_price_col_letter}{len(df) + 1}"
            
            red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            red_formula = f"AND({current_price_col_letter}2<>\"\", {current_price_col_letter}2>0, {current_price_col_letter}2<{min_price_col_letter}2)"
            red_rule = FormulaRule(formula=[red_formula], fill=red_fill, stopIfTrue=False)
            
            green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            green_formula = f"AND({current_price_col_letter}2<>\"\", {current_price_col_letter}2>0, {current_price_col_letter}2>={min_price_col_letter}2)"
            green_rule = FormulaRule(formula=[green_formula], fill=green_fill, stopIfTrue=False)
            
            ws.conditional_formatting.add(data_range, red_rule)
            ws.conditional_formatting.add(data_range, green_rule)
            
            # Цена выше диапазона (>= желательной) — более насыщенный зелёный (применяется поверх обычного зелёного)
            if desired_price_col_idx:
                desired_price_col_letter = ws.cell(row=1, column=desired_price_col_idx).column_letter
                dark_green_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                dark_green_formula = f"AND({current_price_col_letter}2<>\"\", {current_price_col_letter}2>0, {current_price_col_letter}2>={desired_price_col_letter}2)"
                dark_green_rule = FormulaRule(formula=[dark_green_formula], fill=dark_green_fill, stopIfTrue=True)
                ws.conditional_formatting.add(data_range, dark_green_rule)
        
        # Форматирование для цены с акциями
        marketing_price_col_idx = None
        for col_idx, cell in enumerate(ws[1], start=1):
            if cell.value == COL_MARKETING_PRICE:
                marketing_price_col_idx = col_idx
                break
        
        if marketing_price_col_idx and min_price_col_idx:
            marketing_price_col_letter = ws.cell(row=1, column=marketing_price_col_idx).column_letter
            
            green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            green_formula = f"AND({marketing_price_col_letter}2<>\"\", {marketing_price_col_letter}2>0, {marketing_price_col_letter}2>={min_price_col_letter}2)"
            green_rule = FormulaRule(formula=[green_formula], fill=green_fill, stopIfTrue=False)
            
            red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            red_formula = f"AND({marketing_price_col_letter}2<>\"\", {marketing_price_col_letter}2>0, {marketing_price_col_letter}2<{min_price_col_letter}2)"
            red_rule = FormulaRule(formula=[red_formula], fill=red_fill, stopIfTrue=False)
            
            data_range = f"{marketing_price_col_letter}2:{marketing_price_col_letter}{len(df) + 1}"
            ws.conditional_formatting.add(data_range, red_rule)
            ws.conditional_formatting.add(data_range, green_rule)
        
        wb.save(costs_path)
        print("✅ Применено условное форматирование к колонкам.")
    except Exception as e:
        print(f"⚠️ Не удалось применить условное форматирование: {e}")
    
    print("✅ Текущие цены получены и сохранены.")
    return True


def action_get_active_actions(repo_root: Path) -> bool:
    """
    Действие 4: Узнать активные акции.
    Просмотр в каких акциях участвует товар и по какой цене.
    """
    print_step("Узнать активные акции")
    
    costs_path = repo_root / COSTS_FILENAME
    
    if not costs_path.exists():
        print(f"❌ Файл {COSTS_FILENAME} не найден.")
        return False
    
    df, key_col, cost_col = load_costs_df(costs_path)
    print(f"Загружен файл себестоимости: {len(df)} записей.")
    
    # Получаем артикулы
    offer_ids_list = []
    for _, row in df.iterrows():
        art = _artikul_normalize(row.get(key_col))
        if art:
            offer_ids_list.append(art)
    
    # Получаем информацию об акциях
    actions_map, actions_info_list, _ = get_actions_for_products(offer_ids_list)
    
    if not actions_info_list:
        print("⚠️ Активные акции не найдены.")
        return False
    
    log_verbose(f"Найдено акций: {len(actions_info_list)}")
    
    # Создаём DataFrame для листа акций
    actions_df_data = {}
    actions_df_data[key_col] = df[key_col].values
    
    # Проверяем наличие минимальной цены
    if COL_MIN_PRICE not in df.columns:
        print(f"⚠️ Колонка «{COL_MIN_PRICE}» не найдена.")
        if prompt_yes_no("Рассчитать оптимальные цены сейчас?", default_yes=True):
            if not action_calculate_optimal_prices(repo_root):
                return False
            df, key_col, cost_col = load_costs_df(costs_path)
        else:
            print("⚠️ Продолжаем без минимальной цены.")
    
    if COL_MIN_PRICE in df.columns:
        actions_df_data[COL_MIN_PRICE] = df[COL_MIN_PRICE].values
    
    # Получаем цены в акциях для каждого товара
    action_prices_dicts = {}
    for action_info in actions_info_list:
        action_name = action_info["name"]
        action_prices_dicts[action_name] = []
    
    for _, row in df.iterrows():
        art = _artikul_normalize(row.get(key_col))
        if art:
            art_normalized = _normalize_offer_id(art)
            art_actions = actions_map.get(art) or actions_map.get(art_normalized) or {}
        else:
            art_actions = {}
        
        for action_info in actions_info_list:
            action_name = action_info["name"]
            action_price = art_actions.get(action_name)
            if action_price is not None:
                action_prices_dicts[action_name].append(round(action_price))
            else:
                action_prices_dicts[action_name].append(None)
    
    # Добавляем колонки акций
    for action_info in actions_info_list:
        action_name = action_info["name"]
        actions_df_data[action_name] = action_prices_dicts[action_name]
    
    actions_df = pd.DataFrame(actions_df_data)
    
    # Сохраняем в Excel
    try:
        wb = load_workbook(costs_path)
        
        if 'Sheet1' in wb.sheetnames:
            wb['Sheet1'].title = 'Основной'
        
        if "Акции" in wb.sheetnames:
            wb.remove(wb["Акции"])
        
        ws_actions = wb.create_sheet("Акции")
        
        # Записываем заголовки
        for c_idx, col_name in enumerate(actions_df.columns, start=1):
            ws_actions.cell(row=1, column=c_idx, value=col_name)
        
        # Записываем данные
        for r_idx, row in enumerate(actions_df.itertuples(index=False), start=2):
            for c_idx, value in enumerate(row, start=1):
                ws_actions.cell(row=r_idx, column=c_idx, value=value)
        
        # Применяем условное форматирование
        if COL_MIN_PRICE in actions_df.columns:
            min_price_col_idx = None
            for col_idx, col_name in enumerate(actions_df.columns, start=1):
                if col_name == COL_MIN_PRICE:
                    min_price_col_idx = col_idx
                    break
            
            if min_price_col_idx:
                min_price_col_letter = get_column_letter(min_price_col_idx)
                
                for col_idx, col_name in enumerate(actions_df.columns, start=1):
                    if col_name == key_col or col_name == COL_MIN_PRICE:
                        continue
                    
                    action_col_letter = get_column_letter(col_idx)
                    
                    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                    green_formula = f"AND({action_col_letter}2<>\"\", {action_col_letter}2>0, {action_col_letter}2>={min_price_col_letter}2)"
                    green_rule = FormulaRule(formula=[green_formula], fill=green_fill, stopIfTrue=False)
                    
                    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                    red_formula = f"AND({action_col_letter}2<>\"\", {action_col_letter}2>0, {action_col_letter}2<{min_price_col_letter}2)"
                    red_rule = FormulaRule(formula=[red_formula], fill=red_fill, stopIfTrue=False)
                    
                    data_range = f"{action_col_letter}2:{action_col_letter}{len(actions_df) + 1}"
                    ws_actions.conditional_formatting.add(data_range, red_rule)
                    ws_actions.conditional_formatting.add(data_range, green_rule)
        
        wb.save(costs_path)
        print(f"✅ Создан лист «Акции» с {len(actions_info_list)} колонками акций.")
    except Exception as e:
        print(f"⚠️ Ошибка при создании листа «Акции»: {e}")
        import traceback
        print(traceback.format_exc())
        return False
    
    return True


def action_remove_unprofitable_actions(repo_root: Path) -> bool:
    """
    Действие 5: Удалить невыгодные акции.
    Удаление товаров из акций, где цена меньше минимальной.
    """
    print_step("Удалить невыгодные акции")
    
    costs_path = repo_root / COSTS_FILENAME
    
    if not costs_path.exists():
        print(f"❌ Файл {COSTS_FILENAME} не найден.")
        return False
    
    # Проверяем наличие листа "Акции"
    try:
        wb = load_workbook(costs_path)
        if "Акции" not in wb.sheetnames:
            print("⚠️ Лист «Акции» не найден.")
            if prompt_yes_no("Получить информацию об активных акциях сейчас?", default_yes=True):
                if not action_get_active_actions(repo_root):
                    return False
                wb = load_workbook(costs_path)
            else:
                print("❌ Невозможно удалить из акций без информации об акциях.")
                return False
    except Exception as e:
        print(f"❌ Ошибка при открытии файла: {e}")
        return False
    
    df, key_col, cost_col = load_costs_df(costs_path)
    
    # Получаем артикулы и маппинг
    offer_ids_list = []
    for _, row in df.iterrows():
        art = _artikul_normalize(row.get(key_col))
        if art:
            offer_ids_list.append(art)
    
    _, actions_info_list, offer_id_to_product_id = get_actions_for_products(offer_ids_list)
    
    if not actions_info_list:
        print("⚠️ Активные акции не найдены.")
        return False
    
    ws_actions = wb['Акции']
    action_name_to_id = {a["name"]: a["id"] for a in actions_info_list}
    
    if not offer_id_to_product_id or not action_name_to_id:
        print("⚠️ Недостаточно данных для удаления из акций.")
        return False
    
    candidates = collect_deactivation_candidates_from_sheet(
        ws_actions,
        key_col,
        COL_MIN_PRICE,
        action_name_to_id,
        offer_id_to_product_id,
    )
    
    if not candidates:
        print("✅ Товары с ценой ниже минимальной не найдены.")
        return True
    
    total_to_remove = sum(len(ids) for ids in candidates.values())
    print(f"Найдено {total_to_remove} товаров для удаления из {len(candidates)} акций.")
    
    if not prompt_yes_no("Продолжить удаление?", default_yes=False):
        print("❌ Удаление отменено.")
        return False
    
    log_verbose("Удаление товаров из акций...")
    for action_id, product_ids in candidates.items():
        result = deactivate_products_in_action(action_id, product_ids)
        removed = result.get("product_ids", []) or []
        rejected = result.get("rejected", []) or []
        log_verbose(f"Акция {action_id}: удалено {len(removed)}, не удалено {len(rejected)}")
    print("✅ Удаление из акций завершено.")
    return True


def action_add_to_actions(repo_root: Path) -> bool:
    """
    Действие 6: Добавить товары в акции.
    Добавление товаров, если допустимая цена в диапазоне минимальной и желаемой.
    """
    print_step("Добавить товары в акции")
    
    costs_path = repo_root / COSTS_FILENAME
    
    if not costs_path.exists():
        print(f"❌ Файл {COSTS_FILENAME} не найден.")
        return False
    
    # Проверяем наличие минимальной и желательной цены
    df, key_col, cost_col = load_costs_df(costs_path)
    
    if COL_MIN_PRICE not in df.columns or COL_DESIRED_PRICE not in df.columns:
        print(f"⚠️ Колонки «{COL_MIN_PRICE}» или «{COL_DESIRED_PRICE}» не найдены.")
        if prompt_yes_no("Рассчитать оптимальные цены сейчас?", default_yes=True):
            if not action_calculate_optimal_prices(repo_root):
                return False
            df, key_col, cost_col = load_costs_df(costs_path)
        else:
            print("❌ Невозможно добавить в акции без оптимальных цен.")
            return False
    
    # Получаем артикулы и маппинг
    offer_ids_list = []
    for _, row in df.iterrows():
        art = _artikul_normalize(row.get(key_col))
        if art:
            offer_ids_list.append(art)
    
    _, actions_info_list, offer_id_to_product_id = get_actions_for_products(offer_ids_list)
    
    if not actions_info_list:
        print("⚠️ Активные акции не найдены.")
        return False
    
    # Создаём маппинг product_id -> offer_id
    product_id_to_offer_id = {}
    for offer_id, product_id in offer_id_to_product_id.items():
        product_id_to_offer_id[product_id] = offer_id
    
    offer_ids_set_for_candidates = set()
    for oid in offer_ids_list:
        normalized = _normalize_offer_id(oid)
        if normalized:
            offer_ids_set_for_candidates.add(normalized)
            offer_ids_set_for_candidates.add(oid)
    
    # Создаём маппинг product_id -> (min_price, desired_price)
    product_id_to_prices: Dict[int, Tuple[float, float]] = {}
    for _, row in df.iterrows():
        art = _artikul_normalize(row.get(key_col))
        if not art:
            continue
        
        art_normalized = _normalize_offer_id(art)
        product_id = offer_id_to_product_id.get(art_normalized) or offer_id_to_product_id.get(art)
        
        if product_id:
            min_p = row.get(COL_MIN_PRICE)
            des_p = row.get(COL_DESIRED_PRICE)
            if min_p is not None and des_p is not None:
                try:
                    min_price_val = float(min_p)
                    des_price_val = float(des_p)
                    if min_price_val > 0 and des_price_val > 0:
                        product_id_to_prices[product_id] = (min_price_val, des_price_val)
                except (TypeError, ValueError):
                    pass
    
    if not product_id_to_prices:
        print("⚠️ Не найдено товаров с рассчитанными ценами.")
        return False
    
    log_verbose("Проверка кандидатов для добавления в акции...")
    preview_rows = []
    for action_info in actions_info_list:
        action_id = action_info["id"]
        action_name = action_info["name"]
        candidates = get_action_candidates(action_id, product_id_to_offer_id, offer_ids_set_for_candidates)
        if not candidates:
            continue
        products_to_add = []
        for product_id, product_info in candidates.items():
            if product_id not in product_id_to_prices:
                continue
            
            min_price, desired_price = product_id_to_prices[product_id]
            
            max_action_price = product_info.get("max_action_price")
            if max_action_price is None:
                continue
            
            try:
                max_action_price_val = float(max_action_price)
            except (TypeError, ValueError):
                continue
            
            target_price = min(desired_price, max_action_price_val)
            
            if target_price >= min_price:
                current_action_price = product_info.get("action_price", 0)
                if current_action_price == 0 or current_action_price is None:
                    stock = product_info.get("stock", 0) or 0
                    offer_id = product_id_to_offer_id.get(product_id, "—")
                    products_to_add.append({
                        "product_id": product_id,
                        "action_price": int(target_price),
                        "stock": int(stock) if stock else 0
                    })
                    preview_rows.append({
                        "product_id": product_id,
                        "offer_id": offer_id,
                        "action_name": action_name,
                        "action_price": int(target_price),
                        "stock": int(stock) if stock else 0,
                    })

    if not preview_rows:
        print("⚠️ Не найдено товаров, которые можно добавить в акции.")
        return False

    print("✅ Кандидаты для добавления в акции:")
    for row in preview_rows:
        print(f"   {row['offer_id']} | {row['action_name']} | {row['action_price']}")

    if not prompt_yes_no("Продолжить добавление товаров в акции?", default_yes=False):
        print("❌ Добавление отменено.")
        return False

    total_added = 0
    for action_info in actions_info_list:
        action_id = action_info["id"]
        action_name = action_info["name"]
        products_to_add = []
        for row in preview_rows:
            if row["action_name"] != action_name:
                continue
            products_to_add.append({
                "product_id": row["product_id"],
                "action_price": row["action_price"],
                "stock": row["stock"],
            })

        if products_to_add:
            result = activate_products_in_action(action_id, products_to_add)
            added = result.get("product_ids", []) or []
            rejected = result.get("rejected", []) or []
            total_added += len(added)
            log_verbose(f"Акция {action_name}: добавлено {len(added)}, не добавлено {len(rejected)}")
    print(f"✅ Добавление в акции завершено. Всего добавлено: {total_added} товаров.")
    return True


def action_get_current_prices_and_actions(repo_root: Path) -> bool:
    """
    Действие: Узнать текущие цены и акции.
    Последовательно обновляет текущие цены продажи и активные акции в costs.xlsx.
    """
    print_step("Узнать текущие цены и акции")
    prices_ok = action_get_current_prices(repo_root)
    actions_ok = action_get_active_actions(repo_root)
    return bool(prices_ok and actions_ok)


def action_process_discount_requests(repo_root: Path) -> bool:
    """
    Действие 7: Обработать заявки на скидку.
    Одобряет заявки, если заявленная цена >= минимальной цены продажи, иначе отклоняет.
    """
    print_step("Обработать заявки на скидку")
    
    costs_path = repo_root / COSTS_FILENAME
    
    if not costs_path.exists():
        print(f"❌ Файл {COSTS_FILENAME} не найден.")
        return False
    
    # Проверяем наличие колонки минимальной цены
    df, key_col, cost_col = load_costs_df(costs_path)
    
    if COL_MIN_PRICE not in df.columns:
        print(f"⚠️ Колонка «{COL_MIN_PRICE}» не найдена.")
        if prompt_yes_no("Рассчитать оптимальные цены сейчас?", default_yes=True):
            if not action_calculate_optimal_prices(repo_root):
                return False
            df, key_col, cost_col = load_costs_df(costs_path)
        else:
            print("❌ Невозможно обработать заявки без минимальной цены.")
            return False
    
    # Создаём маппинг offer_id -> min_price
    offer_id_to_min_price: Dict[str, float] = {}
    for _, row in df.iterrows():
        art = _artikul_normalize(row.get(key_col))
        if not art:
            continue
        
        art_normalized = _normalize_offer_id(art)
        min_price = row.get(COL_MIN_PRICE)
        
        if min_price is not None:
            try:
                min_price_val = float(min_price)
                if min_price_val > 0:
                    offer_id_to_min_price[art_normalized] = min_price_val
                    # Также добавляем исходный артикул
                    if art != art_normalized:
                        offer_id_to_min_price[art] = min_price_val
            except (TypeError, ValueError):
                pass
    
    if not offer_id_to_min_price:
        print("⚠️ Не найдено товаров с минимальной ценой.")
        return False
    
    print(f"✅ Загружено {len(offer_id_to_min_price)} товаров с минимальной ценой.")
    
    # Получаем заявки на скидку
    print("📡 Получение заявок на скидку...")
    discount_tasks = get_discount_requests(status="NEW", limit=50)
    
    if not discount_tasks:
        print("✅ Новых заявок на скидку не найдено.")
        return True
    
    print(f"✅ Найдено {len(discount_tasks)} заявок на скидку.")
    
    # Ozon в заявках возвращает SKU (product_id), в costs.xlsx записан offer_id (артикул).
    # Получаем маппинг SKU -> offer_id через API.
    skus_from_tasks = []
    for task in discount_tasks:
        sku = task.get("sku")
        if sku is not None:
            try:
                skus_from_tasks.append(int(sku))
            except (TypeError, ValueError):
                pass
    skus_unique = list(dict.fromkeys(skus_from_tasks))
    sku_to_offer_id: Dict[int, str] = {}
    if skus_unique:
        sku_to_offer_id = get_sku_to_offer_id_mapping(skus_unique)
    
    log_verbose(f"Обработка {len(discount_tasks)} заявок...")
    
    tasks_to_approve = []
    tasks_to_decline = []
    task_id_to_sku: Dict[str, str] = {}
    
    for task in discount_tasks:
        sku = task.get("sku")
        task_id = task.get("id")
        if task_id is not None:
            task_id_to_sku[str(task_id)] = str(sku) if sku else "—"
        
        if not sku:
            continue
        
        # Определяем offer_id: сначала по маппингу SKU -> offer_id, иначе считаем sku артикулом (offer_id)
        offer_id_raw = None
        try:
            sku_int = int(sku)
            offer_id_raw = sku_to_offer_id.get(sku_int)
        except (TypeError, ValueError):
            pass
        if not offer_id_raw:
            offer_id_raw = str(sku)
        
        offer_id_normalized = _normalize_offer_id(offer_id_raw)
        min_price = offer_id_to_min_price.get(offer_id_normalized) or offer_id_to_min_price.get(offer_id_raw)
        
        if min_price is None:
            # Если не найдена минимальная цена, отклоняем
            reason = "Минимальная цена не рассчитана"
            tasks_to_decline.append({
                "id": task_id,
                "seller_comment": reason
            })
            continue
        
        requested_price = task.get("requested_price")
        if requested_price is None:
            # Если нет запрошенной цены, отклоняем
            reason = "Запрошенная цена не указана"
            tasks_to_decline.append({
                "id": task_id,
                "seller_comment": reason
            })
            continue
        
        try:
            requested_price_val = float(requested_price)
        except (TypeError, ValueError):
            reason = "Некорректная запрошенная цена"
            tasks_to_decline.append({
                "id": task_id,
                "seller_comment": reason
            })
            continue
        
        # Проверяем условие: одобряем, если requested_price >= min_price
        if requested_price_val >= min_price:
            # Одобряем заявку; причина для отображения и для API (seller_comment)
            reason = f"Одобрено: запрошенная цена {requested_price_val:.0f} ₽ не ниже минимальной {min_price:.0f} ₽"
            # API требует approved_quantity_min > 0
            q_min = task.get("requested_quantity_min")
            try:
                q_min = max(1, int(q_min)) if q_min is not None else 1
            except (TypeError, ValueError):
                q_min = 1
            q_max = task.get("requested_quantity_max")
            try:
                q_max = max(q_min, int(q_max)) if q_max is not None else q_min
            except (TypeError, ValueError):
                q_max = q_min
            approved_task = {
                "id": task_id,
                "approved_price": int(requested_price_val),
                "approved_quantity_min": q_min,
                "approved_quantity_max": q_max,
                "seller_comment": reason
            }
            tasks_to_approve.append(approved_task)
        else:
            # Отклоняем заявку
            reason = f"Отклонено: запрошенная цена {requested_price_val:.2f} ₽ ниже минимальной {min_price:.2f} ₽"
            tasks_to_decline.append({
                "id": task_id,
                "seller_comment": reason
            })
    
    print(f"Одобрить: {len(tasks_to_approve)}, отклонить: {len(tasks_to_decline)}")
    if tasks_to_approve or tasks_to_decline:
        for t in tasks_to_approve:
            sid = str(t.get("id", ""))
            sku_display = task_id_to_sku.get(sid, "—")
            print(f"   ✅ Заявка {sid} (артикул {sku_display}): {t.get('seller_comment', 'Одобрено')}")
        for t in tasks_to_decline:
            sid = str(t.get("id", ""))
            sku_display = task_id_to_sku.get(sid, "—")
            print(f"   ❌ Заявка {sid} (артикул {sku_display}): {t.get('seller_comment', 'Отклонено')}")
    
    if not tasks_to_approve and not tasks_to_decline:
        print("⚠️ Нет заявок для обработки.")
        return True
    
    if not prompt_yes_no("Продолжить обработку заявок?", default_yes=False):
        print("❌ Обработка отменена.")
        return False
    
    # Обрабатываем заявки
    log_verbose("Обработка заявок...")
    
    if tasks_to_approve:
        approve_result = approve_discount_requests(tasks_to_approve)
        ok, fail = approve_result.get('success_count', 0), approve_result.get('fail_count', 0)
        print(f"✅ Одобрено: {ok}" + (f", ошибок: {fail}" if fail else ""))
        if fail:
            for detail in (approve_result.get('fail_details') or [])[:3]:
                print(f"   Заявка {detail.get('task_id')}: {detail.get('error_for_user', '?')}")
    if tasks_to_decline:
        decline_result = decline_discount_requests(tasks_to_decline)
        ok, fail = decline_result.get('success_count', 0), decline_result.get('fail_count', 0)
        print(f"❌ Отклонено: {ok}" + (f", ошибок: {fail}" if fail else ""))
        if fail:
            for detail in (decline_result.get('fail_details') or [])[:3]:
                print(f"   Заявка {detail.get('task_id')}: {detail.get('error_for_user', '?')}")
    
    print("✅ Обработка заявок завершена.")
    return True


def refresh_costs_views(repo_root: Path) -> None:
    """Обновляет данные по текущим ценам и активным акциям в costs.xlsx."""
    print_step("Обновление costs.xlsx после изменений")
    action_get_current_prices_and_actions(repo_root)


def show_price_management_menu(repo_root: Path):
    """
    Показывает меню управления ценой и обрабатывает выбор пользователя.
    """
    while True:
        print_step("Управление ценой")
        print("1. Диапазон рентабельности")
        print("2. Рассчитать оптимальную цену")
        print("3. Узнать текущие цены и акции")
        print("4. Удалить невыгодные акции")
        print("5. Добавить товары в акции")
        print("6. Обработать заявки на скидку")
        print("0. Назад в главное меню")
        
        choice = input("Выберите опцию (0-6): ").strip()
        
        if choice == "1":
            action_set_margin_range(repo_root)
        elif choice == "2":
            action_calculate_optimal_prices(repo_root)
        elif choice == "3":
            action_get_current_prices_and_actions(repo_root)
        elif choice == "4":
            if action_remove_unprofitable_actions(repo_root):
                refresh_costs_views(repo_root)
        elif choice == "5":
            if action_add_to_actions(repo_root):
                refresh_costs_views(repo_root)
        elif choice == "6":
            action_process_discount_requests(repo_root)
        elif choice == "0":
            break
        else:
            print("Пожалуйста, выберите корректную опцию (0-6).")
        
        print()  # Пустая строка для читаемости


def main():
    """Точка входа для запуска модуля как скрипта."""
    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parent
    
    show_price_management_menu(repo_root)


if __name__ == "__main__":
    main()
