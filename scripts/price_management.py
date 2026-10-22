# -*- coding: utf-8 -*-
"""
Модуль управления ценами - разделение логики на отдельные действия.
"""

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

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
    COL_MIN_PRICE,
    COL_DESIRED_PRICE,
    COL_CURRENT_PRICE,
    COL_MARKETING_PRICE,
    COL_CURRENT_MARGIN,
    COL_EXPECTED_PROFIT,
    COL_LOGISTICS,
    compute_current_margin,
    get_report_path,
    get_prev_month_year,
    load_cost_structure,
    generate_monthly_report,
    MONTHS_RU,
    action_exit_price,
    activate_products_in_action,
    pick_action_price,
    get_actions_list,
    describe_action,
    boost_at_price,
    fetch_action_products,
    get_product_id_mapping,
    get_product_skus,
    get_search_visibility,
    ACTION_TYPE_ELASTIC,
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
from file_io import write_costs_dataframe, save_workbook_atomic
from file_lock import locked_costs
from contextlib import ExitStack, closing


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


@locked_costs
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
    structure = load_cost_structure(report_path)
    print(f"Расходы Ozon из отчёта: {structure.summary()}.")
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
        logistics = structure.logistics_for(_artikul_normalize(row.get(key_col)))
        min_p, des_p = compute_prices(cost_val, structure.commission_rate, logistics, min_margin, desired_margin)
        min_prices.append(min_p)
        desired_prices.append(des_p)

    # Если ни одна цена не посчиталась — предупреждаем
    filled_min = sum(1 for p in min_prices if p is not None)
    if filled_min == 0 and len(df) > 0:
        print("⚠️ Не удалось рассчитать ни одной цены. Проверьте:")
        print(f"   — столбец «{cost_col}»: все значения должны быть положительными числами (можно с запятой);")
        print(f"   — комиссия из отчёта: {structure.commission_rate*100:.1f}%. Комиссия с маржой не должна превышать 100%.")
        print(f"   (Маржа мин/жел: {min_margin*100:.0f}% / {desired_margin*100:.0f}%).")

    # Обновляем колонки
    if COL_MIN_PRICE in df.columns:
        df = df.drop(columns=[COL_MIN_PRICE])
    if COL_DESIRED_PRICE in df.columns:
        df = df.drop(columns=[COL_DESIRED_PRICE])
    
    df[COL_MIN_PRICE] = min_prices
    df[COL_DESIRED_PRICE] = desired_prices
    
    write_costs_dataframe(df, costs_path)
    print(f"✅ Рассчитаны оптимальные цены для {len(df)} товаров (маржа {min_margin*100:.0f}% / {desired_margin*100:.0f}%).")
    return True


@locked_costs
def action_get_current_prices(repo_root: Path) -> bool:
    """
    Действие 3: Узнать текущую цену продажи.
    Расчёт колонок Текущая цена на Ozon, Цена с учётом акций и скидок, Ожидаемая рентабельность.
    """
    with ExitStack() as books:
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
    
        structure = load_cost_structure(report_path)
        print(f"Расходы Ozon из отчёта: {structure.summary()}.")
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
        expected_profits = []
        logistics_used = []
    
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
            logistics = structure.logistics_for(art)
            logistics_used.append(round(logistics))
            margin = compute_current_margin(price_for_margin, cost_val, structure.commission_rate, logistics)
            current_margins.append(round(margin * 100, 2) if margin is not None else None)
            # Прибыль с единицы по той же формуле, что и рентабельность: цена × (1 − комиссия) − логистика ₽ − себестоимость.
            expected_profits.append(round(margin * price_for_margin) if margin is not None else None)
    
        # Обновляем колонки
        for c in [COL_CURRENT_PRICE, COL_MARKETING_PRICE, COL_CURRENT_MARGIN, COL_EXPECTED_PROFIT, COL_LOGISTICS]:
            if c in df.columns:
                df = df.drop(columns=[c])
    
        df[COL_CURRENT_PRICE] = current_prices
        df[COL_MARKETING_PRICE] = marketing_prices
        df[COL_CURRENT_MARGIN] = current_margins
        df[COL_EXPECTED_PROFIT] = expected_profits
        df[COL_LOGISTICS] = logistics_used
    
        write_costs_dataframe(df, costs_path)
    
        # Применяем условное форматирование
        try:
            wb = books.enter_context(closing(load_workbook(costs_path)))
            ws = wb["Основной"] if "Основной" in wb.sheetnames else wb.active
        
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
        
            # Отрицательная ожидаемая прибыль — красным
            for col_idx, cell in enumerate(ws[1], start=1):
                if cell.value == COL_EXPECTED_PROFIT:
                    profit_letter = cell.column_letter
                    ws.conditional_formatting.add(
                        f"{profit_letter}2:{profit_letter}{len(df) + 1}",
                        CellIsRule(operator="lessThan", formula=[0], fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")),
                    )
                    break
        
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
        
            save_workbook_atomic(wb, costs_path)
            wb.close()
            print("✅ Применено условное форматирование к колонкам.")
        except Exception as e:
            print(f"⚠️ Не удалось применить условное форматирование: {e}")
    
        print("✅ Текущие цены получены и сохранены.")
        return True


# Колонка бустинга на листе «Акции» стоит рядом с колонкой цены акции.
BOOST_COLUMN_SUFFIX = " — бустинг, %"


def boost_column_name(action_name: str) -> str:
    return f"{action_name}{BOOST_COLUMN_SUFFIX}"


def collect_action_participation(
    offer_ids: List[str],
) -> Tuple[List[Dict[str, Any]], Dict[str, Dict[str, float]], Dict[str, Dict[str, Optional[float]]]]:
    """
    Акции и участие в них товаров: (акции, {артикул: {акция: цена}}, {артикул: {акция: бустинг %}}).
    """
    actions = [action for action in (describe_action(a) for a in get_actions_list()) if action["id"]]
    if not actions:
        return [], {}, {}
    pid_to_offer = get_product_id_mapping(sorted({_normalize_offer_id(oid) for oid in offer_ids if oid}))
    prices: Dict[str, Dict[str, float]] = {}
    boosts: Dict[str, Dict[str, Optional[float]]] = {}
    for action_info in actions:
        for product in fetch_action_products(action_info["id"], "products"):
            offer_id = pid_to_offer.get(product.get("id"))
            price = float(product.get("action_price") or 0)
            if not offer_id or price <= 0:
                continue
            prices.setdefault(offer_id, {})[action_info["name"]] = price
            boosts.setdefault(offer_id, {})[action_info["name"]] = _current_boost(action_info, product, price)
    return actions, prices, boosts


@locked_costs
def action_get_active_actions(repo_root: Path) -> bool:
    """
    Действие 4: Узнать активные акции.
    Просмотр в каких акциях участвует товар и по какой цене.
    """
    with ExitStack() as books:
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
    
        # Получаем акции, цены и бустинг товаров в них
        actions_info_list, prices_by_offer, boosts_by_offer = collect_action_participation(offer_ids_list)
    
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
    
        # Для каждой акции — колонка цены и рядом колонка бустинга
        offers = []
        for _, row in df.iterrows():
            art = _artikul_normalize(row.get(key_col))
            offers.append(_normalize_offer_id(art) if art else "")
        for action_info in actions_info_list:
            action_name = action_info["name"]
            prices = [prices_by_offer.get(offer, {}).get(action_name) for offer in offers]
            actions_df_data[action_name] = [round(price) if price is not None else None for price in prices]
            actions_df_data[boost_column_name(action_name)] = [boosts_by_offer.get(offer, {}).get(action_name) for offer in offers]
    
        actions_df = pd.DataFrame(actions_df_data)
    
        # Сохраняем в Excel
        try:
            wb = books.enter_context(closing(load_workbook(costs_path)))
        
            if 'Sheet1' in wb.sheetnames:
                wb['Sheet1'].title = 'Основной'
        
            if "Акции" in wb.sheetnames:
                wb.remove(wb["Акции"])
        
            ws_actions = wb.create_sheet("Акции")
        
            # Записываем заголовки
            for c_idx, col_name in enumerate(actions_df.columns, start=1):
                ws_actions.cell(row=1, column=c_idx, value=col_name)
        
            # Записываем данные
            boost_col_indices = {
                c_idx for c_idx, col_name in enumerate(actions_df.columns, start=1)
                if str(col_name).endswith(BOOST_COLUMN_SUFFIX)
            }
            for r_idx, row in enumerate(actions_df.itertuples(index=False), start=2):
                for c_idx, value in enumerate(row, start=1):
                    if pd.isna(value):
                        value = None
                    cell = ws_actions.cell(row=r_idx, column=c_idx, value=value)
                    if c_idx in boost_col_indices and value is not None:
                        cell.number_format = '0.#"%"'
        
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
                        if col_name == key_col or col_name == COL_MIN_PRICE or str(col_name).endswith(BOOST_COLUMN_SUFFIX):
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
        
            save_workbook_atomic(wb, costs_path)
            wb.close()
            print(f"✅ Создан лист «Акции»: {len(actions_info_list)} акций, для каждой — цена и бустинг.")
        except Exception as e:
            print(f"⚠️ Ошибка при создании листа «Акции»: {e}")
            import traceback
            print(traceback.format_exc())
            return False
    
        return True


def _load_offer_price_limits(df, key_col: str) -> Dict[str, Tuple[Optional[float], Optional[float]]]:
    """{нормализованный артикул: (минимальная цена, желательная цена)} из costs.xlsx."""
    def _positive(value) -> Optional[float]:
        try:
            number = float(value)
        except (TypeError, ValueError):
            return None
        if pd.isna(number) or number <= 0:
            return None
        return number

    limits: Dict[str, Tuple[Optional[float], Optional[float]]] = {}
    for _, row in df.iterrows():
        art = _artikul_normalize(row.get(key_col))
        if not art:
            continue
        min_price = _positive(row.get(COL_MIN_PRICE)) if COL_MIN_PRICE in df.columns else None
        desired_price = _positive(row.get(COL_DESIRED_PRICE)) if COL_DESIRED_PRICE in df.columns else None
        limits[_normalize_offer_id(art)] = (min_price, desired_price)
    return limits


# Подъём цены в акции меньше этого порога не предлагаем, чтобы не гонять мелкие правки.
RAISE_MIN_STEP_RUB = 10.0
RAISE_MIN_STEP_SHARE = 0.01


def _empty_plan(ok: bool, message: str) -> Dict[str, Any]:
    return {"ok": ok, "message": message, "items": []}


def _load_plan_context(repo_root: Path, need_desired: bool) -> Tuple[Optional[Dict[str, Any]], Optional[Dict[str, Any]]]:
    """
    Общая подготовка для планов: лимиты цен из costs.xlsx, акции и маппинг товаров.
    Возвращает (context, None) или (None, план с ошибкой).
    """
    costs_path = repo_root / COSTS_FILENAME
    if not costs_path.exists():
        return None, _empty_plan(False, f"Файл {COSTS_FILENAME} не найден.")
    df, key_col, _cost_col = load_costs_df(costs_path)
    required = [COL_MIN_PRICE, COL_DESIRED_PRICE] if need_desired else [COL_MIN_PRICE]
    missing = [col for col in required if col not in df.columns]
    if missing:
        names = " и ".join(f"«{col}»" for col in missing)
        return None, _empty_plan(False, f"Не найдены колонки {names}. Сначала рассчитайте оптимальные цены.")
    limits = _load_offer_price_limits(df, key_col)
    if not limits:
        return None, _empty_plan(False, f"В {COSTS_FILENAME} нет артикулов.")
    actions = [describe_action(action) for action in get_actions_list()]
    actions = [action for action in actions if action["id"]]
    if not actions:
        return None, _empty_plan(True, "Активные акции не найдены.")
    return {"limits": limits, "actions": actions, "pid_to_offer": get_product_id_mapping(list(limits))}, None


def _boost_fields(action_info: Dict[str, Any], product: Dict[str, Any]) -> Dict[str, Any]:
    """Диапазон бустинга и цена, при которой достигается максимальный бустинг."""
    if action_info.get("action_type") == ACTION_TYPE_ELASTIC and float(product.get("max_boost") or 0) > 0:
        return {
            "boost_min": float(product.get("min_boost") or 0),
            "boost_max": float(product.get("max_boost") or 0),
            "max_boost_price": float(product.get("price_max_elastic") or 0) or None,
        }
    fixed = action_info.get("fixed_boost")
    return {"boost_min": fixed, "boost_max": fixed, "max_boost_price": None}


def _current_boost(action_info: Dict[str, Any], product: Dict[str, Any], price: float) -> Optional[float]:
    reported = product.get("current_boost")
    if reported:
        return float(reported)
    return boost_at_price(action_info, product, price)


def _moscow_date(value: Optional[str]) -> Optional[str]:
    """Дата из ответа Ozon (UTC, «2026-10-06T21:00:00Z») по московскому времени: «2026-10-07»."""
    if not value:
        return None
    from datetime import datetime, timedelta
    try:
        moment = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return str(value)[:10]
    return (moment + timedelta(hours=3)).date().isoformat()


def _fetch_action_participants(action_id: Any) -> List[Dict[str, Any]]:
    """Все товары-участники акции с повторами: Ozon иногда отвечает дольше 30 секунд."""
    import time
    import requests
    from recommended_prices import fetch_action_page

    products: List[Dict[str, Any]] = []
    last_id = None
    while True:
        for attempt in range(2):
            try:
                batch, next_last_id = fetch_action_page(action_id, "products", last_id, timeout=45)
                break
            except requests.RequestException:
                if attempt == 1:
                    raise
                time.sleep(3)
        products.extend(batch)
        last_id = next_last_id
        # Неполная страница — последняя: лишний запрос по last_id Ozon иногда держит очень долго.
        if len(batch) < 100 or not last_id or last_id == "0":
            return products


def load_actions_calendar(repo_root: Path) -> Dict[str, Any]:
    """
    Календарь акций: даты, тип, бустинг, сколько ваших товаров участвует из возможных,
    даты автодобавления Ozon и список участников с ценой в акции и бустингом.
    """
    from datetime import datetime

    df, key_col, _ = load_costs_df(repo_root / COSTS_FILENAME)
    offers = sorted({_normalize_offer_id(a) for a in (_artikul_normalize(v) for v in df[key_col]) if a})
    pid_to_offer = get_product_id_mapping(offers)
    from concurrent.futures import ThreadPoolExecutor

    raw_actions = [raw for raw in get_actions_list() if describe_action(raw)["id"]]

    def participants_of(raw: Dict[str, Any]) -> Tuple[List[Dict[str, Any]], Optional[str]]:
        try:
            return _fetch_action_participants(describe_action(raw)["id"]), None
        except Exception as exc:
            return [], f"Ozon не отдал список участников: {exc}"

    # Параллельно: одна зависшая акция не задерживает остальные.
    with ThreadPoolExecutor(max_workers=6) as pool:
        loaded = list(pool.map(participants_of, raw_actions))

    actions = []
    for raw, (raw_products, error) in zip(raw_actions, loaded):
        info = describe_action(raw)
        participants = []
        for product in raw_products:
            action_price = float(product.get("action_price") or 0)
            base_price = float(product.get("price") or 0)
            participants.append({
                "offer_id": pid_to_offer.get(product.get("id")) or f"id {product.get('id')}",
                "action_price": action_price,
                "price": base_price,
                "discount": round(1 - action_price / base_price, 3) if base_price and action_price else None,
                "boost": _current_boost(info, product, action_price) if action_price else None,
            })
        participants.sort(key=lambda item: (len(item["offer_id"]), item["offer_id"]))
        actions.append({
            "id": str(info["id"]),
            "name": info["name"],
            "type": info["action_type"],
            "fixed_boost": info["fixed_boost"],
            "start": _moscow_date(raw.get("date_start")),
            "end": _moscow_date(raw.get("date_end")),
            "auto_add_dates": [d for d in (_moscow_date(x) for x in raw.get("auto_add_dates") or []) if d],
            "freeze_date": _moscow_date(raw.get("freeze_date")),
            "can_add": int(raw.get("potential_products_count") or 0),   # сколько товаров ещё можно добавить
            "participating": len(participants) if error is None else int(raw.get("participating_products_count") or 0),
            "error": error,
            "participants": participants,
        })
    actions.sort(key=lambda a: (a["start"] or "", a["name"]))
    return {"loaded": datetime.now().isoformat(timespec="seconds"), "actions": actions}


def _attach_visibility(items: List[Dict[str, Any]]) -> bool:
    """Добавляет к позициям видимость в поиске за 14 дней. True, если Ozon отдал хоть что-то."""
    product_ids = sorted({item["product_id"] for item in items})
    skus = get_product_skus(product_ids)
    visibility = get_search_visibility(sorted(set(skus.values())))
    available = False
    for item in items:
        data = visibility.get(skus.get(item["product_id"]), {})
        item["views"] = data.get("views")
        item["position"] = data.get("position")
        available = available or item["views"] is not None or item["position"] is not None
    return available


def _unify_exit_prices(items: List[Dict[str, Any]]) -> None:
    """У товара одна цена на все акции: каждой позиции удаления — наибольшая цена выхода этого товара."""
    best: Dict[Any, int] = {}
    for item in items:
        if item.get("kind") == "remove" and item.get("exit_price"):
            best[item["product_id"]] = max(item["exit_price"], best.get(item["product_id"], 0))
    for item in items:
        if item.get("kind") == "remove" and item.get("exit_price"):
            item["exit_price"] = best[item["product_id"]]


def _plan_message(items: List[Dict[str, Any]]) -> str:
    labels = {"add": "добавить", "raise": "поднять цену", "remove": "удалить"}
    counts = {kind: sum(1 for item in items if item["kind"] == kind) for kind in labels}
    parts = [f"{labels[kind]}: {count}" for kind, count in counts.items() if count]
    action_count = len({item["action_id"] for item in items})
    return f"{', '.join(parts).capitalize()} — в {action_count} акциях."


def build_action_removal_plan(repo_root: Path) -> Dict[str, Any]:
    """
    Собирает по свежим данным Ozon список товаров, у которых цена в акции
    ниже минимальной цены продажи. Ничего не отправляет.
    """
    context, error_plan = _load_plan_context(repo_root, need_desired=False)
    if error_plan:
        return error_plan
    limits, pid_to_offer = context["limits"], context["pid_to_offer"]

    items: List[Dict[str, Any]] = []
    for action_info in context["actions"]:
        for product in fetch_action_products(action_info["id"], "products"):
            offer_id = pid_to_offer.get(product.get("id"))
            min_price, desired_price = limits.get(offer_id or "", (None, None))
            action_price = float(product.get("action_price") or 0)
            if min_price is None or action_price <= 0 or action_price >= min_price:
                continue
            boost = _current_boost(action_info, product, action_price)
            max_action_price = float(product.get("max_action_price") or 0) or None
            exit_price = action_exit_price(max_action_price, min_price)
            items.append({
                "kind": "remove",
                "action_id": action_info["id"],
                "action_name": action_info["name"],
                "action_type": action_info["action_type"],
                "product_id": product["id"],
                "offer_id": offer_id,
                "action_price": action_price,
                "current_price": action_price,
                "min_price": min_price,
                "desired_price": desired_price,
                "max_action_price": max_action_price,
                "exit_price": exit_price,
                "boost_now": boost,
                "boost_new": None,
                **_boost_fields(action_info, product),
                "selected": exit_price is not None,
            })

    if not items:
        return _empty_plan(True, "Товаров с ценой в акции ниже минимальной не найдено.")
    _unify_exit_prices(items)
    visibility_available = _attach_visibility(items)
    return {"ok": True, "message": _plan_message(items), "items": items, "visibility_available": visibility_available}


def build_action_add_plan(repo_root: Path) -> Dict[str, Any]:
    """
    Собирает план для акций: добавить товары, которые ещё не участвуют,
    и поднять цену участвующим, если её можно приблизить к желательной.
    Цена всегда не ниже минимальной и не выше максимума акции. Ничего не отправляет.
    """
    context, error_plan = _load_plan_context(repo_root, need_desired=True)
    if error_plan:
        return error_plan
    limits, pid_to_offer = context["limits"], context["pid_to_offer"]

    items: List[Dict[str, Any]] = []
    for action_info in context["actions"]:
        base = {
            "action_id": action_info["id"],
            "action_name": action_info["name"],
            "action_type": action_info["action_type"],
        }
        for endpoint in ("candidates", "products"):
            for product in fetch_action_products(action_info["id"], endpoint):
                offer_id = pid_to_offer.get(product.get("id"))
                min_price, desired_price = limits.get(offer_id or "", (None, None))
                if min_price is None or desired_price is None:
                    continue
                try:
                    max_action_price = float(product.get("max_action_price"))
                except (TypeError, ValueError):
                    continue
                price = pick_action_price(min_price, desired_price, max_action_price)
                if price is None:
                    continue
                current_price = float(product.get("action_price") or 0)
                item = {
                    **base,
                    "product_id": product["id"],
                    "offer_id": offer_id,
                    "action_price": price,
                    "min_price": min_price,
                    "desired_price": desired_price,
                    "max_action_price": max_action_price,
                    "stock": int(product.get("stock") or 0),
                    "boost_new": boost_at_price(action_info, product, price),
                    **_boost_fields(action_info, product),
                }
                if current_price <= 0:
                    if endpoint == "products":
                        continue
                    item.update(kind="add", current_price=None, boost_now=None, selected=True)
                else:
                    # Ниже минимальной — задача удаления; поднимаем только заметно.
                    step = max(RAISE_MIN_STEP_RUB, current_price * RAISE_MIN_STEP_SHARE)
                    if current_price < min_price or price - current_price < step:
                        continue
                    boost_now = _current_boost(action_info, product, current_price)
                    boost_new = item["boost_new"]
                    # Подъём, который снижает бустинг, по умолчанию не отмечаем.
                    keeps_boost = boost_now is None or boost_new is None or boost_new >= boost_now
                    item.update(kind="raise", current_price=current_price, boost_now=boost_now, selected=keeps_boost)
                items.append(item)

    if not items:
        return _empty_plan(True, "Не найдено товаров, которые можно выгодно добавить в акции или поднять в цене.")
    visibility_available = _attach_visibility(items)
    return {"ok": True, "message": _plan_message(items), "items": items, "visibility_available": visibility_available}


def execute_action_plan(plan: Dict[str, Any], mode: str) -> Dict[str, Any]:
    """
    Отправляет план в Ozon (mode: "remove" или "add" — добавление и подъём цены)
    и перепроверяет по факту: позиция считается выполненной, только если Ozon
    после отправки показывает новое состояние товара в акции.
    """
    items = [item for item in plan.get("items") or [] if item.get("selected", True)]
    if mode == "remove":
        _unify_exit_prices(items)
    by_action: Dict[Any, List[Dict[str, Any]]] = {}
    for item in items:
        by_action.setdefault(item["action_id"], []).append(item)

    confirmed = 0
    rejected: List[Dict[str, Any]] = []
    unconfirmed: List[Dict[str, Any]] = []
    errors: List[str] = []
    for action_id, action_items in by_action.items():
        action_name = action_items[0]["action_name"]
        by_product = {item["product_id"]: item for item in action_items}
        if mode == "remove":
            # Единая цена Ozon: товар выходит из бустинга, когда его цена выше лимита акции.
            without_price = [item for item in action_items if not item.get("exit_price")]
            for item in without_price:
                errors.append(f"{action_name}: {item['offer_id']} — Ozon не сообщил лимит акции, цену выхода посчитать нельзя.")
            result = activate_products_in_action(
                action_id,
                [{"product_id": item["product_id"], "action_price": item["exit_price"]} for item in action_items if item.get("exit_price")],
            )
        else:
            result = activate_products_in_action(
                action_id,
                [
                    {"product_id": item["product_id"], "action_price": item["action_price"], "stock": item["stock"]}
                    for item in action_items
                ],
            )
        rejected_ids = set()
        for entry in result.get("rejected") or []:
            product_id = entry.get("product_id")
            rejected_ids.add(product_id)
            rejected.append({
                "action_name": action_name,
                "offer_id": by_product.get(product_id, {}).get("offer_id", str(product_id)),
                "reason": entry.get("reason") or "причина не указана",
            })
        errors.extend(f"{action_name}: {error}" for error in result.get("errors") or [])

        actual = {product["id"]: product for product in fetch_action_products(action_id, "products")}
        for product_id, item in by_product.items():
            if product_id in rejected_ids:
                continue
            product = actual.get(product_id)
            if mode == "remove":
                applied = product is None
            else:
                applied = product is not None and abs(float(product.get("action_price") or 0) - item["action_price"]) < 0.5
            if applied:
                confirmed += 1
            else:
                unconfirmed.append({"action_name": action_name, "offer_id": item["offer_id"]})

    verb = "Удалено из акций" if mode == "remove" else "Применено в акциях"
    message = f"{verb}: {confirmed} из {len(items)} (проверено по данным Ozon)."
    if rejected:
        message += f" Ozon отклонил: {len(rejected)}."
    if unconfirmed:
        message += f" Не применилось без объяснения: {len(unconfirmed)}."
    if errors:
        message += f" Ошибок запросов: {len(errors)}."
    return {
        "ok": confirmed == len(items) and not errors,
        "done": confirmed,
        "expected": len(items),
        "rejected": rejected,
        "unconfirmed": unconfirmed,
        "errors": errors,
        "message": message,
    }


def _ru_date(value: str) -> str:
    return f"{value[8:10]}.{value[5:7]}"


def upcoming_action_scope(raw_action: Dict[str, Any], today: str) -> Optional[Dict[str, Any]]:
    """
    Акция «предстоящая», если она ещё не началась или у неё впереди дата автодобавления Ozon.
    Возвращает {"start", "auto_add": [даты ≥ сегодня]} или None. Даты — «YYYY-MM-DD» по Москве.
    """
    start = _moscow_date(raw_action.get("date_start"))
    auto_add = sorted(d for d in (_moscow_date(x) for x in raw_action.get("auto_add_dates") or []) if d and d >= today)
    if (start and start > today) or auto_add:
        return {"start": start, "auto_add": auto_add}
    return None


def plan_upcoming_action_items(
    action_info: Dict[str, Any],
    scope: Dict[str, Any],
    participants: List[Dict[str, Any]],
    candidates: List[Dict[str, Any]],
    limits: Dict[str, Tuple[Optional[float], Optional[float]]],
    pid_to_offer: Dict[Any, str],
) -> List[Dict[str, Any]]:
    """
    Порядок в одной предстоящей акции (ничего не отправляет):
    • участник с ценой ниже минимальной — поднять цену до приемлемой, если акция позволяет, иначе удалить;
    • кандидат, для которого есть цена не ниже минимальной, — добавить сейчас по цене ближе к желаемой
      (раньше, чем Ozon добавит его сам по своей цене);
    • кандидат, у которого максимум акции ниже минимальной цены, — предупреждение: Ozon может добавить его
      сам, а отключить автодобавление через API нельзя.
    """
    base = {"action_id": action_info["id"], "action_name": action_info["name"], "action_type": action_info["action_type"]}
    auto_text = ", ".join(_ru_date(d) for d in scope["auto_add"])
    starts_text = f"акция начнётся {_ru_date(scope['start'])}" if scope.get("start") and not scope["auto_add"] else ""
    items: List[Dict[str, Any]] = []

    def price_of(product: Dict[str, Any], key: str) -> Optional[float]:
        try:
            value = float(product.get(key))
        except (TypeError, ValueError):
            return None
        return value if value > 0 else None

    for product in participants:
        offer_id = pid_to_offer.get(product.get("id"))
        min_price, desired_price = limits.get(offer_id or "", (None, None))
        current = price_of(product, "action_price")
        if min_price is None or current is None or current >= min_price:
            continue
        auto = product.get("add_mode") == "AUTO"
        who = "Ozon добавил товар сам" if auto else "Товар в акции"
        max_price = price_of(product, "max_action_price")
        target = pick_action_price(min_price, desired_price or min_price, max_price) if max_price else None
        item = {
            **base,
            "product_id": product["id"],
            "offer_id": offer_id,
            "min_price": min_price,
            "desired_price": desired_price,
            "max_action_price": max_price,
            "stock": int(product.get("stock") or 0),
            "current_price": current,
            "boost_now": _current_boost(action_info, product, current),
            "added_by_ozon": auto,
            "selected": True,
            **_boost_fields(action_info, product),
        }
        if target is not None and target > current:
            item.update(kind="raise", action_price=target, boost_new=boost_at_price(action_info, product, target),
                        note=f"{who} по цене ниже минимальной — поднимаем до приемлемой.")
        else:
            exit_price = action_exit_price(max_price, min_price)
            item.update(kind="remove", action_price=current, boost_new=None, exit_price=exit_price, selected=exit_price is not None,
                        note=f"{who}, а максимальная цена акции ниже вашей минимальной — выводим из акции.")
        items.append(item)

    for product in candidates:
        offer_id = pid_to_offer.get(product.get("id"))
        min_price, desired_price = limits.get(offer_id or "", (None, None))
        max_price = price_of(product, "max_action_price")
        if min_price is None or max_price is None:
            continue
        target = pick_action_price(min_price, desired_price or min_price, max_price)
        item = {
            **base,
            "product_id": product["id"],
            "offer_id": offer_id,
            "min_price": min_price,
            "desired_price": desired_price,
            "max_action_price": max_price,
            "stock": int(product.get("stock") or 0),
            "current_price": None,
            "boost_now": None,
            "added_by_ozon": False,
            **_boost_fields(action_info, product),
        }
        if target is not None:
            when = f"Ozon может добавить сам {auto_text} — добавляем заранее по своей цене." if auto_text else (starts_text.capitalize() + ".")
            item.update(kind="add", action_price=target, boost_new=boost_at_price(action_info, product, target), selected=True, note=when)
            items.append(item)
        elif auto_text:
            gap = f"{min_price - max_price:,.0f}".replace(",", " ")
            item.update(kind="watch", action_price=max_price, boost_new=None, selected=False,
                        note=f"Ozon может добавить сам {auto_text}: цена в акции будет на {gap} ₽ ниже минимальной.")
            items.append(item)
    return items


def build_upcoming_actions_plan(repo_root: Path, today: Optional[str] = None) -> Dict[str, Any]:
    """
    План «порядок в предстоящих акциях»: акции, которые ещё не начались или куда Ozon скоро
    добавит товары сам. Ничего не отправляет.
    """
    from datetime import datetime, timedelta, timezone
    from concurrent.futures import ThreadPoolExecutor

    today = today or (datetime.now(timezone.utc) + timedelta(hours=3)).date().isoformat()
    context, error_plan = _load_plan_context(repo_root, need_desired=False)
    if error_plan:
        return error_plan
    limits, pid_to_offer = context["limits"], context["pid_to_offer"]
    raw_by_id = {str(raw.get("id")): raw for raw in get_actions_list()}
    targets = []
    for action_info in context["actions"]:
        scope = upcoming_action_scope(raw_by_id.get(str(action_info["id"])) or {}, today)
        if scope:
            targets.append((action_info, scope))
    if not targets:
        return _empty_plan(True, "Предстоящих акций и запланированных автодобавлений Ozon нет.")

    def load(target):
        action_info, scope = target
        return (_fetch_action_participants(action_info["id"]), fetch_action_products(action_info["id"], "candidates"))

    with ThreadPoolExecutor(max_workers=4) as pool:
        loaded = list(pool.map(load, targets))

    items: List[Dict[str, Any]] = []
    for (action_info, scope), (participants, candidates) in zip(targets, loaded):
        items.extend(plan_upcoming_action_items(action_info, scope, participants, candidates, limits, pid_to_offer))

    _unify_exit_prices(items)
    names = ", ".join(f"«{info['name']}»" for info, _ in targets)
    if not items:
        return _empty_plan(True, f"Проверены предстоящие акции ({names}): всё в порядке, менять нечего.")
    labels = {"add": "добавить", "raise": "поднять цену", "remove": "удалить", "watch": "под присмотром"}
    counts = [f"{label}: {sum(1 for item in items if item['kind'] == kind)}" for kind, label in labels.items()
              if any(item["kind"] == kind for item in items)]
    message = f"Предстоящие акции ({names}). " + ", ".join(counts).capitalize() + "."
    visibility_available = _attach_visibility(items)
    return {"ok": True, "message": message, "items": items, "visibility_available": visibility_available}


def execute_upcoming_actions_plan(plan: Dict[str, Any]) -> Dict[str, Any]:
    """Отправляет отмеченные позиции плана предстоящих акций: удаления и добавления/подъёмы отдельно."""
    selected = [item for item in plan.get("items") or [] if item.get("selected") and item["kind"] in ("add", "raise", "remove")]
    results = []
    removals = [item for item in selected if item["kind"] == "remove"]
    additions = [item for item in selected if item["kind"] in ("add", "raise")]
    if removals:
        results.append(execute_action_plan({"items": removals}, "remove"))
    if additions:
        results.append(execute_action_plan({"items": additions}, "add"))
    done = sum(r["done"] for r in results)
    merged = {
        "done": done,
        "expected": len(selected),
        "rejected": [x for r in results for x in r["rejected"]],
        "unconfirmed": [x for r in results for x in r["unconfirmed"]],
        "errors": [x for r in results for x in r["errors"]],
    }
    message = f"Применено в предстоящих акциях: {done} из {len(selected)} (проверено по данным Ozon)."
    if merged["rejected"]:
        message += f" Ozon отклонил: {len(merged['rejected'])}."
    if merged["unconfirmed"]:
        message += f" Не применилось без объяснения: {len(merged['unconfirmed'])}."
    if merged["errors"]:
        message += f" Ошибок запросов: {len(merged['errors'])}."
    merged.update(ok=done == len(selected) and not merged["errors"], message=message)
    return merged


def _format_boost(value: Optional[float]) -> str:
    return "—" if value is None else f"{value:g}%"


def _run_action_plan(plan: Dict[str, Any], mode: str, question: str) -> bool:
    """Печатает план, спрашивает подтверждение, отправляет и печатает итог."""
    message = str(plan.get("message") or "")
    if not plan.get("ok"):
        print(f"❌ {message}")
        return False
    print(message)
    if not plan["items"]:
        return True
    labels = {"add": "добавить", "raise": "поднять", "remove": "удалить"}
    for item in plan["items"]:
        if item["kind"] == "remove":
            prices = f"в акции {item['action_price']:.0f} руб. < мин. {item['min_price']:.0f} руб."
            boost = f"буст {_format_boost(item['boost_now'])} будет потерян"
        elif item["kind"] == "raise":
            prices = f"{item['current_price']:.0f} -> {item['action_price']} руб."
            boost = f"буст {_format_boost(item['boost_now'])} -> {_format_boost(item['boost_new'])}"
        else:
            prices = f"{item['action_price']} руб. (мин. {item['min_price']:.0f} руб.)"
            boost = f"буст {_format_boost(item['boost_new'])}"
        mark = "" if item.get("selected", True) else " (не отмечено: снижает буст)"
        print(f"   [{labels[item['kind']]}{mark}] {item['offer_id']} | {item['action_name']} | {prices} | {boost}")
    if not prompt_yes_no(question, default_yes=False):
        print("❌ Отменено, в Ozon ничего не отправлено.")
        return False
    result = execute_action_plan(plan, mode)
    print(("✅ " if result["ok"] else "⚠️ ") + result["message"])
    for entry in result["rejected"]:
        print(f"   ✖ {entry['offer_id']} | {entry['action_name']}: {entry['reason']}")
    for entry in result["unconfirmed"]:
        print(f"   ✖ {entry['offer_id']} | {entry['action_name']}: Ozon не применил изменение")
    for error in result["errors"]:
        print(f"   ✖ {error}")
    return result["ok"]


def action_remove_unprofitable_actions(repo_root: Path) -> bool:
    """
    Действие 5: Удалить невыгодные акции.
    Удаление товаров из акций, где цена в акции меньше минимальной.
    """
    print_step("Удалить невыгодные акции")
    return _run_action_plan(build_action_removal_plan(repo_root), "remove", "Удалить эти товары из акций?")


def action_add_to_actions(repo_root: Path) -> bool:
    """
    Действие 6: Добавить товары в акции и поднять цены участвующим.
    Цена в диапазоне от минимальной до желательной, не выше максимума акции.
    """
    print_step("Добавить товары в акции")
    return _run_action_plan(build_action_add_plan(repo_root), "add", "Применить эти изменения в акциях?")


@locked_costs
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


def build_discount_request_plan(repo_root: Path) -> Dict[str, Any]:
    costs_path = repo_root / COSTS_FILENAME
    if not costs_path.exists():
        return {"ok": False, "message": f"Файл {COSTS_FILENAME} не найден.", "items": [], "approve_tasks": [], "decline_tasks": []}

    df, key_col, _cost_col = load_costs_df(costs_path)
    if COL_MIN_PRICE not in df.columns:
        return {
            "ok": False,
            "message": f"Колонка «{COL_MIN_PRICE}» не найдена в {COSTS_FILENAME}. Сначала рассчитайте минимальные цены.",
            "items": [],
            "approve_tasks": [],
            "decline_tasks": [],
        }

    offer_id_to_min_price: Dict[str, float] = {}
    for _, row in df.iterrows():
        art = _artikul_normalize(row.get(key_col))
        if not art:
            continue
        art_normalized = _normalize_offer_id(art)
        min_price = row.get(COL_MIN_PRICE)
        try:
            min_price_val = float(min_price) if min_price is not None else 0.0
        except (TypeError, ValueError):
            min_price_val = 0.0
        if min_price_val <= 0:
            continue
        offer_id_to_min_price[art_normalized] = min_price_val
        if art != art_normalized:
            offer_id_to_min_price[art] = min_price_val

    if not offer_id_to_min_price:
        return {"ok": False, "message": "Не найдено товаров с минимальной ценой.", "items": [], "approve_tasks": [], "decline_tasks": []}

    discount_tasks = get_discount_requests(status="NEW", limit=50)
    if not discount_tasks:
        return {"ok": True, "message": "Новых заявок на скидку не найдено.", "items": [], "approve_tasks": [], "decline_tasks": []}

    skus_from_tasks: List[int] = []
    for task in discount_tasks:
        sku = task.get("sku")
        try:
            if sku is not None:
                skus_from_tasks.append(int(sku))
        except (TypeError, ValueError):
            continue
    sku_to_offer_id: Dict[int, str] = get_sku_to_offer_id_mapping(list(dict.fromkeys(skus_from_tasks))) if skus_from_tasks else {}

    items: List[Dict[str, Any]] = []
    approve_tasks: List[Dict[str, Any]] = []
    decline_tasks: List[Dict[str, Any]] = []

    for task in discount_tasks:
        task_id = task.get("id")
        sku = task.get("sku")
        requested_price = task.get("requested_price")
        requested_quantity_min = task.get("requested_quantity_min")
        requested_quantity_max = task.get("requested_quantity_max")

        offer_id_raw: Optional[str] = None
        try:
            if sku is not None:
                offer_id_raw = sku_to_offer_id.get(int(sku))
        except (TypeError, ValueError):
            offer_id_raw = None
        if not offer_id_raw and sku is not None:
            offer_id_raw = str(sku)

        offer_id_normalized = _normalize_offer_id(offer_id_raw) if offer_id_raw else ""
        min_price = offer_id_to_min_price.get(offer_id_normalized) or offer_id_to_min_price.get(offer_id_raw or "")

        try:
            requested_price_val = float(requested_price) if requested_price is not None else None
        except (TypeError, ValueError):
            requested_price_val = None

        try:
            q_min = max(1, int(requested_quantity_min)) if requested_quantity_min is not None else 1
        except (TypeError, ValueError):
            q_min = 1
        try:
            q_max = max(q_min, int(requested_quantity_max)) if requested_quantity_max is not None else q_min
        except (TypeError, ValueError):
            q_max = q_min

        decision = "decline"
        reason = ""
        approved_task: Optional[Dict[str, Any]] = None
        declined_task: Optional[Dict[str, Any]] = None
        price_state = "unknown"
        price_delta: Optional[float] = None

        if requested_price_val is not None and min_price is not None:
            price_delta = requested_price_val - min_price
            price_state = "ok" if price_delta >= 0 else "below_min"

        def build_approve_payload(comment: str) -> Dict[str, Any]:
            return {
                "id": task_id,
                "approved_price": int(requested_price_val),
                "approved_quantity_min": q_min,
                "approved_quantity_max": q_max,
                "seller_comment": comment,
            }

        if not sku:
            reason = "SKU отсутствует в заявке."
        elif min_price is None:
            reason = "Минимальная цена не рассчитана для этого товара."
            if requested_price_val is not None:
                approved_task = build_approve_payload(
                    f"Одобрено вручную: товара нет в {COSTS_FILENAME}, минимальная цена не найдена."
                )
        elif requested_price_val is None:
            reason = "В заявке не указана корректная запрошенная цена."
        elif requested_price_val >= min_price:
            decision = "approve"
            reason = f"Запрошенная цена {requested_price_val:.2f} ₽ не ниже минимальной цены {min_price:.2f} ₽."
            approved_task = build_approve_payload(
                f"Одобрено: запрошенная цена {requested_price_val:.0f} ₽ не ниже минимальной {min_price:.0f} ₽"
            )
            approve_tasks.append(approved_task)
        else:
            reason = f"Запрошенная цена {requested_price_val:.2f} ₽ ниже минимальной цены {min_price:.2f} ₽."
            approved_task = build_approve_payload(
                f"Одобрено вручную: запрошенная цена {requested_price_val:.0f} ₽ ниже минимальной {min_price:.0f} ₽"
            )

        if decision != "approve":
            declined_task = {
                "id": task_id,
                "seller_comment": f"Отклонено: {reason}",
            }
            decline_tasks.append(declined_task)

        items.append(
            {
                "id": task_id,
                "sku": str(sku or "—"),
                "offer_id": offer_id_raw or "—",
                "requested_price": requested_price_val,
                "requested_quantity_min": q_min,
                "requested_quantity_max": q_max,
                "min_price": min_price,
                "price_state": price_state,
                "price_delta": price_delta,
                "recommended_action": decision,
                "recommendation_reason": reason,
                "approve_payload": approved_task,
                "decline_payload": declined_task,
            }
        )

    return {
        "ok": True,
        "message": f"Найдено {len(items)} новых заявок на скидку.",
        "items": items,
        "approve_tasks": approve_tasks,
        "decline_tasks": decline_tasks,
    }


def process_discount_request_item(item: Dict[str, Any], action: str) -> Dict[str, Any]:
    if action == "approve":
        payload = item.get("approve_payload")
        if not payload:
            return {"ok": False, "message": "Эту заявку нельзя одобрить автоматически: нет корректного payload."}
        result = approve_discount_requests([payload])
        ok = result.get("success_count", 0) > 0 and result.get("fail_count", 0) == 0
        return {"ok": ok, "message": "Заявка одобрена." if ok else "Не удалось одобрить заявку.", "result": result}

    if action == "decline":
        payload = item.get("decline_payload")
        if not payload:
            return {"ok": False, "message": "Не найден payload для отклонения заявки."}
        result = decline_discount_requests([payload])
        ok = result.get("success_count", 0) > 0 and result.get("fail_count", 0) == 0
        return {"ok": ok, "message": "Заявка отклонена." if ok else "Не удалось отклонить заявку.", "result": result}

    return {"ok": False, "message": f"Неизвестное действие: {action}"}


def process_discount_request_plan(plan: Dict[str, Any]) -> Dict[str, Any]:
    approve_result = {"success_count": 0, "fail_count": 0, "fail_details": []}
    decline_result = {"success_count": 0, "fail_count": 0, "fail_details": []}

    approve_tasks = plan.get("approve_tasks") or []
    decline_tasks = plan.get("decline_tasks") or []
    if approve_tasks:
        approve_result = approve_discount_requests(approve_tasks)
    if decline_tasks:
        decline_result = decline_discount_requests(decline_tasks)

    expected_count = len(approve_tasks) + len(decline_tasks)
    success_count = approve_result.get("success_count", 0) + decline_result.get("success_count", 0)
    fail_count = approve_result.get("fail_count", 0) + decline_result.get("fail_count", 0)
    ok = fail_count == 0 and success_count == expected_count
    message = f"Автоматически обработано: одобрено {approve_result.get('success_count', 0)}, отклонено {decline_result.get('success_count', 0)}."
    if not ok:
        message += " Не все заявки обработаны. Обновите список перед повторной отправкой."
    return {
        "ok": ok,
        "approve_result": approve_result,
        "decline_result": decline_result,
        "message": message,
    }


def show_pricing_menu(repo_root: Path) -> None:
    """Показывает меню управления ценами и скидками."""
    while True:
        print_step("Цены и скидки")
        print("1. Диапазон рентабельности")
        print("2. Рассчитать оптимальную цену")
        print("3. Обработать заявки на скидку")
        print("0. Назад")

        choice = input("Выберите опцию (0-3): ").strip()

        if choice == "1":
            action_set_margin_range(repo_root)
        elif choice == "2":
            action_calculate_optimal_prices(repo_root)
        elif choice == "3":
            action_process_discount_requests(repo_root)
        elif choice == "0":
            break
        else:
            print("Пожалуйста, выберите корректную опцию (0-3).")

        print()


def show_actions_menu(repo_root: Path) -> None:
    """Показывает меню управления акциями."""
    while True:
        print_step("Акции")
        print("1. Узнать текущие цены и акции")
        print("2. Удалить невыгодные акции")
        print("3. Добавить товары в акции")
        print("0. Назад")

        choice = input("Выберите опцию (0-3): ").strip()

        if choice == "1":
            action_get_current_prices_and_actions(repo_root)
        elif choice == "2":
            if action_remove_unprofitable_actions(repo_root):
                refresh_costs_views(repo_root)
        elif choice == "3":
            if action_add_to_actions(repo_root):
                refresh_costs_views(repo_root)
        elif choice == "0":
            break
        else:
            print("Пожалуйста, выберите корректную опцию (0-3).")

        print()


def show_price_management_menu(repo_root: Path) -> None:
    """Показывает корневое меню разделов цен и акций."""
    while True:
        print_step("Цены и акции")
        print("1. Цены и скидки")
        print("2. Акции")
        print("0. Выход")

        choice = input("Выберите раздел (0-2): ").strip()

        if choice == "1":
            show_pricing_menu(repo_root)
        elif choice == "2":
            show_actions_menu(repo_root)
        elif choice == "0":
            break
        else:
            print("Пожалуйста, выберите корректную опцию (0-2).")

        print()


def main(argv: Optional[List[str]] = None) -> int:
    """Точка входа для запуска модуля как скрипта."""
    parser = argparse.ArgumentParser(description="Меню разделов цен и акций.")
    parser.add_argument(
        "--menu",
        choices=("all", "pricing", "actions"),
        default="all",
        help="Сразу открыть конкретный раздел.",
    )
    args = parser.parse_args(argv)

    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parent

    if args.menu == "pricing":
        show_pricing_menu(repo_root)
    elif args.menu == "actions":
        show_actions_menu(repo_root)
    else:
        show_price_management_menu(repo_root)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
