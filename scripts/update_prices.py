# -*- coding: utf-8 -*-
"""
Обновление минимальных цен товаров на Ozon.
Читает минимальные цены из costs.xlsx и обновляет их на Ozon через API v1/product/import/prices.
"""

import os
import sys
from pathlib import Path
from typing import List, Dict, Any, Optional

import pandas as pd
import requests
from dotenv import load_dotenv

# Загружаем переменные окружения для API Ozon
load_dotenv()
OZON_CLIENT_ID = os.getenv('OZON_CLIENT_ID')
OZON_API_KEY = os.getenv('OZON_API_KEY')

OZON_HEADERS = {
    'Client-Id': OZON_CLIENT_ID or '',
    'Api-Key': OZON_API_KEY or '',
    'Content-Type': 'application/json'
}

COSTS_FILENAME = "costs.xlsx"
COL_MIN_PRICE = "Минимальная цена продажи"


def _artikul_normalize(v):
    """Нормализация артикула для сопоставления (строка, без пробелов по краям)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def _normalize_offer_id(offer_id: str) -> str:
    """
    Преобразует артикул в строку без .0 (например, '1101.0' -> '1101').
    """
    try:
        return str(int(float(offer_id)))
    except (ValueError, TypeError):
        return str(offer_id).strip()


def load_costs_df(costs_path: Path) -> tuple[pd.DataFrame, str]:
    """Читает costs.xlsx, возвращает DataFrame и имя колонки с артикулами."""
    if not costs_path.exists():
        raise FileNotFoundError(f"Файл себестоимости не найден: {costs_path}")

    df = pd.read_excel(costs_path)
    lower_cols = {c.lower(): c for c in df.columns}
    key_col = None
    for v in ["prefix", "префикс", "код", "артикул", "offer_id"]:
        if v in lower_cols:
            key_col = lower_cols[v]
            break
    
    if not key_col:
        raise ValueError(
            "В costs.xlsx не найден столбец артикула. "
            "Ожидаются: 'артикул' (или prefix/код/offer_id)."
        )
    
    if COL_MIN_PRICE not in df.columns:
        raise ValueError(f"В costs.xlsx не найдена колонка «{COL_MIN_PRICE}». Сначала запустите расчёт рекомендуемых цен.")
    
    return df, key_col


def get_current_prices_from_ozon(offer_ids: List[str]) -> Dict[str, Optional[float]]:
    """
    Получает текущие цены продажи товаров с Ozon по их offer_id (артикулам).
    Использует API v5/product/info/prices.
    Возвращает словарь {offer_id: цена}.
    """
    if not OZON_CLIENT_ID or not OZON_API_KEY:
        return {}
    
    if not offer_ids:
        return {}
    
    # Нормализуем артикулы
    normalized_offer_ids = []
    for oid in offer_ids:
        normalized = _normalize_offer_id(oid)
        if normalized:
            normalized_offer_ids.append(normalized)
    
    if not normalized_offer_ids:
        return {}
    
    url = "https://api-seller.ozon.ru/v5/product/info/prices"
    prices_map = {}
    batch_size = 100
    
    # Разбиваем на батчи по 100 артикулов
    for i in range(0, len(normalized_offer_ids), batch_size):
        batch = normalized_offer_ids[i:i + batch_size]
        
        cursor = ""
        has_more = True
        
        while has_more:
            payload = {
                "cursor": cursor,
                "filter": {
                    "offer_id": batch,
                    "visibility": "ALL"
                },
                "limit": 100
            }
            
            try:
                response = requests.post(url, headers=OZON_HEADERS, json=payload, timeout=30)
                response.raise_for_status()
                data = response.json()
                
                items = data.get("items", [])
                cursor = data.get("cursor", "")
                
                if len(items) == 0:
                    break
                
                for item in items:
                    offer_id_raw = item.get("offer_id", "")
                    offer_id_normalized = _normalize_offer_id(str(offer_id_raw)) if offer_id_raw else None
                    
                    price = None
                    price_obj = item.get("price", {})
                    
                    if isinstance(price_obj, dict):
                        if "price" in price_obj:
                            try:
                                price = float(price_obj["price"])
                            except (TypeError, ValueError):
                                pass
                        
                        if price is None and "old_price" in price_obj:
                            try:
                                price = float(price_obj["old_price"])
                            except (TypeError, ValueError):
                                pass
                    else:
                        try:
                            price = float(price_obj)
                        except (TypeError, ValueError):
                            pass
                    
                    if offer_id_normalized and price is not None:
                        # Сохраняем по нормализованному ключу
                        prices_map[offer_id_normalized] = price
                        # Также сохраняем по исходным артикулам для сопоставления
                        for orig_oid in offer_ids:
                            if _normalize_offer_id(orig_oid) == offer_id_normalized:
                                prices_map[orig_oid] = price
                                break
                
                has_more = bool(cursor) and len(items) >= 100
                    
            except requests.exceptions.RequestException as e:
                print(f"⚠️ Ошибка при запросе цен для батча: {e}")
                break
    
    return prices_map


def update_min_prices_on_ozon(updates: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Обновляет минимальные цены товаров на Ozon через API v1/product/import/prices.
    
    Args:
        updates: Список словарей с данными для обновления:
            [{"offer_id": "1101", "min_price": "1500"}, ...]
    
    Returns:
        Результат обновления от API.
    """
    if not OZON_CLIENT_ID or not OZON_API_KEY:
        raise RuntimeError("OZON_CLIENT_ID или OZON_API_KEY не настроены. Укажите их в .env файле.")
    
    if not updates:
        return {"result": {"task_id": None}, "errors": []}
    
    url = "https://api-seller.ozon.ru/v1/product/import/prices"
    
    payload = {
        "prices": updates
    }
    
    try:
        response = requests.post(url, headers=OZON_HEADERS, json=payload, timeout=30)
        response.raise_for_status()
        data = response.json()
        # API может вернуть результат в разных форматах
        # Если это словарь - возвращаем как есть, если список - оборачиваем
        if isinstance(data, list):
            return {"result": data, "errors": []}
        return data
    except requests.exceptions.RequestException as e:
        if hasattr(e, 'response') and e.response is not None:
            try:
                error_data = e.response.json()
                raise RuntimeError(f"Ошибка API Ozon: {error_data}")
            except:
                raise RuntimeError(f"Ошибка API Ozon: {e.response.text}")
        raise RuntimeError(f"Ошибка при обновлении цен: {e}")


def run(repo_root: Path) -> None:
    """Основная функция обновления цен."""
    costs_path = repo_root / COSTS_FILENAME
    
    print("📊 Загрузка данных из costs.xlsx...")
    df, key_col = load_costs_df(costs_path)
    
    # Фильтруем только строки с заполненными артикулами и минимальными ценами
    updates = []
    skipped = []
    
    for _, row in df.iterrows():
        art = _artikul_normalize(row.get(key_col))
        min_price = row.get(COL_MIN_PRICE)
        
        if not art:
            skipped.append("пустой артикул")
            continue
        
        if pd.isna(min_price) or min_price is None:
            skipped.append(f"артикул {art} - нет минимальной цены")
            continue
        
        try:
            min_price_val = float(min_price)
            if min_price_val <= 0:
                skipped.append(f"артикул {art} - некорректная цена {min_price_val}")
                continue
        except (TypeError, ValueError):
            skipped.append(f"артикул {art} - некорректная цена {min_price}")
            continue
        
        # Нормализуем артикул
        offer_id = _normalize_offer_id(art)
        
        # Формируем запрос на обновление (только min_price)
        updates.append({
            "offer_id": offer_id,
            "min_price": str(int(min_price_val))  # API ожидает строку
        })
    
    if not updates:
        print("⚠️ Нет товаров для обновления цен.")
        if skipped:
            print(f"   Пропущено: {len(skipped)} записей")
        return
    
    print(f"✅ Найдено {len(updates)} товаров для обновления минимальной цены.")
    if skipped:
        print(f"⚠️ Пропущено {len(skipped)} записей (пустые артикулы или цены)")
    
    # Получаем текущие цены с Ozon для проверки
    print("\n📡 Получение текущих цен продажи с Ozon для проверки...")
    offer_ids_for_check = [item["offer_id"] for item in updates]
    current_prices = get_current_prices_from_ozon(offer_ids_for_check)
    
    if current_prices:
        print(f"✅ Получено цен для {len(current_prices)} товаров.")
    else:
        print("⚠️ Не удалось получить цены с Ozon (возможно, не настроены API ключи).")
        print("   Продолжаем обновление без проверки...")
    
    # Проверяем, что минимальная цена не больше текущей цены продажи
    validated_updates = []
    price_warnings = []
    
    for update_item in updates:
        offer_id = update_item["offer_id"]
        min_price_val = float(update_item["min_price"])
        
        # Проверяем текущую цену
        current_price = current_prices.get(offer_id) or current_prices.get(_normalize_offer_id(offer_id))
        
        if current_price is not None:
            if min_price_val > current_price:
                price_warnings.append({
                    "offer_id": offer_id,
                    "min_price": min_price_val,
                    "current_price": current_price
                })
                # Пропускаем товары, где минимальная цена больше текущей
                continue
        
        validated_updates.append(update_item)
    
    if price_warnings:
        print(f"\n⚠️ Обнаружено {len(price_warnings)} товаров, где минимальная цена больше текущей цены продажи:")
        print("   (эти товары будут пропущены)")
        for warn in price_warnings[:10]:  # Показываем первые 10
            print(f"   - Артикул {warn['offer_id']}: мин.цена {warn['min_price']:.0f} > текущая {warn['current_price']:.0f}")
        if len(price_warnings) > 10:
            print(f"   ... и ещё {len(price_warnings) - 10} товаров")
        print("   💡 Решение: сначала обновите цену продажи, затем минимальную цену")
    
    if not validated_updates:
        print("\n❌ Нет товаров для обновления после проверки цен.")
        return
    
    print(f"\n✅ После проверки осталось {len(validated_updates)} товаров для обновления.")
    updates = validated_updates
    
    # Разбиваем на батчи по 1000 товаров (лимит API)
    batch_size = 1000
    total_updated = 0
    total_errors = 0
    
    for i in range(0, len(updates), batch_size):
        batch = updates[i:i + batch_size]
        batch_num = i // batch_size + 1
        total_batches = (len(updates) + batch_size - 1) // batch_size
        
        print(f"\n📤 Отправка батча {batch_num}/{total_batches} ({len(batch)} товаров)...")
        
        try:
            result = update_min_prices_on_ozon(batch)
            
            # API возвращает {"result": [список результатов по каждому товару]}
            if isinstance(result, dict) and "result" in result:
                results_list = result["result"]
                
                if isinstance(results_list, list):
                    batch_updated = 0
                    batch_errors = []
                    
                    for item_result in results_list:
                        if isinstance(item_result, dict):
                            offer_id = item_result.get("offer_id", "unknown")
                            updated = item_result.get("updated", False)
                            errors = item_result.get("errors", [])
                            
                            if updated:
                                batch_updated += 1
                            elif errors:
                                # Собираем ошибки для вывода
                                for err in errors:
                                    if isinstance(err, dict):
                                        err_msg = err.get("message", str(err))
                                        err_code = err.get("code", "")
                                        batch_errors.append(f"{offer_id}: {err_code} - {err_msg}")
                                    else:
                                        batch_errors.append(f"{offer_id}: {err}")
                    
                    total_updated += batch_updated
                    total_errors += len(results_list) - batch_updated
                    
                    if batch_updated > 0:
                        print(f"   ✅ Успешно обновлено: {batch_updated} товаров")
                    
                    if batch_errors:
                        print(f"   ⚠️ Ошибки при обновлении: {len(batch_errors)} товаров")
                        # Группируем ошибки по типу
                        error_counts = {}
                        for err in batch_errors:
                            if "MinPrice must be less or equals than Price" in err:
                                error_counts["min_price_too_high"] = error_counts.get("min_price_too_high", 0) + 1
                            elif "NOT_FOUND" in err:
                                error_counts["not_found"] = error_counts.get("not_found", 0) + 1
                            else:
                                error_counts["other"] = error_counts.get("other", 0) + 1
                        
                        if "min_price_too_high" in error_counts:
                            print(f"      - Минимальная цена больше текущей цены продажи: {error_counts['min_price_too_high']} товаров")
                            print(f"        (нужно сначала обновить цену продажи, затем минимальную)")
                        if "not_found" in error_counts:
                            print(f"      - Товар не найден на Ozon: {error_counts['not_found']} товаров")
                        if "other" in error_counts:
                            print(f"      - Другие ошибки: {error_counts['other']} товаров")
                        
                        # Показываем примеры ошибок
                        unique_errors = list(set(batch_errors))[:3]
                        for err in unique_errors:
                            print(f"        Пример: {err}")
                else:
                    print(f"   ⚠️ Неожиданный формат result: {type(results_list)}")
            else:
                print(f"   ⚠️ Неожиданный формат ответа: {result}")
            
        except Exception as e:
            print(f"   ❌ Ошибка при обновлении батча: {e}")
            import traceback
            print(f"   Детали: {traceback.format_exc()}")
            total_errors += len(batch)
    
    print(f"\n✅ Обновление завершено.")
    print(f"   Обновлено товаров: {total_updated}")
    if total_errors > 0:
        print(f"   Ошибок: {total_errors}")


def main():
    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parent
    
    try:
        run(repo_root)
    except FileNotFoundError as e:
        print(f"Ошибка: {e}")
        sys.exit(1)
    except ValueError as e:
        print(f"Ошибка: {e}")
        sys.exit(1)
    except RuntimeError as e:
        print(f"Ошибка: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
