# -*- coding: utf-8 -*-
"""
Модуль для интерпретации результатов с помощью ИИ (Groq API).
"""

import os
import requests
from typing import Dict, List, Any, Optional
from dotenv import load_dotenv

# Загружаем переменные окружения
load_dotenv()
GROQ_API_KEY = os.getenv('GROQ_API_KEY')

GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"

# Модель по умолчанию
DEFAULT_MODEL = "llama-3.3-70b-versatile"


def interpret_discount_requests_results(
    approved_count: int,
    declined_count: int,
    approved_tasks: List[Dict[str, Any]],
    declined_tasks: List[Dict[str, Any]],
    total_tasks: int
) -> Optional[str]:
    """
    Интерпретирует результаты обработки заявок на скидку с помощью Groq API.
    
    Args:
        approved_count: Количество одобренных заявок
        declined_count: Количество отклонённых заявок
        approved_tasks: Список одобренных заявок
        declined_tasks: Список отклонённых заявок
        total_tasks: Общее количество заявок
    
    Returns:
        Интерпретация результатов или None в случае ошибки
    """
    if not GROQ_API_KEY:
        print("⚠️ GROQ_API_KEY не настроен в .env файле. Пропускаем интерпретацию.")
        return None
    
    # Формируем сводку данных
    approved_pct = (approved_count / total_tasks * 100) if total_tasks > 0 else 0.0
    declined_pct = (declined_count / total_tasks * 100) if total_tasks > 0 else 0.0
    
    summary = f"""
Результаты обработки заявок на скидку:

Общая статистика:
- Всего заявок: {total_tasks}
- Одобрено: {approved_count} ({approved_pct:.1f}%)
- Отклонено: {declined_count} ({declined_pct:.1f}%)

Одобренные заявки ({len(approved_tasks)}):
"""
    
    # Добавляем информацию об одобренных заявках
    for i, task in enumerate(approved_tasks[:10], 1):  # Показываем первые 10
        approved_price = task.get('approved_price', 'N/A')
        quantity = task.get('approved_quantity_max', 'N/A')
        summary += f"{i}. Заявка ID {task.get('id', 'N/A')}: Цена {approved_price} руб., Количество: {quantity}\n"
    
    if len(approved_tasks) > 10:
        summary += f"... и ещё {len(approved_tasks) - 10} заявок\n"
    
    summary += f"\nОтклонённые заявки ({len(declined_tasks)}):\n"
    
    # Группируем причины отклонения
    decline_reasons = {}
    for task in declined_tasks:
        reason = task.get('seller_comment', 'Не указана причина')
        if reason not in decline_reasons:
            decline_reasons[reason] = 0
        decline_reasons[reason] += 1
    
    for reason, count in decline_reasons.items():
        summary += f"- {reason}: {count} заявок\n"
    
    # Формируем запрос к ИИ
    system_prompt = """Ты опытный аналитик данных для маркетплейса Ozon. 
Твоя задача - анализировать результаты обработки заявок на скидку и давать практические рекомендации.

Анализируй:
1. Соотношение одобренных и отклонённых заявок
2. Основные причины отклонения заявок
3. Средние цены в одобренных заявках
4. Давай конкретные рекомендации по улучшению ценовой политики

Отвечай кратко, по делу, на русском языке. Используй эмодзи для визуального выделения."""

    user_prompt = f"Проанализируй следующие результаты обработки заявок на скидку:\n\n{summary}\n\nДай краткий анализ и практические рекомендации."

    try:
        response = requests.post(
            GROQ_API_URL,
            headers={
                "Authorization": f"Bearer {GROQ_API_KEY}",
                "Content-Type": "application/json"
            },
            json={
                "model": DEFAULT_MODEL,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                "temperature": 0.7,
                "max_tokens": 1000
            },
            timeout=30
        )
        
        response.raise_for_status()
        data = response.json()
        
        interpretation = data.get("choices", [{}])[0].get("message", {}).get("content", "")
        
        if interpretation:
            return interpretation
        else:
            print("⚠️ ИИ не вернул интерпретацию.")
            return None
            
    except requests.exceptions.RequestException as e:
        print(f"⚠️ Ошибка при запросе к Groq API: {e}")
        if hasattr(e, 'response') and e.response is not None:
            try:
                error_data = e.response.json()
                print(f"   Детали ошибки: {error_data}")
            except:
                print(f"   Ответ сервера: {e.response.text[:200]}")
        return None
    except Exception as e:
        print(f"⚠️ Неожиданная ошибка при интерпретации: {e}")
        return None


def interpret_price_analysis(
    costs_df_summary: Dict[str, Any],
    current_prices_summary: Dict[str, Any],
    actions_summary: Dict[str, Any]
) -> Optional[str]:
    """
    Интерпретирует результаты анализа цен с помощью Groq API.
    
    Args:
        costs_df_summary: Сводка по себестоимости
        current_prices_summary: Сводка по текущим ценам
        actions_summary: Сводка по акциям
    
    Returns:
        Интерпретация результатов или None в случае ошибки
    """
    if not GROQ_API_KEY:
        print("⚠️ GROQ_API_KEY не настроен в .env файле. Пропускаем интерпретацию.")
        return None
    
    # Формируем сводку данных
    summary = f"""
Анализ ценовой политики:

Себестоимость:
- Всего товаров: {costs_df_summary.get('total_products', 0)}
- Средняя себестоимость: {costs_df_summary.get('avg_cost', 0):.2f} руб.
- Минимальная цена продажи: {costs_df_summary.get('min_price_avg', 0):.2f} руб.
- Желательная цена продажи: {costs_df_summary.get('desired_price_avg', 0):.2f} руб.

Текущие цены на Ozon:
- Товаров с ценой: {current_prices_summary.get('products_with_price', 0)}
- Средняя текущая цена: {current_prices_summary.get('avg_current_price', 0):.2f} руб.
- Средняя цена с акциями: {current_prices_summary.get('avg_marketing_price', 0):.2f} руб.
- Товаров ниже минимальной цены: {current_prices_summary.get('below_min_price', 0)}

Акции:
- Активных акций: {actions_summary.get('active_actions', 0)}
- Товаров в акциях: {actions_summary.get('products_in_actions', 0)}
- Товаров с невыгодными ценами в акциях: {actions_summary.get('unprofitable_in_actions', 0)}
"""
    
    system_prompt = """Ты опытный аналитик данных для маркетплейса Ozon. 
Твоя задача - анализировать ценовую политику и давать практические рекомендации.

Анализируй:
1. Соответствие текущих цен минимальным и желательным ценам
2. Проблемные товары (ниже минимальной цены)
3. Эффективность акций
4. Давай конкретные рекомендации по оптимизации цен

Отвечай кратко, по делу, на русском языке. Используй эмодзи для визуального выделения."""

    user_prompt = f"Проанализируй следующие данные по ценам:\n\n{summary}\n\nДай краткий анализ и практические рекомендации."

    try:
        response = requests.post(
            GROQ_API_URL,
            headers={
                "Authorization": f"Bearer {GROQ_API_KEY}",
                "Content-Type": "application/json"
            },
            json={
                "model": DEFAULT_MODEL,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                "temperature": 0.7,
                "max_tokens": 1000
            },
            timeout=30
        )
        
        response.raise_for_status()
        data = response.json()
        
        interpretation = data.get("choices", [{}])[0].get("message", {}).get("content", "")
        
        if interpretation:
            return interpretation
        else:
            print("⚠️ ИИ не вернул интерпретацию.")
            return None
            
    except requests.exceptions.RequestException as e:
        print(f"⚠️ Ошибка при запросе к Groq API: {e}")
        return None
    except Exception as e:
        print(f"⚠️ Неожиданная ошибка при интерпретации: {e}")
        return None
