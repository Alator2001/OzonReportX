# -*- coding: utf-8 -*-
"""
Модуль для интерпретации результатов с помощью локального Ollama.
"""

import os
from typing import Dict, List, Any, Optional

import requests
from dotenv import load_dotenv

load_dotenv()

OLLAMA_HOST = (os.getenv("OLLAMA_HOST") or "http://localhost:11434").rstrip("/")
OLLAMA_CHAT_URL = f"{OLLAMA_HOST}/api/chat"
DEFAULT_MODEL = "qwen3:4b"
OLLAMA_MODEL = (os.getenv("OLLAMA_MODEL") or "").strip() or DEFAULT_MODEL


def check_ollama_model() -> Tuple[bool, str]:
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


def _ask_ollama(system_prompt: str, user_prompt: str, max_tokens: int = 1000) -> Optional[str]:
    ok, message = check_ollama_model()
    if not ok:
        print(f"⚠️ {message}")
        return None

    try:
        response = requests.post(
            OLLAMA_CHAT_URL,
            json={
                "model": OLLAMA_MODEL,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                "stream": False,
                "options": {
                    "temperature": 0.4,
                    "num_predict": max_tokens,
                },
            },
            timeout=120,
        )
        response.raise_for_status()
        data = response.json()
        content = data.get("message", {}).get("content", "")
        return content or None
    except requests.exceptions.RequestException as e:
        print(f"⚠️ Ошибка при запросе к Ollama: {e}")
        return None
    except Exception as e:
        print(f"⚠️ Неожиданная ошибка при интерпретации: {e}")
        return None


def interpret_discount_requests_results(
    approved_count: int,
    declined_count: int,
    approved_tasks: List[Dict[str, Any]],
    declined_tasks: List[Dict[str, Any]],
    total_tasks: int,
) -> Optional[str]:
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

    for i, task in enumerate(approved_tasks[:10], 1):
        approved_price = task.get("approved_price", "N/A")
        quantity = task.get("approved_quantity_max", "N/A")
        summary += f"{i}. Заявка ID {task.get('id', 'N/A')}: Цена {approved_price} руб., Количество: {quantity}\n"

    if len(approved_tasks) > 10:
        summary += f"... и ещё {len(approved_tasks) - 10} заявок\n"

    summary += f"\nОтклонённые заявки ({len(declined_tasks)}):\n"

    decline_reasons: Dict[str, int] = {}
    for task in declined_tasks:
        reason = task.get("seller_comment", "Не указана причина")
        decline_reasons[reason] = decline_reasons.get(reason, 0) + 1

    for reason, count in decline_reasons.items():
        summary += f"- {reason}: {count} заявок\n"

    system_prompt = (
        "Ты аналитик Ozon. Кратко и по делу интерпретируй результаты заявок на скидку, "
        "выдели ключевые паттерны и дай практические рекомендации."
    )
    user_prompt = (
        f"Проанализируй следующие результаты обработки заявок на скидку:\n\n{summary}\n\n"
        "Дай краткий анализ и практические рекомендации."
    )
    return _ask_ollama(system_prompt, user_prompt)


def interpret_price_analysis(
    costs_df_summary: Dict[str, Any],
    current_prices_summary: Dict[str, Any],
    actions_summary: Dict[str, Any],
) -> Optional[str]:
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

    system_prompt = (
        "Ты аналитик Ozon. Кратко оцени качество ценовой политики, найди риски и дай "
        "практические рекомендации по оптимизации цен."
    )
    user_prompt = (
        f"Проанализируй следующие данные по ценам:\n\n{summary}\n\n"
        "Дай краткий анализ и практические рекомендации."
    )
    return _ask_ollama(system_prompt, user_prompt)
