# -*- coding: utf-8 -*-
"""
Агентный цикл AI-ассистента: вопрос → модель вызывает инструменты → ответ.

Протокол — встроенный вызов инструментов Ollama: сообщение ассистента с `tool_calls`, затем по сообщению
`role: tool` на каждый вызов. Цикл следит за окном контекста (сжимает старые результаты), не повторяет
одинаковые вызовы, после MAX_STEPS просит ответить по собранным данным и переспрашивает при пустом ответе.

    python scripts/ai_agent.py "Как дела в магазине?"     # задать вопрос из терминала
"""

import json
import sys
import time
from contextlib import closing
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

try:
    from scripts import ai_chat, ai_tools
    from scripts import ai_warehouse as warehouse
except ModuleNotFoundError:
    import ai_chat
    import ai_tools
    import ai_warehouse as warehouse

MAX_STEPS = 8
TEMPERATURE = 0.3
# Сколько прошлых пар «вопрос — ответ» видит модель. Данные прошлых вопросов не передаются: нужно — запросит снова.
HISTORY_TURNS = 3
# Оценка токенов по символам: русский текст с таблицами — около 2,3 символа на токен (замер на qwen3.5).
CHARS_PER_TOKEN = 2.3
# Запас окна под размышления и ответ модели.
RESERVED_OUTPUT_TOKENS = 6000
# Один результат инструмента — не больше этой доли окна; старые результаты сжимаются до SHRUNK_RESULT_CHARS.
MAX_RESULT_SHARE = 0.35
SHRUNK_RESULT_CHARS = 1200

TOOL_LABELS = {
    "store_overview": "обзор магазина",
    "month_metrics": "итоги месяцев",
    "sales_metrics": "показатели продаж",
    "product_table": "таблица товаров",
    "product_profile": "карточка товара",
    "supply_needs": "план пополнения",
    "find_products": "поиск товара",
    "calculate": "расчёт",
    "describe_data": "каталог данных",
    "run_sql": "запрос к базе",
    "live_prices": "цены в Ozon сейчас",
    "live_stocks": "остатки в Ozon сейчас",
    "live_promotions": "акции в Ozon сейчас",
    "live_discount_requests": "заявки на скидку",
}

SYSTEM_PROMPT = """Ты — AI-аналитик продавца маркетплейсов Ozon и Wildberries в программе OzonReportX. Сегодня {today}.
Работаешь локально: данные магазина никуда не уходят.

# Какие данные есть
{data_summary}

# Как работать
- Данные получай только через инструменты, вызывай их сам; независимые вызовы — сразу несколько.
- Общий вопрос («как дела», «что улучшить», «что происходит») — начни со store_overview.
- Один товар — product_profile. Вопрос «какие товары…» по нескольким признакам — product_table.
- Что пополнить и сколько денег нужно — supply_needs. Рейтинги, сравнения месяцев, FBO/FBS, отмены — sales_metrics.
- Чистая прибыль, реклама, комиссии, логистика по месяцам — month_metrics. Wildberries — marketplace="wb".
- Нужного инструмента нет — describe_data, затем run_sql (SQLite, только SELECT). Представления *_latest уже свежие.
- Текущие цены, остатки, акции, заявки на скидку прямо сейчас — live_*; это запросы в Ozon, вызывай, когда нужна
  именно текущая картина или снимок устарел.
- Не знаешь артикул — find_products. Не повторяй вызов с теми же аргументами: результат уже есть выше.

# Точность
- Все цифры — только из ответов инструментов. Не считай в уме: любую разницу, долю, процент изменения,
  «во сколько раз» или сумму нескольких чисел посчитай инструментом calculate (можно сразу несколько выражений).
  Если цифры нет — так и скажи.
- Называй период каждой цифры. Текущий месяц не закончен — не сравнивай его напрямую с полными месяцами.
- «profit» в sales_metrics — прибыль по заказам до рекламы и хранения; чистая прибыль — только из month_metrics.
- Разбирая товар, проверь все сигналы: остаток, индекс цен (RED — цена невыгодна), акции, органику, отмены, поиск.
- Никогда не советуй цену ниже минимальной цены продажи товара: это ниже минимальной маржи. Если для выгодного
  индекса цен нужна цена ниже минимальной (live_prices пишет «снижать нельзя»), скажи, что конкурировать ценой
  нельзя, и предложи другое: карточка, отзывы, доставка, реклама.
- Ты ничего не меняешь в Ozon. Предлагая действие, называй раздел программы: «Цены и скидки», «Акции»,
  «Реклама», «Поставки», «Анализ», «Поиск».

# Ответ
- На русском, по делу: сначала вывод в 1–2 предложениях, затем цифры (таблица markdown для сравнений),
  в конце 2–4 конкретных действия.
- Без названий инструментов, таблиц базы и SQL в тексте ответа.
"""

FINAL_NUDGE = "Данных достаточно. Не вызывай инструменты — ответь пользователю по уже полученным данным."
EMPTY_NUDGE = "Ответ получился пустым. Сформулируй ответ пользователю по уже полученным данным, обычным текстом."


@dataclass
class AgentResult:
    answer: Optional[str] = None
    error: Optional[str] = None
    cancelled: bool = False
    steps: int = 0
    tool_calls: List[Tuple[str, Dict[str, Any]]] = field(default_factory=list)
    model: Optional[str] = None
    seconds: float = 0.0


def estimate_tokens(messages: List[Dict[str, Any]]) -> int:
    chars = 0
    for message in messages:
        chars += len(message.get("content") or "")
        for call in message.get("tool_calls") or []:
            chars += len(json.dumps(call, ensure_ascii=False))
    return int(chars / CHARS_PER_TOKEN)


def data_summary(root: Path) -> str:
    """Короткая справка о данных для системного промпта: какие месяцы есть и насколько свежи снимки."""
    try:
        with closing(warehouse.connect(root)) as conn:
            catalog = {item["table"]: item for item in warehouse.describe(conn)}
    except Exception as exc:  # без базы ассистент всё равно отвечает: инструменты вернут понятную ошибку
        return f"- Аналитическая база недоступна: {exc}"
    lines = []

    def span(table: str) -> Optional[str]:
        item = catalog.get(table) or {}
        value = item.get("range")
        return f"{value[0]} … {value[1]}" if value and item.get("rows") else None

    for table, label in (("ozon_orders", "Месячные отчёты Ozon"), ("wb_orders", "Месячные отчёты Wildberries")):
        if span(table):
            lines.append(f"- {label}: {span(table)}")
    for table, label in (("organic_products", "Анализ органики"), ("product_snapshots", "Снимки цены и остатка"),
                         ("fbo_supply_plan", "План поставки"), ("search_sku_daily", "Поисковая аналитика"),
                         ("abc_xyz", "ABC/XYZ-отчёты")):
        item = catalog.get(table) or {}
        if item.get("rows"):
            when = span(table)
            lines.append(f"- {label}: {('по ' + when.split(' … ')[-1][:10]) if when else 'есть'}")
    errors = [f"{t}: {i['error']}" for t, i in catalog.items() if i.get("error")]
    if errors:
        lines.append("- Не прочитаны: " + "; ".join(errors))
    return "\n".join(lines) or "- Данных пока нет: сформируйте месячный отчёт в программе."


class Agent:
    def __init__(
        self,
        root: Path,
        tools: Optional[ai_tools.AnalyticsTools] = None,
        chat: Callable[..., ai_chat.ChatResponse] = ai_chat.chat,
        today: Optional[date] = None,
        max_steps: int = MAX_STEPS,
    ):
        self.root = Path(root)
        self.tools = tools or ai_tools.AnalyticsTools(self.root, today=today)
        self.chat = chat
        self.today = today
        self.max_steps = max_steps

    def system_prompt(self) -> str:
        today = self.today or date.today()
        return SYSTEM_PROMPT.format(today=today.isoformat(), data_summary=data_summary(self.root))

    def _fit_context(self, messages: List[Dict[str, Any]], num_ctx: int) -> None:
        """Сжимает самые старые результаты инструментов, пока запрос не влезет в окно с запасом на ответ."""
        budget = max(2000, num_ctx - RESERVED_OUTPUT_TOKENS)
        for message in messages:
            if estimate_tokens(messages) <= budget:
                return
            if message.get("role") == "tool" and not message.get("_shrunk"):
                content = message["content"]
                if len(content) > SHRUNK_RESULT_CHARS:
                    message["content"] = content[:SHRUNK_RESULT_CHARS] + "\n[…результат сокращён, чтобы уместиться в память модели; если нужен целиком — вызови инструмент снова]"
                message["_shrunk"] = True

    def run(
        self,
        question: str,
        history: Optional[List[Dict[str, str]]] = None,
        on_event: Optional[Callable[[str, Any], None]] = None,
        should_cancel: Optional[Callable[[], bool]] = None,
    ) -> AgentResult:
        """
        on_event(kind, payload): "status" (текст этапа), "thinking"/"answer_chunk" (кусочки),
        "answer_reset" (начатый текст оказался не ответом), "tool" (вызов и результат), "model" (данные запроса).
        """
        emit = on_event or (lambda kind, payload: None)
        cancelled = should_cancel or (lambda: False)
        started = time.monotonic()
        result = AgentResult(model=ai_chat.OLLAMA_MODEL)
        num_ctx = ai_chat.OLLAMA_NUM_CTX
        max_result_chars = int((num_ctx - RESERVED_OUTPUT_TOKENS) * MAX_RESULT_SHARE * CHARS_PER_TOKEN)

        messages: List[Dict[str, Any]] = [{"role": "system", "content": self.system_prompt()}]
        for item in (history or [])[-2 * HISTORY_TURNS:]:
            if item.get("role") in ("user", "assistant") and item.get("content"):
                messages.append({"role": item["role"], "content": item["content"]})
        messages.append({"role": "user", "content": question})
        seen: Dict[str, str] = {}
        nudged_empty = False
        final_step = False

        try:
            while True:
                if cancelled():
                    result.cancelled = True
                    return result
                result.steps += 1
                use_tools = not final_step and result.steps <= self.max_steps
                emit("status", "Думаю над вопросом…" if result.steps == 1 else ("Формулирую ответ…" if not use_tools else "Анализирую данные…"))
                self._fit_context(messages, num_ctx)
                streamed: List[str] = []

                def on_chunk(chunk: str) -> None:
                    streamed.append(chunk)
                    emit("answer_chunk", chunk)

                response = self.chat(
                    [{k: v for k, v in m.items() if not k.startswith("_")} for m in messages],
                    tools=ai_tools.TOOL_SCHEMAS if use_tools else None,
                    temperature=TEMPERATURE,
                    on_chunk=on_chunk,
                    on_thinking=lambda chunk: emit("thinking", chunk),
                    should_cancel=cancelled,
                )
                result.model = response.model or result.model
                emit("model", {"step": result.steps, "tools": use_tools, **response.runtime()})
                if response.cancelled:
                    result.cancelled = True
                    return result
                if response.error:
                    result.error = response.error
                    return result

                calls = ai_chat.parse_tool_calls(response.tool_calls) if use_tools else []
                if calls:
                    if streamed:
                        emit("answer_reset", None)  # текст перед вызовом инструментов — не ответ
                    messages.append({"role": "assistant", "content": response.content, "tool_calls": response.tool_calls})
                    labels = ", ".join(dict.fromkeys(TOOL_LABELS.get(name, name) for name, _ in calls))
                    emit("status", f"Запрашиваю данные: {labels}")
                    for name, args in calls:
                        if cancelled():
                            result.cancelled = True
                            return result
                        key = name + json.dumps(args, ensure_ascii=False, sort_keys=True)
                        tool_started = time.monotonic()
                        if key in seen:
                            text = "Этот вызов с теми же аргументами уже был выше — используй тот результат."
                        else:
                            text = ai_tools.render(self.tools.call(name, args))
                            if len(text) > max_result_chars:
                                text = text[:max_result_chars] + "\n[…обрезано: результат больше памяти модели. Сузьте запрос: фильтры, группировка, limit]"
                            seen[key] = text
                        result.tool_calls.append((name, args))
                        emit("tool", {"name": name, "args": args, "result": text, "seconds": round(time.monotonic() - tool_started, 2)})
                        messages.append({"role": "tool", "tool_name": name, "content": text})
                    if result.steps >= self.max_steps:
                        final_step = True
                        messages.append({"role": "user", "content": FINAL_NUDGE})
                    continue

                if response.content:
                    result.answer = response.content
                    return result
                if nudged_empty:
                    result.error = "Модель не дала ответа. Попробуйте переформулировать вопрос."
                    return result
                nudged_empty = True
                final_step = True
                messages.append({"role": "user", "content": EMPTY_NUDGE})
        finally:
            result.seconds = round(time.monotonic() - started, 1)
            self.tools.close()


def main() -> None:
    question = " ".join(sys.argv[1:]).strip() or "Как дела в магазине?"
    agent = Agent(warehouse.ROOT)

    def on_event(kind: str, payload: Any) -> None:
        if kind == "status":
            print(f"· {payload}", flush=True)
        elif kind == "tool":
            print(f"  → {payload['name']} {json.dumps(payload['args'], ensure_ascii=False)} ({payload['seconds']} с)", flush=True)

    result = agent.run(question, on_event=on_event)
    print()
    print(result.answer or f"Ошибка: {result.error}")
    print(f"\n[{result.model}, шагов {result.steps}, вызовов {len(result.tool_calls)}, {result.seconds} с]")


if __name__ == "__main__":
    main()
