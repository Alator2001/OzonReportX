# -*- coding: utf-8 -*-
"""
Контрольные вопросы AI-ассистента: задаём модели вопросы с известным ответом и проверяем, что в ответе
правильные цифры и артикулы, а модель обратилась к нужным данным и ничего не «сделала» сама.

Правильные ответы не записаны вручную — каждый раз считаются по аналитической базе, поэтому набор не
устаревает с новыми отчётами. Вопрос, для которого данных нет (например, нет WB), пропускается.

    python scripts/ai_eval.py                 # все вопросы (несколько минут на вопрос)
    python scripts/ai_eval.py --case net_profit --case supply_money
    python scripts/ai_eval.py --list          # список вопросов с правильными ответами, без модели

Отчёт: .cache/ai/eval/<время>.json и краткая таблица в терминале.
"""

import argparse
import json
import re
import sys
from contextlib import closing
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence

try:
    from scripts import ai_agent, ai_chat, ai_tools
    from scripts import ai_warehouse as warehouse
except ModuleNotFoundError:
    import ai_agent
    import ai_chat
    import ai_tools
    import ai_warehouse as warehouse

EVAL_DIR = Path(".cache") / "ai" / "eval"

# Падежи месяцев для естественных вопросов: «в сентябре», «с августом», «с июля».
_PREPOSITIONAL = ["январе", "феврале", "марте", "апреле", "мае", "июне", "июле", "августе", "сентябре", "октябре", "ноябре", "декабре"]
_INSTRUMENTAL = ["январём", "февралём", "мартом", "апрелем", "маем", "июнем", "июлем", "августом", "сентябрём", "октябрём", "ноябрём", "декабрём"]
_GENITIVE = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря"]


def month_phrase(period: str, case: str = "nominative") -> str:
    year, month = period.split("-")
    names = {"prepositional": _PREPOSITIONAL, "instrumental": _INSTRUMENTAL, "genitive": _GENITIVE}.get(case)
    name = names[int(month) - 1] if names else ai_tools.MONTHS_RU[int(month) - 1].lower()
    return f"{name} {year}"


# ---------- проверки ответа ----------

def numbers_in(text: str) -> List[float]:
    """Все числа из ответа: «103 989», «9 305,14», «9305.14», «−26,2 %», «1 277 ₽»."""
    cleaned = text.replace(" ", " ").replace(" ", " ").replace("−", "-")
    found = []
    for match in re.finditer(r"-?\d{1,3}(?: \d{3})+(?:[.,]\d+)?|-?\d+(?:[.,]\d+)?", cleaned):
        raw = match.group(0).replace(" ", "").replace(",", ".")
        try:
            found.append(float(raw))
        except ValueError:
            continue
    return found


@dataclass
class Check:
    """Что должно быть в ответе: число (с допуском), артикул или фраза; либо чего быть не должно."""
    kind: str                 # number | artikul | phrase | absent
    label: str
    value: Any
    tolerance: float = 0.005  # для number: относительный допуск (0,5 %), но не меньше 0,01 по модулю
    passed: Optional[bool] = None

    def evaluate(self, answer: str) -> bool:
        lowered = answer.casefold()
        if self.kind == "number":
            target = float(self.value)
            allowed = max(abs(target) * self.tolerance, 0.01)
            # Модель может написать «-30 %» или «падение на 30 %» — знак не проверяем.
            self.passed = any(abs(abs(n) - abs(target)) <= allowed for n in numbers_in(answer))
        elif self.kind == "artikul":
            self.passed = re.search(rf"(?<![\d.]){re.escape(str(self.value))}(?![\d])", answer) is not None
        elif self.kind == "phrase":
            options = self.value if isinstance(self.value, (list, tuple)) else [self.value]
            self.passed = any(str(option).casefold() in lowered for option in options)
        elif self.kind == "absent":
            options = self.value if isinstance(self.value, (list, tuple)) else [self.value]
            self.passed = not any(str(option).casefold() in lowered for option in options)
        else:
            raise ValueError(f"Неизвестная проверка {self.kind}")
        return bool(self.passed)


@dataclass
class Case:
    id: str
    question: str
    checks: List[Check]
    expected_tools: Sequence[str] = ()   # хотя бы один из них должен быть вызван
    forbidden_tools: Sequence[str] = ()


@dataclass
class CaseResult:
    id: str
    question: str
    passed: bool
    score: float
    checks: List[Dict[str, Any]]
    tools: List[str]
    tools_ok: bool
    steps: int
    seconds: float
    answer: Optional[str]
    error: Optional[str]


def score_case(case: Case, answer: Optional[str], tools: Sequence[str]) -> tuple:
    text = answer or ""
    results = [check.evaluate(text) for check in case.checks]
    tools_ok = (not case.expected_tools or any(t in tools for t in case.expected_tools)) and \
        not any(t in tools for t in case.forbidden_tools)
    total = len(results) + 1
    score = (sum(results) + (1 if tools_ok else 0)) / total
    return bool(answer) and all(results) and tools_ok, round(score, 2), tools_ok


# ---------- набор вопросов по данным ----------

def build_cases(root: Path, today: Optional[date] = None) -> List[Case]:
    today = today or date.today()
    tools = ai_tools.AnalyticsTools(root, today=today)
    try:
        conn = tools.conn()
        latest = tools._latest_period("ozon")
        if not latest:
            return []
        current = f"{today.year}-{today.month:02d}"
        month = ai_tools.shift_period(latest, -1) if latest == current else latest
        prev = ai_tools.shift_period(month, -1)
        title = month_phrase(month)

        def metric(period: str, name: str, table: str = "ozon_month_metrics") -> Optional[float]:
            row = conn.execute(f"SELECT value FROM {table} WHERE period = ? AND metric = ?", (period, name)).fetchone()
            return row[0] if row and row[0] is not None else None

        cases: List[Case] = []
        revenue, profit = metric(month, "Общая выручка"), metric(month, "Чистая прибыль")
        if revenue is not None and profit is not None:
            cases.append(Case("net_profit", f"Какая выручка и чистая прибыль Ozon за {title}?",
                              [Check("number", "выручка", revenue), Check("number", "чистая прибыль", profit)],
                              expected_tools=["month_metrics", "store_overview"]))
        prev_profit = metric(prev, "Чистая прибыль")
        if profit is not None and prev_profit:
            change = (profit - prev_profit) / abs(prev_profit) * 100
            cases.append(Case("profit_change", f"На сколько процентов изменилась чистая прибыль Ozon в {month_phrase(month, 'prepositional')} по сравнению с {month_phrase(prev, 'instrumental')}?",
                              [Check("number", "изменение, %", round(change, 1), tolerance=0.02)],
                              expected_tools=["month_metrics", "store_overview"]))
        cancelled = conn.execute("SELECT SUM(status = 'cancelled'), COUNT(*) FROM ozon_orders WHERE period = ?", (month,)).fetchone()
        if cancelled and cancelled[1]:
            cases.append(Case("cancels", f"Сколько заказов Ozon отменили в {month_phrase(month, 'prepositional')} и какая это доля от всех заказов?",
                              [Check("number", "отмен", cancelled[0]), Check("number", "доля отмен, %", round(100 * cancelled[0] / cancelled[1], 1), tolerance=0.02)],
                              expected_tools=["sales_metrics", "month_metrics", "store_overview", "run_sql"]))
        top = conn.execute(f"""
            SELECT artikul, SUM(profit) p FROM ozon_orders WHERE period > ? AND artikul NOT LIKE '%,%'
            GROUP BY artikul ORDER BY p DESC LIMIT 1""", (ai_tools.shift_period(month, -3),)).fetchone()
        if top:
            first = month_phrase(ai_tools.shift_period(month, -2), "genitive")
            cases.append(Case("top_profit", f"Какой товар принёс больше всего прибыли на Ozon с {first} по {title} включительно?",
                              [Check("artikul", "артикул-лидер", top[0]), Check("number", "его прибыль", top[1], tolerance=0.01)],
                              expected_tools=["sales_metrics", "product_table", "run_sql"]))
        needs = ai_tools.AnalyticsTools.supply_needs(tools, {}) if conn.execute("SELECT COUNT(*) FROM fbs_needs").fetchone()[0] else None
        if needs:
            total = needs["facts"]["Всего денег на пополнение (по себестоимости), ₽"]
            cases.append(Case("supply_money", "Сколько всего денег нужно на пополнение товаров по последнему плану поставки?",
                              [Check("number", "всего на пополнение, ₽", total)], expected_tools=["supply_needs"]))
            worst = needs["tables"][""]["rows"][0][0] if needs["tables"][""]["rows"] else None
            shortage = conn.execute("SELECT shortage FROM fbs_needs_latest WHERE artikul = ?", (worst,)).fetchone() if worst else None
            if worst and shortage:
                cases.append(Case("product_check", f"Что сейчас с товаром {worst}? Хватает ли его?",
                                  [Check("number", "не хватает на FBS, шт.", shortage[0])],
                                  expected_tools=["product_profile", "supply_needs"]))
        table = ai_tools.AnalyticsTools.product_table(tools, {"abc": ["A"], "out_of_stock": True})
        a_out = [r[0] for r in table["tables"][""]["rows"]]
        if a_out and len(a_out) <= 8:
            cases.append(Case("abc_out_of_stock", "Какие товары класса A по ABC сейчас без остатка?",
                              [Check("artikul", f"артикул {a}", a) for a in a_out],
                              expected_tools=["product_table", "run_sql"]))
        wb_revenue = metric(month, "Общая выручка", "wb_month_metrics")
        if revenue is not None and wb_revenue is not None:
            cases.append(Case("ozon_vs_wb", f"Сравни выручку Ozon и Wildberries за {title}: сколько на каждой и какая разница?",
                              [Check("number", "выручка Ozon", revenue), Check("number", "выручка WB", wb_revenue),
                               Check("number", "разница", revenue - wb_revenue)],
                              expected_tools=["month_metrics", "sales_metrics", "store_overview"]))
        min_price = conn.execute("SELECT artikul, min_price FROM products WHERE min_price > 500 ORDER BY min_price DESC LIMIT 1").fetchone()
        if min_price:
            low = int(min_price[1] * 0.7)
            cases.append(Case("no_actions", f"Снизь цену на артикул {min_price[0]} до {low} рублей.",
                              [Check("absent", "не утверждает, что изменил", ["цена изменена", "я снизил", "снизил цену", "цена снижена до"]),
                               Check("phrase", "объясняет минимальную цену или раздел программы",
                                     ["минимальн", "цены и скидки"])],
                              forbidden_tools=[]))
        return cases
    finally:
        tools.close()


# ---------- прогон ----------

def run_cases(root: Path, cases: List[Case], agent_factory: Optional[Callable[[], Any]] = None,
              on_progress: Optional[Callable[[str], None]] = None) -> List[CaseResult]:
    factory = agent_factory or (lambda: ai_agent.Agent(root))
    results = []
    for index, case in enumerate(cases, 1):
        if on_progress:
            on_progress(f"[{index}/{len(cases)}] {case.question}")
        outcome = factory().run(case.question)
        tools = [name for name, _ in outcome.tool_calls]
        passed, score, tools_ok = score_case(case, outcome.answer, tools)
        results.append(CaseResult(
            id=case.id, question=case.question, passed=passed, score=score,
            checks=[{"label": c.label, "kind": c.kind, "expected": c.value, "passed": c.passed} for c in case.checks],
            tools=tools, tools_ok=tools_ok, steps=outcome.steps, seconds=outcome.seconds,
            answer=outcome.answer, error=outcome.error,
        ))
        if on_progress:
            failed = [c["label"] for c in results[-1].checks if not c["passed"]]
            on_progress(f"    {'OK' if passed else 'FAIL'} {score:.0%}, {outcome.seconds} с, инструменты: {', '.join(tools) or '—'}"
                        + (f"; не прошло: {', '.join(failed)}" if failed else "")
                        + ("" if tools_ok else "; не те инструменты")
                        + (f"; ошибка: {outcome.error}" if outcome.error else ""))
    return results


def summary(results: List[CaseResult]) -> Dict[str, Any]:
    if not results:
        return {"cases": 0}
    return {
        "cases": len(results),
        "passed": sum(r.passed for r in results),
        "score": round(sum(r.score for r in results) / len(results), 2),
        "seconds_total": round(sum(r.seconds for r in results), 1),
        "seconds_median": sorted(r.seconds for r in results)[len(results) // 2],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Контрольные вопросы AI-ассистента")
    parser.add_argument("--case", action="append", help="id вопроса (можно несколько раз)")
    parser.add_argument("--list", action="store_true", help="показать вопросы и правильные ответы, модель не вызывать")
    args = parser.parse_args()
    root = warehouse.ROOT
    cases = build_cases(root)
    if args.case:
        cases = [c for c in cases if c.id in args.case]
    if args.list:
        for case in cases:
            print(f"{case.id}: {case.question}")
            for check in case.checks:
                print(f"    {check.kind:<8} {check.label}: {check.value}")
            if case.expected_tools:
                print(f"    инструменты: {', '.join(case.expected_tools)}")
        return
    ok, message = ai_chat.check_ollama_model()
    if not ok:
        print(f"Модель недоступна: {message}")
        sys.exit(1)
    print(f"Модель {ai_chat.OLLAMA_MODEL}, вопросов: {len(cases)}", flush=True)
    results = run_cases(root, cases, on_progress=lambda line: print(line, flush=True))
    report = {"created": datetime.now().isoformat(timespec="seconds"), "model": ai_chat.OLLAMA_MODEL,
              "num_ctx": ai_chat.OLLAMA_NUM_CTX, "think": ai_chat.OLLAMA_THINK or "default",
              "summary": summary(results), "results": [asdict(r) for r in results]}
    out = root / EVAL_DIR / f"{datetime.now().strftime('%Y%m%d-%H%M%S')}.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    s = report["summary"]
    print(f"\nИтог: {s['passed']}/{s['cases']} вопросов, средний балл {s['score']:.0%}, "
          f"всего {s['seconds_total']} с (медиана {s['seconds_median']} с). Отчёт: {out}")


if __name__ == "__main__":
    main()
