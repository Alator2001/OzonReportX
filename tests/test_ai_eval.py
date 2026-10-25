"""Offline checks for the AI assistant control questions: number parsing, answer checks, scoring, case building."""
import sys
import unittest
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "tests"))
import ai_agent  # noqa: E402
import ai_eval  # noqa: E402
from test_ai_warehouse import WarehouseTestCase  # noqa: E402


class CheckTests(unittest.TestCase):
    def test_numbers_in_russian_formats(self):
        text = "Выручка 103 989 ₽, прибыль 9 305,14 ₽ (−26,2 %), заказов 114, итог 176 690 ₽, цена 1277.50"
        self.assertEqual(ai_eval.numbers_in(text), [103989.0, 9305.14, -26.2, 114.0, 176690.0, 1277.5])

    def test_number_check_tolerance_and_sign(self):
        self.assertTrue(ai_eval.Check("number", "x", 9305.140000000003).evaluate("прибыль 9 305,14 ₽"))
        self.assertTrue(ai_eval.Check("number", "x", -26.2, tolerance=0.02).evaluate("упала на 26,16%"))
        self.assertFalse(ai_eval.Check("number", "x", 81727).evaluate("разница 71 727 ₽"))
        self.assertFalse(ai_eval.Check("number", "x", 103989).evaluate("около 104 тыс."))

    def test_artikul_phrase_absent(self):
        self.assertTrue(ai_eval.Check("artikul", "a", "1102").evaluate("Товар 1102 — нет остатка"))
        self.assertFalse(ai_eval.Check("artikul", "a", "110").evaluate("Товар 1102"))
        self.assertTrue(ai_eval.Check("phrase", "p", ["минимальн", "цены и скидки"]).evaluate("Ниже Минимальной нельзя"))
        self.assertFalse(ai_eval.Check("absent", "n", ["я снизил"]).evaluate("Готово, я снизил цену"))

    def test_score_case(self):
        case = ai_eval.Case("c", "q", [ai_eval.Check("number", "n", 5)], expected_tools=["month_metrics"])
        self.assertEqual(ai_eval.score_case(case, "Ответ 5", ["month_metrics"]), (True, 1.0, True))
        self.assertEqual(ai_eval.score_case(case, "Ответ 5", ["run_sql"]), (False, 0.5, False))
        self.assertEqual(ai_eval.score_case(case, None, ["month_metrics"])[0], False)

    def test_month_phrases(self):
        self.assertEqual(ai_eval.month_phrase("2026-09", "prepositional"), "сентябре 2026")
        self.assertEqual(ai_eval.month_phrase("2026-08", "instrumental"), "августом 2026")
        self.assertEqual(ai_eval.month_phrase("2026-07", "genitive"), "июля 2026")
        self.assertEqual(ai_eval.month_phrase("2026-09"), "сентябрь 2026")


class CasesTests(WarehouseTestCase):
    def test_cases_follow_the_data(self):
        cases = {c.id: c for c in ai_eval.build_cases(self.root, today=date(2026, 10, 24))}
        self.assertIn("net_profit", cases)
        self.assertEqual([c.value for c in cases["net_profit"].checks], [103989.0, 9305.14])
        self.assertIn("сентябрь 2026", cases["net_profit"].question)
        self.assertEqual([c.value for c in cases["ozon_vs_wb"].checks], [103989.0, 22262.0, 81727.0])
        self.assertEqual(cases["cancels"].checks[0].value, 1)
        self.assertEqual(cases["top_profit"].checks[0].value, "7012")
        self.assertIn("no_actions", cases)
        self.assertNotIn("supply_money", cases)  # плана поставки в данных нет — вопрос не задаём
        self.assertNotIn("profit_change", cases)  # прошлого месяца нет

    def test_current_month_is_skipped_for_month_questions(self):
        cases = {c.id: c for c in ai_eval.build_cases(self.root, today=date(2026, 9, 15))}
        self.assertNotIn("сентябрь 2026", cases.get("net_profit", ai_eval.Case("", "", [])).question)

    def test_run_cases_with_scripted_agent(self):
        case = ai_eval.Case("c", "Выручка?", [ai_eval.Check("number", "выручка", 960)], expected_tools=["sales_metrics"])

        class FakeAgent:
            def run(self, question):
                return ai_agent.AgentResult(answer="Выручка 960 ₽", steps=2, seconds=1.5, tool_calls=[("sales_metrics", {})])

        lines = []
        results = ai_eval.run_cases(self.root, [case], agent_factory=FakeAgent, on_progress=lines.append)
        self.assertTrue(results[0].passed)
        self.assertEqual(ai_eval.summary(results)["passed"], 1)
        self.assertTrue(lines[-1].strip().startswith("OK 100%"))


if __name__ == "__main__":
    unittest.main()
