"""Offline checks for the AI assistant agent loop with a scripted model: tool protocol, limits, retries, context budget."""
import sys
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "tests"))
import ai_agent  # noqa: E402
import ai_chat  # noqa: E402
from test_ai_warehouse import WarehouseTestCase  # noqa: E402


def call(name, **args):
    return {"function": {"name": name, "arguments": args}}


class ScriptedModel:
    """Отвечает заранее заданными ответами и запоминает, что ему прислали."""

    def __init__(self, *responses):
        self.responses = list(responses)
        self.requests = []

    def __call__(self, messages, tools=None, temperature=0.3, on_chunk=None, on_thinking=None, should_cancel=None, **_):
        self.requests.append({"messages": [dict(m) for m in messages], "tools": tools})
        response = self.responses.pop(0)
        if response.content and on_chunk:
            on_chunk(response.content)
        return response


def reply(content="", tool_calls=None, **kwargs):
    return ai_chat.ChatResponse(content=content, tool_calls=tool_calls or [], model="test-model", **kwargs)


class AgentTestCase(WarehouseTestCase):
    def run_agent(self, model, question="Как дела?", history=None, max_steps=ai_agent.MAX_STEPS):
        self.events = []
        agent = ai_agent.Agent(self.root, chat=model, today=date(2026, 10, 24), max_steps=max_steps)
        return agent.run(question, history, on_event=lambda kind, payload: self.events.append((kind, payload)))


class LoopTests(AgentTestCase):
    def test_tool_call_then_answer_uses_native_protocol(self):
        model = ScriptedModel(reply(tool_calls=[call("sales_metrics", period_from="2026-09", period_to="2026-09", metrics=["revenue"])]),
                              reply("Выручка за сентябрь — 960 ₽."))
        result = self.run_agent(model)
        self.assertEqual(result.answer, "Выручка за сентябрь — 960 ₽.")
        self.assertEqual(result.tool_calls, [("sales_metrics", {"period_from": "2026-09", "period_to": "2026-09", "metrics": ["revenue"]})])
        second = model.requests[1]["messages"]
        self.assertEqual(second[-2]["role"], "assistant")
        self.assertEqual(second[-2]["tool_calls"][0]["function"]["name"], "sales_metrics")
        self.assertEqual(second[-1]["role"], "tool")
        self.assertEqual(second[-1]["tool_name"], "sales_metrics")
        self.assertIn("| 960 |", second[-1]["content"])
        self.assertIsNotNone(model.requests[1]["tools"])
        self.assertIn(("status", "Запрашиваю данные: показатели продаж"), self.events)

    def test_system_prompt_lists_available_data(self):
        model = ScriptedModel(reply("ok"))
        self.run_agent(model)
        system = model.requests[0]["messages"][0]["content"]
        self.assertIn("Сегодня 2026-10-24", system)
        self.assertIn("Месячные отчёты Ozon: 2026-09 … 2026-09", system)
        self.assertIn("Месячные отчёты Wildberries: 2026-09 … 2026-09", system)

    def test_history_is_trimmed_to_recent_turns(self):
        history = []
        for i in range(5):
            history += [{"role": "user", "content": f"q{i}"}, {"role": "assistant", "content": f"a{i}"}]
        model = ScriptedModel(reply("ok"))
        self.run_agent(model, question="новый", history=history)
        contents = [m["content"] for m in model.requests[0]["messages"][1:]]
        self.assertEqual(contents, ["q2", "a2", "q3", "a3", "q4", "a4", "новый"])

    def test_repeated_call_is_not_executed_twice(self):
        same = call("describe_data")
        model = ScriptedModel(reply(tool_calls=[same]), reply(tool_calls=[same]), reply("ok"))
        with patch.object(ai_agent.ai_tools.AnalyticsTools, "describe_data", autospec=True,
                          return_value={"title": "Каталог"}) as describe:
            result = self.run_agent(model)
        self.assertEqual(describe.call_count, 1)
        self.assertIn("уже был выше", model.requests[2]["messages"][-1]["content"])
        self.assertEqual(result.answer, "ok")

    def test_step_limit_forces_answer_without_tools(self):
        model = ScriptedModel(reply(tool_calls=[call("describe_data")]), reply(tool_calls=[call("find_products", query="доска")]),
                              reply("Итог по собранному."))
        result = self.run_agent(model, max_steps=2)
        self.assertEqual(result.answer, "Итог по собранному.")
        self.assertIsNone(model.requests[2]["tools"])
        self.assertEqual(model.requests[2]["messages"][-1]["content"], ai_agent.FINAL_NUDGE)

    def test_empty_answer_is_retried_once_without_tools(self):
        model = ScriptedModel(reply(""), reply("Теперь ответ."))
        self.assertEqual(self.run_agent(model).answer, "Теперь ответ.")
        self.assertIsNone(model.requests[1]["tools"])
        self.assertEqual(model.requests[1]["messages"][-1]["content"], ai_agent.EMPTY_NUDGE)
        model = ScriptedModel(reply(""), reply(""))
        result = self.run_agent(model)
        self.assertIsNone(result.answer)
        self.assertIn("не дала ответа", result.error)

    def test_text_before_tool_call_is_reset(self):
        model = ScriptedModel(reply("Сейчас посмотрю…", tool_calls=[call("describe_data")]), reply("Ответ"))
        self.run_agent(model)
        kinds = [kind for kind, _ in self.events]
        self.assertLess(kinds.index("answer_reset"), kinds.index("tool"))

    def test_errors_and_cancel(self):
        result = self.run_agent(ScriptedModel(reply(error="Нет связи с Ollama")))
        self.assertEqual(result.error, "Нет связи с Ollama")
        result = self.run_agent(ScriptedModel(reply(cancelled=True)))
        self.assertTrue(result.cancelled)
        agent = ai_agent.Agent(self.root, chat=ScriptedModel(), today=date(2026, 10, 24))
        self.assertTrue(agent.run("q", should_cancel=lambda: True).cancelled)

    def test_unknown_tool_error_goes_back_to_model(self):
        model = ScriptedModel(reply(tool_calls=[call("delete_everything")]), reply("Не могу."))
        self.run_agent(model)
        self.assertIn("не найден", model.requests[1]["messages"][-1]["content"])


class ContextTests(AgentTestCase):
    def test_large_result_is_cut(self):
        model = ScriptedModel(reply(tool_calls=[call("run_sql", sql="SELECT 1")]), reply("ok"))
        huge = {"title": "x", "facts": {"данные": "я" * 100000}}
        with patch.object(ai_agent.ai_tools.AnalyticsTools, "run_sql", autospec=True, return_value=huge):
            self.run_agent(model)
        content = model.requests[1]["messages"][-1]["content"]
        self.assertLess(len(content), 100000)
        self.assertIn("обрезано", content)

    def test_old_results_shrink_when_window_is_full(self):
        agent = ai_agent.Agent(self.root, chat=ScriptedModel())
        messages = [{"role": "system", "content": "s"}, {"role": "tool", "content": "a" * 20000},
                    {"role": "tool", "content": "b" * 20000}, {"role": "user", "content": "q"}]
        agent._fit_context(messages, num_ctx=ai_agent.RESERVED_OUTPUT_TOKENS + 12000)
        self.assertLess(len(messages[1]["content"]), 2000)
        self.assertEqual(len(messages[2]["content"]), 20000)
        self.assertLessEqual(ai_agent.estimate_tokens(messages), 12000)


if __name__ == "__main__":
    unittest.main()
