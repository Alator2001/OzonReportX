"""Offline checks for the Ollama client of the AI assistant: payload, streaming, tool calls, model checks."""
import json
import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import ai_chat  # noqa: E402


def response(data, status=200):
    resp = MagicMock(status_code=status)
    resp.json.return_value = data
    resp.raise_for_status.return_value = None
    return resp


def stream_response(chunks, status=200):
    resp = MagicMock(status_code=status)
    resp.iter_lines.return_value = [json.dumps(chunk, ensure_ascii=False) for chunk in chunks]
    resp.__enter__.return_value = resp
    resp.__exit__.return_value = False
    return resp


TOOL_CALL = {"function": {"name": "store_overview", "arguments": {"period": "Сентябрь 2026"}}}


class TransportTestCase(unittest.TestCase):
    def setUp(self):
        self.capabilities = ["completion", "tools", "thinking"]
        patcher = patch.object(ai_chat, "model_capabilities", side_effect=lambda model=None: self.capabilities)
        patcher.start()
        self.addCleanup(patcher.stop)


class PayloadTests(TransportTestCase):
    def test_context_window_keep_alive_and_tools(self):
        payload = ai_chat._build_ollama_payload([], 0.2, None, True, tools=[{"type": "function"}])
        self.assertEqual(payload["options"]["num_ctx"], ai_chat.OLLAMA_NUM_CTX)
        self.assertEqual(payload["keep_alive"], ai_chat.OLLAMA_KEEP_ALIVE)
        self.assertEqual(payload["tools"], [{"type": "function"}])
        self.assertNotIn("tools", ai_chat._build_ollama_payload([], 0.2, None, True))

    def test_think_is_sent_only_when_configured_and_supported(self):
        with patch.object(ai_chat, "OLLAMA_THINK", "0"):
            self.assertIs(ai_chat._build_ollama_payload([], 0.2, None, False)["think"], False)
            self.capabilities = ["completion", "tools"]
            self.assertNotIn("think", ai_chat._build_ollama_payload([], 0.2, None, False))
        with patch.object(ai_chat, "OLLAMA_THINK", ""):
            self.capabilities = ["completion", "tools", "thinking"]
            self.assertNotIn("think", ai_chat._build_ollama_payload([], 0.2, None, False))


class ChatTests(TransportTestCase):
    def test_stream_collects_content_thinking_and_tool_calls(self):
        chunks = [
            {"message": {"content": "", "thinking": "Нужна "}},
            {"message": {"content": "", "thinking": "сводка"}},
            {"message": {"content": "", "tool_calls": [TOOL_CALL]}},
            {"message": {"content": "Итог"}, "done": True, "prompt_eval_count": 300, "eval_count": 20, "done_reason": "stop", "model": "m"},
        ]
        thinking, text = [], []
        with patch.object(ai_chat.requests, "post", return_value=stream_response(chunks)) as post:
            result = ai_chat.chat([{"role": "user", "content": "q"}], tools=[{"type": "function"}],
                                  on_chunk=text.append, on_thinking=thinking.append)
        self.assertTrue(post.call_args.kwargs["json"]["stream"])
        self.assertEqual(result.thinking, "Нужна сводка")
        self.assertEqual("".join(thinking), "Нужна сводка")
        self.assertEqual(result.content, "Итог")
        self.assertEqual(text, ["Итог"])
        self.assertEqual(ai_chat.parse_tool_calls(result.tool_calls), [("store_overview", {"period": "Сентябрь 2026"})])
        self.assertEqual((result.prompt_tokens, result.output_tokens, result.model), (300, 20, "m"))
        self.assertFalse(result.context_overflow)

    def test_context_overflow(self):
        chunks = [{"message": {"content": "ok"}, "done": True, "prompt_eval_count": ai_chat.OLLAMA_NUM_CTX}]
        with patch.object(ai_chat.requests, "post", return_value=stream_response(chunks)):
            self.assertTrue(ai_chat.chat([]).runtime()["context_overflow"])

    def test_http_error_and_network_error_do_not_raise(self):
        failed = stream_response([], status=400)
        failed.json.return_value = {"error": "model does not support tools"}
        with patch.object(ai_chat.requests, "post", return_value=failed):
            self.assertIn("does not support tools", ai_chat.chat([]).error)
        with patch.object(ai_chat.requests, "post", side_effect=ai_chat.requests.exceptions.ConnectionError()):
            self.assertIn("Нет связи с Ollama", ai_chat.chat([]).error)

    def test_cancel_stops_reading(self):
        chunks = [{"message": {"content": "a"}}, {"message": {"content": "b"}, "done": True}]
        with patch.object(ai_chat.requests, "post", return_value=stream_response(chunks)):
            self.assertTrue(ai_chat.chat([], should_cancel=lambda: True).cancelled)

    def test_string_arguments_are_decoded(self):
        calls = [{"function": {"name": "sales_metrics", "arguments": '{"limit": 5}'}},
                 {"function": {"name": "x", "arguments": "не json"}}, {"bad": 1}]
        self.assertEqual(ai_chat.parse_tool_calls(calls), [("sales_metrics", {"limit": 5}), ("x", {})])


class ModelCheckTests(unittest.TestCase):
    def setUp(self):
        ai_chat._model_capabilities_cache.clear()
        self.addCleanup(ai_chat._model_capabilities_cache.clear)

    def test_capabilities_are_cached_per_model(self):
        with patch.object(ai_chat.requests, "post", return_value=response({"capabilities": ["completion", "tools"]})) as post:
            self.assertTrue(ai_chat.supports_tools("m1"))
            self.assertTrue(ai_chat.supports_tools("m1"))
        self.assertEqual(post.call_count, 1)

    def test_unreachable_ollama_is_not_cached(self):
        with patch.object(ai_chat.requests, "post", side_effect=ai_chat.requests.exceptions.ConnectionError()):
            self.assertFalse(ai_chat.supports_tools("m2"))
        self.assertNotIn((ai_chat.OLLAMA_HOST, "m2"), ai_chat._model_capabilities_cache)

    def test_model_without_tools_is_rejected(self):
        tags = response({"models": [{"name": "old:1b"}]})
        with patch.object(ai_chat, "OLLAMA_MODEL", "old:1b"), \
                patch.object(ai_chat.requests, "get", return_value=tags), \
                patch.object(ai_chat.requests, "post", return_value=response({"capabilities": ["completion"]})):
            ok, message = ai_chat.check_ollama_model()
        self.assertFalse(ok)
        self.assertIn("не умеет вызывать инструменты", message)

    def test_missing_model(self):
        with patch.object(ai_chat, "OLLAMA_MODEL", "nope"), \
                patch.object(ai_chat.requests, "get", return_value=response({"models": [{"name": "qwen3.5:9b"}]})):
            ok, message = ai_chat.check_ollama_model()
        self.assertFalse(ok)
        self.assertIn("ollama pull nope", message)


if __name__ == "__main__":
    unittest.main()
