# -*- coding: utf-8 -*-
"""
Клиент локального Ollama для AI-ассистента: настройки, возможности модели, один запрос к /api/chat.

Цикл «вопрос → инструменты → ответ» — в `ai_agent.py`, инструменты — в `ai_tools.py`.
"""

import json
import os
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional, Tuple

import requests
from dotenv import load_dotenv

load_dotenv()
OLLAMA_HOST = (os.getenv("OLLAMA_HOST") or "http://localhost:11434").rstrip("/")
OLLAMA_CHAT_URL = f"{OLLAMA_HOST}/api/chat"


def _env_int(name: str, default: int) -> int:
    try:
        return int((os.getenv(name) or "").strip() or default)
    except ValueError:
        return default


DEFAULT_MODEL = "qwen3:4b"
# Приложение меняет модель на лету: ai_chat.OLLAMA_MODEL = "<имя>".
OLLAMA_MODEL = (os.getenv("OLLAMA_MODEL") or "").strip() or DEFAULT_MODEL
# Окно контекста. Без явного num_ctx Ollama берёт 4096 токенов и молча обрезает
# всё, что не поместилось, — данные инструментов терялись.
OLLAMA_NUM_CTX = _env_int("OLLAMA_NUM_CTX", 32768)
# Сколько модель остаётся в памяти после ответа: следующий вопрос не ждёт загрузки.
OLLAMA_KEEP_ALIVE = (os.getenv("OLLAMA_KEEP_ALIVE") or "").strip() or "30m"
# Режим размышлений: "1" — включить, "0" — выключить, пусто — как задано в модели.
OLLAMA_THINK = (os.getenv("OLLAMA_THINK") or "").strip()
# Таймаут чтения ответа: большой контекст на ноутбуке читается минутами.
OLLAMA_READ_TIMEOUT = _env_int("OLLAMA_READ_TIMEOUT", 600)

_model_capabilities_cache: Dict[Tuple[str, str], List[str]] = {}


def model_capabilities(model: Optional[str] = None) -> Optional[List[str]]:
    """Возможности модели из /api/show (tools, thinking, ...); None — Ollama не ответил."""
    model = model or OLLAMA_MODEL
    key = (OLLAMA_HOST, model)
    if key in _model_capabilities_cache:
        return _model_capabilities_cache[key]
    try:
        response = requests.post(f"{OLLAMA_HOST}/api/show", json={"model": model}, timeout=5)
        response.raise_for_status()
        capabilities = response.json().get("capabilities")
    except (requests.exceptions.RequestException, ValueError):
        return None
    result = [str(item) for item in capabilities] if isinstance(capabilities, list) else []
    _model_capabilities_cache[key] = result
    return result


def supports_tools(model: Optional[str] = None) -> bool:
    """Ассистент работает только через встроенный вызов инструментов."""
    return "tools" in (model_capabilities(model) or [])


def check_ollama_model() -> Tuple[bool, str]:
    """Ollama доступен, выбранная модель установлена и умеет вызывать инструменты."""
    try:
        response = requests.get(f"{OLLAMA_HOST}/api/tags", timeout=5)
        response.raise_for_status()
        data = response.json()
    except requests.exceptions.RequestException:
        return False, f"Ollama недоступен по адресу {OLLAMA_HOST}"
    available = {item.get("name") for item in data.get("models", []) or [] if isinstance(item, dict) and item.get("name")}
    if OLLAMA_MODEL not in available:
        return False, f"Модель {OLLAMA_MODEL} не найдена в Ollama. Скачайте её командой: ollama pull {OLLAMA_MODEL}"
    capabilities = model_capabilities()
    if capabilities is not None and "tools" not in capabilities:
        return False, (
            f"Модель {OLLAMA_MODEL} не умеет вызывать инструменты, а без них ассистент не видит данных. "
            "Выберите модель с поддержкой tools, например qwen3.5:9b или qwen3:4b."
        )
    return True, ""


def _build_ollama_payload(
    messages: List[Dict[str, Any]],
    temperature: float,
    max_tokens: Optional[int],
    stream: bool,
    tools: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    payload: Dict[str, Any] = {
        "model": OLLAMA_MODEL,
        "messages": messages,
        "stream": stream,
        "keep_alive": OLLAMA_KEEP_ALIVE,
        "options": {
            "temperature": temperature,
            "num_ctx": OLLAMA_NUM_CTX,
        },
    }
    if max_tokens is not None:
        payload["options"]["num_predict"] = max_tokens
    if tools:
        payload["tools"] = tools
    if OLLAMA_THINK in ("0", "1") and "thinking" in (model_capabilities() or []):
        payload["think"] = OLLAMA_THINK == "1"
    return payload


def _context_overflow(prompt_eval_count: Any) -> bool:
    """Запрос занял почти всё окно — Ollama мог обрезать начало (системный промпт, данные)."""
    return isinstance(prompt_eval_count, int) and prompt_eval_count >= OLLAMA_NUM_CTX * 0.95


@dataclass
class ChatResponse:
    content: str = ""
    thinking: str = ""
    tool_calls: List[Dict[str, Any]] = field(default_factory=list)
    prompt_tokens: Optional[int] = None
    output_tokens: Optional[int] = None
    done_reason: Optional[str] = None
    model: Optional[str] = None
    error: Optional[str] = None
    cancelled: bool = False

    @property
    def context_overflow(self) -> bool:
        return _context_overflow(self.prompt_tokens)

    def runtime(self) -> Dict[str, Any]:
        return {
            "model": self.model, "host": OLLAMA_HOST, "num_ctx": OLLAMA_NUM_CTX, "done_reason": self.done_reason,
            "prompt_tokens": self.prompt_tokens, "output_tokens": self.output_tokens,
            "context_overflow": self.context_overflow, "error": self.error,
        }


def parse_tool_calls(tool_calls: List[Dict[str, Any]]) -> List[Tuple[str, Dict[str, Any]]]:
    """tool_calls Ollama → [(имя, аргументы)]; аргументы-строки (JSON) раскрываются."""
    result = []
    for call in tool_calls or []:
        function = call.get("function") if isinstance(call, dict) else None
        if not isinstance(function, dict) or not function.get("name"):
            continue
        args = function.get("arguments")
        if isinstance(args, str):
            try:
                args = json.loads(args)
            except json.JSONDecodeError:
                args = {}
        result.append((str(function["name"]), args if isinstance(args, dict) else {}))
    return result


def chat(
    messages: List[Dict[str, Any]],
    tools: Optional[List[Dict[str, Any]]] = None,
    temperature: float = 0.3,
    max_tokens: Optional[int] = None,
    on_chunk: Optional[Callable[[str], None]] = None,
    on_thinking: Optional[Callable[[str], None]] = None,
    should_cancel: Optional[Callable[[], bool]] = None,
) -> ChatResponse:
    """
    Один потоковый запрос к Ollama. Текст ответа и размышления отдаются по кусочкам в on_chunk / on_thinking.
    Ошибка сети или Ollama — в ChatResponse.error, исключений нет.
    """
    result = ChatResponse(model=OLLAMA_MODEL)
    content: List[str] = []
    thinking: List[str] = []
    payload = _build_ollama_payload(messages, temperature, max_tokens, stream=True, tools=tools)
    try:
        with requests.post(OLLAMA_CHAT_URL, json=payload, timeout=(10, OLLAMA_READ_TIMEOUT), stream=True) as response:
            if response.status_code >= 400:
                try:
                    detail = response.json().get("error") or response.text
                except ValueError:
                    detail = response.text
                result.error = f"Ollama ответил {response.status_code}: {str(detail)[:300]}"
                return result
            for raw_line in response.iter_lines(decode_unicode=True):
                if should_cancel and should_cancel():
                    result.cancelled = True
                    return result
                if not raw_line:
                    continue
                data = json.loads(raw_line)
                if data.get("error"):
                    result.error = f"Ollama: {data['error']}"
                    return result
                message = data.get("message") if isinstance(data.get("message"), dict) else {}
                chunk = message.get("content")
                if isinstance(chunk, str) and chunk:
                    content.append(chunk)
                    if on_chunk:
                        on_chunk(chunk)
                thought = message.get("thinking")
                if isinstance(thought, str) and thought:
                    thinking.append(thought)
                    if on_thinking:
                        on_thinking(thought)
                if message.get("tool_calls"):
                    result.tool_calls.extend(message["tool_calls"])
                if data.get("done"):
                    result.prompt_tokens = data.get("prompt_eval_count")
                    result.output_tokens = data.get("eval_count")
                    result.done_reason = data.get("done_reason")
                    result.model = data.get("model") or result.model
                    break
    except requests.exceptions.RequestException as exc:
        result.error = f"Нет связи с Ollama ({OLLAMA_HOST}): {exc.__class__.__name__}"
    except ValueError as exc:
        result.error = f"Ollama прислал непонятный ответ: {exc}"
    result.content = "".join(content).strip()
    result.thinking = "".join(thinking).strip()
    return result
