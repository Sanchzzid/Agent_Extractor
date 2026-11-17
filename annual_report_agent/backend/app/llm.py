
from __future__ import annotations

import base64
import json
import re
import threading
import time
from functools import lru_cache
from typing import Any, Callable

from json_repair import repair_json
from langchain_core.messages import AIMessage, BaseMessage, HumanMessage, SystemMessage
from langchain_openai import ChatOpenAI

from . import config, prompts

_SEMAPHORE = threading.BoundedSemaphore(max(1, config.LLM_MAX_CONCURRENCY))
_USAGE_LOCK = threading.Lock()


class LLMError(RuntimeError):
    """The model could not produce a usable answer after all retries."""


class ContextTooLarge(LLMError):
    """The request (or the answer) does not fit in the model's context window."""


class OutputTooLong(LLMError):
    """The model hit max_tokens (usually a repetition loop or far too many rows)."""


def estimate_tokens(text: str) -> int:
    # Conservative: numeric tables tokenize worse than prose.
    return int(len(text) / 2.8) + 1


@lru_cache(maxsize=64)
def _chat(max_tokens: int, temperature: float) -> ChatOpenAI:
    config.check_llm_settings()
    return ChatOpenAI(
        model=config.LLM_MODEL,
        base_url=config.LLM_ENDPOINT or None,
        api_key=config.LLM_API_KEY,
        temperature=temperature,
        max_tokens=max_tokens,
        timeout=config.LLM_TIMEOUT_SECONDS,
        max_retries=0,  # retries are handled below so we can classify errors
        model_kwargs={"response_format": {"type": "json_object"}} if config.LLM_JSON_MODE else {},
    )


def image_part(png_bytes: bytes) -> dict[str, Any]:
    encoded = base64.b64encode(png_bytes).decode("ascii")
    return {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{encoded}"}}


def parse_json_object(text: str) -> dict[str, Any]:
    """Parse model output into a dict, tolerating fences and small syntax errors."""
    raw = (text or "").strip()
    if not raw:
        raise ValueError("empty reply")
    raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.I).strip()
    try:
        value = json.loads(raw)
    except json.JSONDecodeError:
        start, end = raw.find("{"), raw.rfind("}")
        candidate = raw[start:end + 1] if start != -1 and end > start else raw
        try:
            value = json.loads(candidate)
        except json.JSONDecodeError:
            value = repair_json(candidate, return_objects=True)
    if not isinstance(value, dict):
        raise ValueError("reply is not a JSON object")
    return value


def _is_context_error(exc: Exception) -> bool:
    msg = str(exc).lower()
    return "maximum context length" in msg or "context length" in msg or "too many tokens" in msg


def _is_retryable(exc: Exception) -> bool:
    msg = str(exc).lower()
    name = type(exc).__name__.lower()
    return any(s in msg or s in name for s in (
        "timeout", "timed out", "connection", "rate", "429", "500", "502", "503", "504",
        "overloaded", "temporarily", "server error", "apiconnection", "internalserver"))


def _invoke(messages: list[BaseMessage], max_tokens: int, temperature: float) -> AIMessage:
    delay = 2.0
    max_tokens = max(200, (max_tokens // 250) * 250)  # bucket so client objects are reused
    for attempt in range(config.LLM_NETWORK_RETRIES + 1):
        try:
            with _SEMAPHORE:
                return _chat(max_tokens, temperature).invoke(messages)
        except Exception as exc:  # noqa: BLE001 - classify any SDK error
            if "lengthfinishreason" in type(exc).__name__.lower() or "length limit" in str(exc).lower():
                raise OutputTooLong("reply hit the output token limit") from exc
            if _is_context_error(exc):
                raise ContextTooLarge(str(exc)[:300]) from exc
            if attempt >= config.LLM_NETWORK_RETRIES or not _is_retryable(exc):
                raise LLMError(f"{type(exc).__name__}: {str(exc)[:300]}") from exc
            time.sleep(delay)
            delay = min(delay * 2, 30)
    raise LLMError("unreachable")


def _output_budget(messages: list[BaseMessage], image_count: int) -> int:
    used = sum(estimate_tokens(m.content if isinstance(m.content, str) else
                               " ".join(p.get("text", "") for p in m.content if isinstance(p, dict)))
               for m in messages)
    used += image_count * config.IMAGE_TOKEN_COST
    room = config.LLM_CONTEXT_TOKENS - used - 500
    if room < 600:
        raise ContextTooLarge(f"prompt too large (~{used} tokens)")
    return min(config.LLM_MAX_OUTPUT_TOKENS, room)


def call_json(prompt: str,
              validate: Callable[[dict[str, Any]], dict[str, Any]] | None = None,
              images: list[bytes] | None = None,
              usage: dict[str, int] | None = None,
              max_output: int | None = None) -> dict[str, Any]:
    """Ask the model for JSON; re-ask with the error until it parses and validates.

    max_output caps the reply size (small answers like page lists or verifier issues);
    a reply that runs over the cap is retried at a slightly higher temperature, which
    breaks the repetition loops some models fall into at temperature 0.
    """
    images = images or []
    content: Any = prompt
    if images:
        content = [{"type": "text", "text": prompt}] + [image_part(b) for b in images]
    messages: list[BaseMessage] = [SystemMessage(prompts.SYSTEM_JSON), HumanMessage(content=content)]

    last_error = "unknown error"
    too_long = False
    for attempt in range(config.LLM_JSON_RETRIES + 1):
        max_tokens = _output_budget(messages, len(images))
        if max_output:
            max_tokens = min(max_tokens, max_output)
        temperature = 0.0 if attempt == 0 else 0.3
        try:
            reply = _invoke(messages, max_tokens, temperature)
        except OutputTooLong as exc:
            too_long, last_error = True, str(exc)
            continue
        _record_usage(usage, reply)
        text = reply.content if isinstance(reply.content, str) else str(reply.content)
        finish = (reply.response_metadata or {}).get("finish_reason")
        try:
            if finish == "length":
                too_long = True
                raise ValueError("reply was cut off (too long); answer more concisely")
            too_long = False
            data = parse_json_object(text)
            return validate(data) if validate else data
        except Exception as exc:  # noqa: BLE001 - parse / validation failure
            last_error = f"{type(exc).__name__}: {exc}"[:400]
            # Keep the conversation short: original prompt + bad reply + repair request.
            messages = messages[:2] + [AIMessage(content=text[:4000]),
                                       HumanMessage(content=prompts.json_repair_prompt(last_error))]
    if too_long:  # caller may split the batch into smaller pieces
        raise ContextTooLarge(f"reply too long after {config.LLM_JSON_RETRIES + 1} attempts")
    raise LLMError(f"invalid JSON after {config.LLM_JSON_RETRIES + 1} attempts: {last_error}")


def _record_usage(usage: dict[str, int] | None, reply: AIMessage) -> None:
    if usage is None:
        return
    meta = getattr(reply, "usage_metadata", None) or {}
    with _USAGE_LOCK:
        usage["calls"] = usage.get("calls", 0) + 1
        usage["input_tokens"] = usage.get("input_tokens", 0) + int(meta.get("input_tokens", 0) or 0)
        usage["output_tokens"] = usage.get("output_tokens", 0) + int(meta.get("output_tokens", 0) or 0)
