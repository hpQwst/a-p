from __future__ import annotations

from contextlib import contextmanager
from contextvars import ContextVar
from datetime import datetime
import json
from pathlib import Path
from typing import Any, Iterator


_CURRENT_LOG_PATH: ContextVar[str] = ContextVar("auto_ppt_ai_debug_log_path", default="")
_CURRENT_USAGE_PATH: ContextVar[str] = ContextVar("auto_ppt_ai_usage_log_path", default="")
_CURRENT_REQUEST_META: ContextVar[dict[str, Any]] = ContextVar("auto_ppt_ai_request_meta", default={})


@contextmanager
def ai_debug_log(path: Path | str | None) -> Iterator[None]:
    token = _CURRENT_LOG_PATH.set(str(path or ""))
    try:
        yield
    finally:
        _CURRENT_LOG_PATH.reset(token)


def current_ai_debug_log_path() -> Path | None:
    value = _CURRENT_LOG_PATH.get()
    return Path(value) if value else None


def set_ai_debug_log_path(path: Path | str | None):
    return _CURRENT_LOG_PATH.set(str(path or ""))


def reset_ai_debug_log_path(token) -> None:
    _CURRENT_LOG_PATH.reset(token)


def set_ai_usage_log_path(path: Path | str | None):
    return _CURRENT_USAGE_PATH.set(str(path or ""))


def reset_ai_usage_log_path(token) -> None:
    _CURRENT_USAGE_PATH.reset(token)


def log_debug_event(event: str, payload: dict[str, Any] | None = None) -> None:
    path = current_ai_debug_log_path()
    event_payload = payload or {}
    if path is not None:
        _append_jsonl(path, {"event": event, **event_payload})
    if event == "ai_error":
        _log_compact_error(event_payload)


def log_ai_request(operation: str, request_payload: dict[str, Any]) -> None:
    _CURRENT_REQUEST_META.set(
        {
            "operation": operation,
            "model_requested": str(request_payload.get("model") or ""),
            "reasoning_effort": str((request_payload.get("reasoning") or {}).get("effort") or ""),
            "input_bytes_utf8": _json_size(request_payload.get("input") or []),
            "request_bytes_utf8": _json_size(request_payload),
            "sent_at": datetime.now().isoformat(timespec="milliseconds"),
        }
    )
    log_debug_event(
        "ai_request",
        {
            "operation": operation,
            "payload_bytes_utf8": _json_size(request_payload),
            "input": request_payload.get("input"),
            "text": request_payload.get("text"),
            "request": request_payload,
        },
    )


def log_ai_response(
    operation: str,
    output_text: str,
    elapsed_ms: int | None = None,
    response: Any | None = None,
) -> None:
    response_payload = _serialize_response(response)
    log_debug_event(
        "ai_response",
        {
            "operation": operation,
            "elapsed_ms": elapsed_ms,
            "output_bytes_utf8": len((output_text or "").encode("utf-8")),
            "output": {
                "output_text": output_text,
                "response": response_payload,
            },
            "output_text": output_text,
            "response": response_payload,
        },
    )
    _log_compact_usage(operation, output_text, elapsed_ms, response, response_payload)


def _append_jsonl(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    record = {"timestamp": datetime.now().isoformat(timespec="milliseconds"), **payload}
    with path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(record, ensure_ascii=False, default=str))
        handle.write("\n")


def _json_size(value: Any) -> int:
    return len(json.dumps(value, ensure_ascii=False, default=str, separators=(",", ":")).encode("utf-8"))


def _serialize_response(response: Any) -> Any:
    if response is None:
        return None
    for method_name in ("model_dump", "to_dict", "dict"):
        method = getattr(response, method_name, None)
        if callable(method):
            try:
                return method()
            except Exception:
                pass
    return repr(response)


def _log_compact_usage(
    operation: str,
    output_text: str,
    elapsed_ms: int | None,
    response: Any,
    response_payload: Any,
) -> None:
    value = _CURRENT_USAGE_PATH.get()
    if not value:
        return
    request_meta = _CURRENT_REQUEST_META.get()
    payload = response_payload if isinstance(response_payload, dict) else {}
    usage = payload.get("usage") if isinstance(payload.get("usage"), dict) else {}
    if not usage:
        serialized_usage = _serialize_response(getattr(response, "usage", None))
        usage = serialized_usage if isinstance(serialized_usage, dict) else {}
    input_details = usage.get("input_tokens_details") or {}
    output_details = usage.get("output_tokens_details") or {}
    input_tokens = _int_or_zero(usage.get("input_tokens"))
    cached_input_tokens = _int_or_zero(input_details.get("cached_tokens"))
    output_tokens = _int_or_zero(usage.get("output_tokens"))
    total_tokens = _int_or_zero(usage.get("total_tokens")) or input_tokens + output_tokens
    model_returned = str(payload.get("model") or getattr(response, "model", "") or "")
    model_requested = str(request_meta.get("model_requested") or "")
    record = {
        "operation": operation,
        "status": "ok",
        "sent_at": request_meta.get("sent_at") or "",
        "returned_at": datetime.now().isoformat(timespec="milliseconds"),
        "duration_ms": elapsed_ms,
        "model_requested": model_requested,
        "model_returned": model_returned,
        "reasoning_effort": request_meta.get("reasoning_effort") or "",
        "input_bytes_utf8": request_meta.get("input_bytes_utf8") or 0,
        "request_bytes_utf8": request_meta.get("request_bytes_utf8") or 0,
        "output_bytes_utf8": len((output_text or "").encode("utf-8")),
        "input_tokens": input_tokens,
        "cached_input_tokens": cached_input_tokens,
        "output_tokens": output_tokens,
        "reasoning_tokens": _int_or_zero(output_details.get("reasoning_tokens")),
        "total_tokens": total_tokens,
        "estimated_cost_usd": _estimated_cost_usd(
            model_returned or model_requested,
            input_tokens,
            cached_input_tokens,
            output_tokens,
        ),
    }
    _append_jsonl(Path(value), record)


def _log_compact_error(payload: dict[str, Any]) -> None:
    value = _CURRENT_USAGE_PATH.get()
    if not value:
        return
    request_meta = _CURRENT_REQUEST_META.get()
    raw_error = str(payload.get("error") or "")
    error_type = raw_error.split("(", 1)[0].strip()[:80]
    _append_jsonl(
        Path(value),
        {
            "operation": payload.get("operation") or request_meta.get("operation") or "",
            "status": "error",
            "sent_at": request_meta.get("sent_at") or "",
            "returned_at": datetime.now().isoformat(timespec="milliseconds"),
            "duration_ms": payload.get("elapsed_ms"),
            "model_requested": request_meta.get("model_requested") or "",
            "reasoning_effort": request_meta.get("reasoning_effort") or "",
            "input_bytes_utf8": request_meta.get("input_bytes_utf8") or 0,
            "request_bytes_utf8": request_meta.get("request_bytes_utf8") or 0,
            "error_type": error_type,
        },
    )


def _int_or_zero(value: Any) -> int:
    try:
        return max(int(value or 0), 0)
    except (TypeError, ValueError):
        return 0


def _estimated_cost_usd(
    model: str,
    input_tokens: int,
    cached_input_tokens: int,
    output_tokens: int,
) -> float | None:
    # USD por 1M tokens. A telemetria mantem tokens brutos para permitir
    # recalculo caso a tabela publica de precos mude.
    prices = {
        "gpt-5.6-luna": (1.00, 0.10, 6.00),
        "gpt-5.6-terra": (2.50, 0.25, 15.00),
    }
    normalized = model.strip().lower()
    key = next((name for name in prices if normalized.startswith(name)), "")
    if not key:
        return None
    input_rate, cached_rate, output_rate = prices[key]
    uncached_tokens = max(input_tokens - cached_input_tokens, 0)
    cost = (
        uncached_tokens * input_rate
        + cached_input_tokens * cached_rate
        + output_tokens * output_rate
    ) / 1_000_000
    return round(cost, 8)
