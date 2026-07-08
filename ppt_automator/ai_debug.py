from __future__ import annotations

from contextlib import contextmanager
from contextvars import ContextVar
from datetime import datetime
import json
from pathlib import Path
from typing import Any, Iterator


_CURRENT_LOG_PATH: ContextVar[str] = ContextVar("auto_ppt_ai_debug_log_path", default="")


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


def log_debug_event(event: str, payload: dict[str, Any] | None = None) -> None:
    path = current_ai_debug_log_path()
    if path is None:
        return
    _append_jsonl(path, {"event": event, **(payload or {})})


def log_ai_request(operation: str, request_payload: dict[str, Any]) -> None:
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
