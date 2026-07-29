from __future__ import annotations

from dataclasses import dataclass
from typing import Any
import json
import time

from .ai import build_openai_client, reasoning_for_operation
from .ai_debug import log_ai_request, log_ai_response, log_debug_event


@dataclass(frozen=True)
class SlideUnderstandingInput:
    slide_number: int
    slide_text: str
    targets: list[dict[str, Any]]
    xlsx_manifests: list[dict[str, Any]]
    xlsx_dumps: list[str]
    manual_context: str = ""


def suggest_slide_understanding(payload: SlideUnderstandingInput, root: Path | str | None = None) -> dict[str, Any]:
    client, model = build_openai_client(root, operation="slide_understanding")
    user_payload = {
        "slide_number": payload.slide_number,
        "slide_text": payload.slide_text,
        "targets": payload.targets,
        "xlsx_manifests": payload.xlsx_manifests,
        "xlsx_plaintext_dumps": payload.xlsx_dumps,
        "manual_context": payload.manual_context,
        "instructions": [
            "Interprete o papel de cada target pelo contrato OpenXML, titulos e contexto textual do slide.",
            "Use xlsx_manifests como indice semantico principal dos arquivos; use xlsx_plaintext_dumps para conferir celulas, coordenadas e valores raw.",
            "Interprete os XLSX somente pelo dump textual compacto/JSON. Nao use rasterizacao de planilha.",
            "Nao invente valores. Nesta etapa, apenas descreva partes uteis e como extrair.",
            "Se o arquivo nao tiver dados suficientes, retorne questions_for_user.",
        ],
    }
    content: list[dict[str, Any]] = [{"type": "input_text", "text": json.dumps(user_payload, ensure_ascii=False)}]
    request_kwargs = {
        "model": model,
        "store": False,
        "reasoning": reasoning_for_operation("slide_understanding"),
        "input": [
            {
                "role": "system",
                "content": (
                    "Voce entende slides de PowerPoint e planilhas complexas. "
                    "Seu trabalho e criar um entendimento por slide, mapeando alvos visuais para partes uteis dos XLSX. "
                    "Use valores raw apenas como referencia textual; nao crie matriz final nesta etapa."
                ),
            },
            {"role": "user", "content": content},
        ],
        "text": {
            "format": {
                "type": "json_schema",
                "name": "slide_understanding",
                "schema": _schema(),
                "strict": True,
            }
        },
    }
    log_ai_request("slide_understanding", request_kwargs)
    started = time.perf_counter()
    try:
        response = client.responses.create(**request_kwargs)
    except Exception as exc:
        log_debug_event(
            "ai_error",
            {
                "operation": "slide_understanding",
                "elapsed_ms": round((time.perf_counter() - started) * 1000),
                "error": repr(exc),
            },
        )
        raise
    text = getattr(response, "output_text", "") or _response_text_fallback(response)
    log_ai_response("slide_understanding", text, round((time.perf_counter() - started) * 1000), response)
    return json.loads(text)


def _schema() -> dict[str, Any]:
    return {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "slide_number": {"type": "integer"},
            "slide_understanding": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "title": {"type": "string"},
                    "main_metric": {"type": "string"},
                    "targets": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "target_id": {"type": "string"},
                                "visual_label": {"type": "string"},
                                "semantic_name": {"type": "string"},
                                "role": {"type": "string"},
                            },
                            "required": ["target_id", "visual_label", "semantic_name", "role"],
                        },
                    },
                },
                "required": ["title", "main_metric", "targets"],
            },
            "xlsx_understanding": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "file_name": {"type": "string"},
                        "meaning": {"type": "string"},
                        "usable_parts": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "additionalProperties": False,
                                "properties": {
                                    "part_id": {"type": "string"},
                                    "meaning": {"type": "string"},
                                    "how_to_extract": {"type": "string"},
                                },
                                "required": ["part_id", "meaning", "how_to_extract"],
                            },
                        },
                    },
                    "required": ["file_name", "meaning", "usable_parts"],
                },
            },
            "questions_for_user": {"type": "array", "items": {"type": "string"}},
        },
        "required": ["slide_number", "slide_understanding", "xlsx_understanding", "questions_for_user"],
    }


def _response_text_fallback(response: Any) -> str:
    try:
        return response.output[0].content[0].text
    except Exception:
        return str(response)
