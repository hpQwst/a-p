from __future__ import annotations

from dataclasses import dataclass
from typing import Any
import json
import time

from .ai import build_openai_client, reasoning_for_operation
from .ai_debug import log_ai_request, log_ai_response, log_debug_event


@dataclass(frozen=True)
class SlideMatrixBuildInput:
    slide_number: int
    slide_understanding: dict[str, Any]
    targets: list[dict[str, Any]]
    xlsx_manifests: list[dict[str, Any]]
    xlsx_dumps: list[str]
    target_ids: list[str] | None = None
    manual_context: str = ""


def build_slide_matrices_with_ai(payload: SlideMatrixBuildInput, root: Path | str | None = None) -> dict[str, Any]:
    client, model = build_openai_client(root, operation="slide_matrix_builder")
    user_payload = {
        "slide_number": payload.slide_number,
        "slide_understanding": payload.slide_understanding,
        "targets": payload.targets,
        "target_ids_to_process": payload.target_ids or [],
        "xlsx_manifests": payload.xlsx_manifests,
        "xlsx_plaintext_dumps": payload.xlsx_dumps,
        "manual_context": payload.manual_context,
        "rules": [
            "Escolha a fonte e monte final_edit_data para cada target solicitado usando somente valores do dump XLSX.",
            "Preserve a ordem e os labels completos do XLSX; use o PPT apenas para a orientacao fisica e capacidade da matriz.",
            "series_rows_categories_columns: headers=[blank,categorias...] e rows=[[serie,valores...],...].",
            "categories_rows_series_columns: headers=[blank,series...] e rows=[[categoria,valores...],...].",
            "table_cells: rows na mesma ordem da fonte; mantenha linhas fixas do template apenas quando nao existirem no XLSX.",
            "Labels, periodos e codigos sao type=text; valores quantitativos sao type=number. Nao arredonde.",
            "source_trace deve resumir o range usado, sem repetir cada celula.",
        ],
    }
    content: list[dict[str, Any]] = [{"type": "input_text", "text": json.dumps(user_payload, ensure_ascii=False)}]
    request_kwargs = {
        "model": model,
        "store": False,
        "reasoning": reasoning_for_operation("slide_matrix_builder"),
        "input": [
            {
                "role": "system",
                "content": (
                    "Monte matrizes tipadas para o Editar dados do PowerPoint a partir dos dumps XLSX. "
                    "Nao invente valores; preserve ordem, nomes e precisao da fonte."
                ),
            },
            {"role": "user", "content": content},
        ],
        "text": {
            "format": {
                "type": "json_schema",
                "name": "slide_target_matrices",
                "schema": _schema(),
                "strict": True,
            }
        },
    }
    log_ai_request("slide_matrix_builder", request_kwargs)
    started = time.perf_counter()
    try:
        response = client.responses.create(**request_kwargs)
    except Exception as exc:
        log_debug_event(
            "ai_error",
            {
                "operation": "slide_matrix_builder",
                "elapsed_ms": round((time.perf_counter() - started) * 1000),
                "error": repr(exc),
            },
        )
        raise
    text = getattr(response, "output_text", "") or _response_text_fallback(response)
    log_ai_response("slide_matrix_builder", text, round((time.perf_counter() - started) * 1000), response)
    return json.loads(text)


def _schema() -> dict[str, Any]:
    typed_cell = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "value": {"type": "string"},
            "type": {"type": "string", "enum": ["text", "number"]},
        },
        "required": ["value", "type"],
    }
    return {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "slide_number": {"type": "integer"},
            "target_outputs": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "target_id": {"type": "string"},
                        "object_type": {"type": "string", "enum": ["chart", "table", "text"]},
                        "target_name": {"type": "string"},
                        "source_file": {"type": "string"},
                        "source_part": {"type": "string"},
                        "edit_orientation": {
                            "type": "string",
                            "enum": ["series_rows_categories_columns", "categories_rows_series_columns", "table_cells", "text_value"],
                        },
                        "visual_number_format": {"type": "string"},
                        "confidence": {"type": "number", "minimum": 0, "maximum": 1},
                        "extraction_explanation": {"type": "string"},
                        "final_edit_data": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "headers": {"type": "array", "items": typed_cell},
                                "rows": {
                                    "type": "array",
                                    "items": {"type": "array", "items": typed_cell},
                                },
                            },
                            "required": ["headers", "rows"],
                        },
                        "source_trace": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "additionalProperties": False,
                                "properties": {
                                    "source_file": {"type": "string"},
                                    "source_sheet": {"type": "string"},
                                    "source_range": {"type": "string"},
                                },
                                "required": ["source_file", "source_sheet", "source_range"],
                            },
                        },
                    },
                    "required": [
                        "target_id",
                        "object_type",
                        "target_name",
                        "source_file",
                        "source_part",
                        "edit_orientation",
                        "visual_number_format",
                        "confidence",
                        "extraction_explanation",
                        "final_edit_data",
                        "source_trace",
                    ],
                },
            },
            "questions_for_user": {"type": "array", "items": {"type": "string"}},
        },
        "required": ["slide_number", "target_outputs", "questions_for_user"],
    }


def _response_text_fallback(response: Any) -> str:
    try:
        return response.output[0].content[0].text
    except Exception:
        return str(response)
