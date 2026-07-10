from __future__ import annotations

from dataclasses import dataclass
from typing import Any
import json
import time

from .ai import build_openai_client
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
            "Monte final_edit_data tipado para cada target solicitado.",
            "Use xlsx_manifests como indice semantico principal dos arquivos; use xlsx_plaintext_dumps para conferir celulas, coordenadas e valores raw.",
            "Use somente valores presentes no raw dos XLSX. Nunca invente valores.",
            "Nunca arredonde valores numericos. Copie source_raw exatamente.",
            "Use os labels, categorias, series e headers do XLSX sempre que eles forem a fonte correspondente; nao substitua por labels resumidos ou antigos do PPT.",
            "O PPT serve para entender o papel visual e o formato esperado, mas o texto final dos rotulos deve vir do XLSX quando houver match.",
            "Para object_type=chart, final_edit_data deve estar no formato exato do workbook do PowerPoint, nao em uma orientacao semantica livre.",
            "Para chart com expected_orientation=series_rows_categories_columns, use headers=[blank, categorias...] e rows=[[serie, valores...], ...].",
            "Para chart com expected_orientation=categories_rows_series_columns, use headers=[blank, series...] e rows=[[categoria, valores...], ...].",
            "Para object_type=table, final_edit_data deve espelhar as celulas visiveis da tabela PowerPoint. Preserve labels existentes e preencha as celulas de valor correspondentes.",
            "Para tabelas chave-valor, use a primeira coluna como chave/label e coloque o valor correspondente na segunda coluna.",
            "Retorne edit_orientation com a orientacao usada: series_rows_categories_columns, categories_rows_series_columns ou table_cells.",
            "Retorne visual_number_format quando a formatacao visual for importante, por exemplo 0.0 para mostrar uma casa decimal sem arredondar o workbook.",
            "Headers, series, categorias, periodos, codigos e labels devem ser type=text e force_text=true.",
            "Valores quantitativos devem ser type=number.",
            "Labels como Nov/25, Dez/25, 1Q26, 01/2026, 001 e 10-15 sao texto forcado.",
            "Inclua source_trace quando souber a celula de origem.",
        ],
    }
    content: list[dict[str, Any]] = [{"type": "input_text", "text": json.dumps(user_payload, ensure_ascii=False)}]
    request_kwargs = {
        "model": model,
        "store": False,
        "input": [
            {
                "role": "system",
                "content": (
                    "Voce monta matrizes finais para o Editar dados do PowerPoint. "
                    "A saida deve ser exata, tipada e rastreavel. "
                    "Voce nao interpreta valores por rasterizacao de planilha; usa apenas dumps textuais compactos dos XLSX. "
                    "Quando um label do XLSX for mais completo que o label atual do PPT, use o label do XLSX na matriz final."
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
            "force_text": {"type": "boolean"},
            "source_raw": {"type": "string"},
        },
        "required": ["value", "type", "force_text", "source_raw"],
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
                                "matrix_preview": {
                                    "type": "array",
                                    "items": {"type": "array", "items": {"type": "string"}},
                                },
                                "headers": {"type": "array", "items": typed_cell},
                                "rows": {
                                    "type": "array",
                                    "items": {"type": "array", "items": typed_cell},
                                },
                            },
                            "required": ["matrix_preview", "headers", "rows"],
                        },
                        "source_trace": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "additionalProperties": False,
                                "properties": {
                                    "output_position": {"type": "string"},
                                    "value": {"type": "string"},
                                    "source_file": {"type": "string"},
                                    "source_sheet": {"type": "string"},
                                    "source_cell": {"type": "string"},
                                },
                                "required": ["output_position", "value", "source_file", "source_sheet", "source_cell"],
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
