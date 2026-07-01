from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any
import base64
import json

from .ai import build_openai_client


@dataclass(frozen=True)
class SlideMatrixBuildInput:
    slide_number: int
    slide_image_path: Path | None
    visual_map: list[dict[str, str]]
    slide_understanding: dict[str, Any]
    targets: list[dict[str, Any]]
    xlsx_dumps: list[str]
    target_ids: list[str] | None = None
    manual_context: str = ""


def build_slide_matrices_with_ai(payload: SlideMatrixBuildInput, root: Path | str | None = None) -> dict[str, Any]:
    client, model = build_openai_client(root)
    user_payload = {
        "slide_number": payload.slide_number,
        "visual_target_map": payload.visual_map,
        "slide_understanding": payload.slide_understanding,
        "targets": payload.targets,
        "target_ids_to_process": payload.target_ids or [],
        "xlsx_plaintext_dumps": payload.xlsx_dumps,
        "manual_context": payload.manual_context,
        "rules": [
            "Monte final_edit_data tipado para cada target solicitado.",
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
    if payload.slide_image_path and payload.slide_image_path.exists():
        content.insert(0, {"type": "input_image", "image_url": _image_data_url(payload.slide_image_path)})
    response = client.responses.create(
        model=model,
        store=False,
        input=[
            {
                "role": "system",
                "content": (
                    "Voce monta matrizes finais para o Editar dados do PowerPoint. "
                    "A saida deve ser exata, tipada e rastreavel. "
                    "Voce nao interpreta valores por imagem de planilha; usa apenas dumps raw dos XLSX. "
                    "Quando um label do XLSX for mais completo que o label atual do PPT, use o label do XLSX na matriz final."
                ),
            },
            {"role": "user", "content": content},
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": "slide_target_matrices",
                "schema": _schema(),
                "strict": True,
            }
        },
    )
    text = getattr(response, "output_text", "") or _response_text_fallback(response)
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


def _image_data_url(path: Path) -> str:
    return "data:image/png;base64," + base64.b64encode(path.read_bytes()).decode("ascii")


def _response_text_fallback(response: Any) -> str:
    try:
        return response.output[0].content[0].text
    except Exception:
        return str(response)
