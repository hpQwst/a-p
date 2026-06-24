from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any
import json

from .ai import build_openai_client
from .table_normalizer import TransformPlan


@dataclass(frozen=True)
class AiTransformDiagnostic:
    target: str
    status: str
    confidence: float
    action: str
    reason: str
    row_mapping: dict[str, str]
    column_mapping: dict[str, str]
    recommended_edit_data: dict[str, Any]


def suggest_transform_diagnostics(
    plans: list[TransformPlan],
    root: Path | str | None = None,
) -> list[AiTransformDiagnostic]:
    client, model = build_openai_client(root)

    payload = {
        "task": (
            "Revise se a matriz proposta para cada target do PowerPoint respeita o contrato do "
            "Editar dados e usa corretamente o datasource XLSX. Quando houver problema, devolva tambem "
            "a tabela corrigida exatamente no formato que o Editar dados do PowerPoint deveria receber."
        ),
        "plans": [_plan_payload(plan) for plan in plans],
        "rules": [
            "O contrato do PPT vem do workbook embutido do grafico/tabela e deve ser preservado.",
            "Confira linhas e colunas do PPT contra linhas e colunas detectadas no XLSX.",
            "Confirme se e preciso alinhar ou transpor.",
            "Nao invente valores. A tabela corrigida deve usar apenas valores presentes no datasource XLSX ou na matriz proposta quando ela estiver correta.",
            "Sinalize review quando nomes nao batem semanticamente ou valores parecem deslocados.",
            "recommended_edit_data.headers deve representar a primeira linha do Editar dados, incluindo a celula vazia inicial quando existir.",
            "recommended_edit_data.rows deve representar todas as linhas da tabela do Editar dados, com o rotulo da linha na primeira coluna quando existir.",
            "Para percentuais, prefira strings legiveis como 16,2%, 0,8% ou 0,0% na tabela recomendada.",
        ],
    }
    schema = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "diagnostics": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "target": {"type": "string"},
                        "status": {"type": "string", "enum": ["ok", "review"]},
                        "confidence": {"type": "number", "minimum": 0, "maximum": 1},
                        "action": {"type": "string"},
                        "reason": {"type": "string"},
                        "row_mapping": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "additionalProperties": False,
                                "properties": {
                                    "ppt": {"type": "string"},
                                    "xlsx": {"type": "string"},
                                },
                                "required": ["ppt", "xlsx"],
                            },
                        },
                        "column_mapping": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "additionalProperties": False,
                                "properties": {
                                    "ppt": {"type": "string"},
                                    "xlsx": {"type": "string"},
                                },
                                "required": ["ppt", "xlsx"],
                            },
                        },
                        "recommended_edit_data": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "orientation": {"type": "string"},
                                "headers": {
                                    "type": "array",
                                    "items": {"type": "string"},
                                },
                                "rows": {
                                    "type": "array",
                                    "items": {
                                        "type": "array",
                                        "items": {"type": "string"},
                                    },
                                },
                            },
                            "required": ["orientation", "headers", "rows"],
                        },
                    },
                    "required": [
                        "target",
                        "status",
                        "confidence",
                        "action",
                        "reason",
                        "row_mapping",
                        "column_mapping",
                        "recommended_edit_data",
                    ],
                },
            }
        },
        "required": ["diagnostics"],
    }
    response = client.responses.create(
        model=model,
        store=False,
        input=[
            {
                "role": "system",
                "content": (
                    "Voce e um auditor de automacao de PowerPoint. "
                    "Seu papel e validar transformacoes de dados para graficos e tabelas, com muito cuidado. "
                    "Quando algo estiver errado, voce deve ser acionavel: alem de explicar o erro, devolva a tabela correta "
                    "no mesmo contrato de linhas e colunas do Editar dados do PowerPoint."
                ),
            },
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": "ppt_transform_diagnostics",
                "schema": schema,
                "strict": True,
            }
        },
    )
    text = getattr(response, "output_text", "") or _response_text_fallback(response)
    data = json.loads(text)
    requested_targets = {plan.target_id for plan in plans}
    diagnostics = []
    for item in data.get("diagnostics", []):
        target = str(item.get("target") or "")
        if target not in requested_targets:
            continue
        diagnostics.append(
            AiTransformDiagnostic(
                target=target,
                status=str(item.get("status") or "review"),
                confidence=float(item.get("confidence") or 0),
                action=str(item.get("action") or ""),
                reason=str(item.get("reason") or ""),
                row_mapping=_mapping_to_dict(item.get("row_mapping")),
                column_mapping=_mapping_to_dict(item.get("column_mapping")),
                recommended_edit_data=_recommended_edit_data(item.get("recommended_edit_data")),
            )
        )
    return diagnostics


def _plan_payload(plan: TransformPlan) -> dict[str, Any]:
    ppt_rows, ppt_columns = _plan_axes(plan)
    headers, rows = _edit_data_matrix(plan, ppt_rows, ppt_columns)
    return {
        "target": plan.target_id,
        "object_type": plan.object_type,
        "slide": plan.target.slide_number,
        "shape_name": plan.target.shape_name,
        "nearby_text": plan.target.nearby_text,
        "datasource": plan.datasource.file_name,
        "ppt_edit_data_contract": {
            "orientation": plan.orientation_ppt,
            "rows": ppt_rows,
            "columns": ppt_columns,
            "current_values": plan.target.expected_values[:8],
        },
        "xlsx_detected": {
            "orientation": plan.datasource.orientation,
            "categories": plan.datasource.categories,
            "series": plan.datasource.series,
            "preview_rows": plan.datasource.preview_rows[:20],
        },
        "proposed_transform": {
            "action": plan.action,
            "rows": ppt_rows,
            "columns": ppt_columns,
            "values": plan.values[:20],
            "edit_data_headers": headers,
            "edit_data_rows": rows[:20],
            "reason": plan.reason,
            "warnings": plan.warnings,
        },
    }


def _plan_axes(plan: TransformPlan) -> tuple[list[str], list[str]]:
    if plan.object_type == "chart" and plan.orientation_ppt == "series_rows_categories_columns":
        return plan.series, plan.categories
    if plan.object_type == "chart":
        return plan.categories, plan.series
    return plan.series, plan.categories


def _edit_data_matrix(plan: TransformPlan, rows_axis: list[str], columns_axis: list[str]) -> tuple[list[str], list[list[Any]]]:
    if plan.object_type == "chart":
        return ["", *columns_axis], [[rows_axis[i] if i < len(rows_axis) else "", *row] for i, row in enumerate(plan.values)]
    return list(columns_axis), [list(row) for row in plan.values]


def _response_text_fallback(response: Any) -> str:
    try:
        return response.output[0].content[0].text
    except Exception:
        return str(response)


def _mapping_to_dict(value: Any) -> dict[str, str]:
    if isinstance(value, dict):
        return {str(k): str(v) for k, v in value.items()}
    if not isinstance(value, list):
        return {}
    output: dict[str, str] = {}
    for item in value:
        if not isinstance(item, dict):
            continue
        left = str(item.get("ppt") or "").strip()
        right = str(item.get("xlsx") or "").strip()
        if left or right:
            output[left] = right
    return output


def _recommended_edit_data(value: Any) -> dict[str, Any]:
    if not isinstance(value, dict):
        return {"orientation": "", "headers": [], "rows": []}
    return {
        "orientation": str(value.get("orientation") or ""),
        "headers": [str(item) for item in value.get("headers") or []],
        "rows": [[str(cell) for cell in row] for row in value.get("rows") or [] if isinstance(row, list)],
    }
