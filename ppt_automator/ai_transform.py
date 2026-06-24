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


def suggest_transform_diagnostics(
    plans: list[TransformPlan],
    root: Path | str | None = None,
) -> list[AiTransformDiagnostic]:
    client, model = build_openai_client(root)

    payload = {
        "task": (
            "Revise se a matriz proposta para cada target do PowerPoint respeita o contrato do "
            "Editar dados e usa corretamente o datasource XLSX."
        ),
        "plans": [_plan_payload(plan) for plan in plans],
        "rules": [
            "O contrato do PPT vem do workbook embutido do grafico/tabela e deve ser preservado.",
            "Confira linhas e colunas do PPT contra linhas e colunas detectadas no XLSX.",
            "Confirme se e preciso alinhar ou transpor.",
            "Nao invente valores. A matriz proposta deve vir do datasource.",
            "Sinalize review quando nomes nao batem semanticamente ou valores parecem deslocados.",
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
                    },
                    "required": [
                        "target",
                        "status",
                        "confidence",
                        "action",
                        "reason",
                        "row_mapping",
                        "column_mapping",
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
                    "Seu papel e validar transformacoes de dados para graficos e tabelas, com muito cuidado."
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
            )
        )
    return diagnostics


def _plan_payload(plan: TransformPlan) -> dict[str, Any]:
    ppt_rows, ppt_columns = _plan_axes(plan)
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
            "preview_rows": plan.datasource.preview_rows[:8],
        },
        "proposed_transform": {
            "action": plan.action,
            "rows": ppt_rows,
            "columns": ppt_columns,
            "values": plan.values[:8],
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
