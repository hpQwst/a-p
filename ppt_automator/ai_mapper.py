from __future__ import annotations

from dataclasses import dataclass
import json
import os
from typing import Any

from .ai import build_openai_client
from .ppt_discovery import PptTarget
from .table_normalizer import TransformPlan, source_match_candidates
from .xlsx_parser import ParsedXlsxTable


@dataclass(frozen=True)
class AiSourceMatchSuggestion:
    target: str
    datasource: str
    confidence: float
    reason: str


def suggest_source_matches_with_ai(
    targets: list[PptTarget],
    sources: list[ParsedXlsxTable],
    existing_plan_ids: set[str] | None = None,
    root: Any = None,
    max_targets: int | None = None,
    candidates_per_target: int = 4,
) -> list[AiSourceMatchSuggestion]:
    plan_ids = existing_plan_ids or set()
    eligible_targets = [
        target
        for target in targets
        if target.object_type in {"chart", "table"} and target.shape_name not in plan_ids
    ]
    if not eligible_targets or not sources:
        return []

    target_limit = max_targets if max_targets is not None else _env_int("AUTO_PPT_AI_MATCH_TARGET_LIMIT", 40)
    min_confidence = _env_float("AUTO_PPT_AI_MATCH_MIN_CONFIDENCE", 0.55)
    min_local_score = _env_float("AUTO_PPT_AI_MATCH_MIN_LOCAL_SCORE", 0.25)
    ranked_targets = [
        target
        for target in eligible_targets
        if _top_local_score(target, sources) >= min_local_score
    ]
    request_targets = ranked_targets[: max(target_limit, 1)]
    if not request_targets:
        return []
    valid_targets = {target.shape_name for target in request_targets}
    valid_sources = {source.file_name for source in sources}

    client, model = build_openai_client(root)
    payload = {
        "task": (
            "Escolha o XLSX mais provavel para cada target ainda sem datasource automatico. "
            "Use o contrato do Editar dados do PowerPoint e a estrutura detectada do XLSX."
        ),
        "cost_control": {
            "targets_sent": len(request_targets),
            "total_unmatched_targets": len(eligible_targets),
            "candidate_filtered_targets": len(ranked_targets),
            "candidates_per_target": candidates_per_target,
            "rule": "Cada target recebe somente os melhores candidatos locais para reduzir tokens.",
        },
        "targets": [
            _ai_match_target_payload(target, sources, candidates_per_target)
            for target in request_targets
        ],
        "rules": [
            "Escolha no maximo um datasource por target.",
            "datasource deve ser exatamente um file_name de candidates.",
            "Prefira compatibilidade entre categorias/series do Editar dados e do XLSX.",
            "Use o texto proximo do slide apenas como apoio semantico.",
            "Se nenhum candidato fizer sentido, omita o target da resposta.",
            "Nao invente valores nem altere dados; esta etapa so escolhe o arquivo.",
        ],
    }
    schema = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "suggestions": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "target": {"type": "string"},
                        "datasource": {"type": "string"},
                        "confidence": {"type": "number", "minimum": 0, "maximum": 1},
                        "reason": {"type": "string"},
                    },
                    "required": ["target", "datasource", "confidence", "reason"],
                },
            }
        },
        "required": ["suggestions"],
    }
    response = client.responses.create(
        model=model,
        store=False,
        input=[
            {
                "role": "system",
                "content": (
                    "Voce e um analista de automacao de PowerPoint. "
                    "Seu trabalho e mapear targets PPT para arquivos XLSX com cautela. "
                    "Quando estiver incerto, omita a sugestao."
                ),
            },
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": "ppt_target_datasource_match",
                "schema": schema,
                "strict": True,
            }
        },
    )
    text = getattr(response, "output_text", "") or _response_text_fallback(response)
    data = json.loads(text)
    output: list[AiSourceMatchSuggestion] = []
    for item in data.get("suggestions", []):
        target = str(item.get("target") or "")
        datasource = str(item.get("datasource") or "")
        confidence = float(item.get("confidence") or 0)
        if target not in valid_targets or datasource not in valid_sources or confidence < min_confidence:
            continue
        output.append(
            AiSourceMatchSuggestion(
                target=target,
                datasource=datasource,
                confidence=confidence,
                reason=str(item.get("reason") or ""),
            )
        )
    return output


def build_ai_mapping_payload(targets: list[PptTarget], sources: list[ParsedXlsxTable]) -> dict[str, Any]:
    return {
        "targets": [_target_payload(target) for target in targets],
        "datasources": [_source_payload(source) for source in sources],
        "expected_response": {
            "target": "shape_name",
            "object_type": "chart|table|text|shape",
            "datasource": "arquivo.xlsx",
            "orientation_xlsx": "series_rows_categories_columns",
            "orientation_ppt": "categories_rows_series_columns",
            "action": "transpose|align|fill_table_cells",
            "series": [],
            "categories": [],
            "preserve_percentage_decimal": True,
            "confidence": 0.95,
            "reason": "explicacao curta",
        },
    }


def plans_to_ai_payload(plans: list[TransformPlan]) -> list[dict[str, Any]]:
    return [
        {
            "target": plan.target_id,
            "object_type": plan.object_type,
            "datasource": plan.datasource.file_name,
            "orientation_xlsx": plan.orientation_xlsx,
            "orientation_ppt": plan.orientation_ppt,
            "action": plan.action,
            "series": plan.series,
            "categories": plan.categories,
            "preserve_percentage_decimal": plan.preserve_percentage_decimal,
            "number_format": plan.number_format,
            "confidence": plan.confidence,
            "reason": plan.reason,
        }
        for plan in plans
    ]


def _target_payload(target: PptTarget) -> dict[str, Any]:
    return {
        "slide_index": target.slide_index,
        "slide_number": target.slide_number,
        "shape_name": target.shape_name,
        "shape_id": target.shape_id,
        "object_type": target.object_type,
        "position": {
            "left_in": target.left_in,
            "top_in": target.top_in,
            "width_in": target.width_in,
            "height_in": target.height_in,
        },
        "nearby_text": target.nearby_text,
        "slide_text": target.slide_text,
        "chart_xml": target.chart_xml,
        "workbook_embedded": target.workbook_embedded,
        "expected_orientation": target.expected_orientation,
        "expected_categories": target.expected_categories,
        "expected_series": target.expected_series,
        "table_cells": target.table_cells,
    }


def _source_payload(source: ParsedXlsxTable) -> dict[str, Any]:
    return {
        "source_id": source.source_id,
        "file_name": source.file_name,
        "sheet_name": source.sheet_name,
        "orientation": source.orientation,
        "categories": source.categories,
        "series": source.series,
        "preview_rows": source.preview_rows[:8],
        "metadata": source.metadata,
    }


def _ai_match_target_payload(
    target: PptTarget,
    sources: list[ParsedXlsxTable],
    candidates_per_target: int,
) -> dict[str, Any]:
    candidates = source_match_candidates(target, sources, limit=max(candidates_per_target, 1))
    return {
        "target": _target_payload_compact(target),
        "candidates": [
            {
                "file_name": candidate.source.file_name,
                "local_score": round(candidate.score, 4),
                "local_reason": candidate.reason,
                "xlsx": _source_payload_compact(candidate.source),
            }
            for candidate in candidates
        ],
    }


def _top_local_score(target: PptTarget, sources: list[ParsedXlsxTable]) -> float:
    candidates = source_match_candidates(target, sources, limit=1)
    return candidates[0].score if candidates else 0.0


def _target_payload_compact(target: PptTarget) -> dict[str, Any]:
    return {
        "target": target.shape_name,
        "slide": target.slide_number,
        "object_type": target.object_type,
        "nearby_text": _short(target.nearby_text, 500),
        "ppt_contract": {
            "orientation": target.expected_orientation or ("table_cells" if target.object_type == "table" else ""),
            "categories": _take(target.expected_categories, 16),
            "series": _take(target.expected_series, 16),
            "values_preview": _take_rows(target.expected_values, 8, 10),
            "table_cells_preview": _take_rows(target.table_cells, 8, 10),
        },
    }


def _source_payload_compact(source: ParsedXlsxTable) -> dict[str, Any]:
    return {
        "source_id": source.source_id,
        "file_name": source.file_name,
        "sheet_name": source.sheet_name,
        "used_range": source.used_range,
        "orientation": source.orientation,
        "categories": _take(source.categories, 16),
        "series": _take(source.series, 16),
        "preview_rows": _take_rows(source.preview_rows, 10, 12),
        "metadata": source.metadata,
    }


def _take(values: list[Any], limit: int) -> list[Any]:
    return list(values[:limit])


def _take_rows(rows: list[list[Any]], row_limit: int, col_limit: int) -> list[list[Any]]:
    return [list(row[:col_limit]) for row in rows[:row_limit]]


def _short(value: Any, limit: int) -> str:
    text = "" if value is None else str(value)
    return text if len(text) <= limit else f"{text[:limit]}..."


def _response_text_fallback(response: Any) -> str:
    try:
        return response.output[0].content[0].text
    except Exception:
        return str(response)


def _env_int(name: str, default: int) -> int:
    value = os.getenv(name, "").strip()
    if not value:
        return default
    try:
        return int(value)
    except ValueError:
        return default


def _env_float(name: str, default: float) -> float:
    value = os.getenv(name, "").strip()
    if not value:
        return default
    try:
        return float(value)
    except ValueError:
        return default
