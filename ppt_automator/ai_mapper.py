from __future__ import annotations

from dataclasses import dataclass, field
import json
import os
import time
from typing import Any

from .ai import build_openai_client, reasoning_for_operation
from .ai_debug import log_ai_request, log_ai_response, log_debug_event
from .ppt_discovery import PptTarget
from .source_manifest import xlsx_source_manifest
from .table_normalizer import TransformPlan, source_match_candidates
from .xlsx_parser import ParsedXlsxTable


@dataclass(frozen=True)
class AiSourceMatchSuggestion:
    target: str
    datasource: str
    confidence: float
    reason: str
    recipe_suggestion: dict[str, Any] = field(default_factory=dict)


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
        if target.object_type in {"chart", "table"} and target.target_id not in plan_ids
    ]
    if not eligible_targets or not sources:
        return []

    target_limit = max_targets if max_targets is not None else _env_int("AUTO_PPT_AI_MATCH_TARGET_LIMIT", 40)
    min_confidence = _env_float("AUTO_PPT_AI_MATCH_MIN_CONFIDENCE", 0.55)
    min_local_score = _env_float("AUTO_PPT_AI_MATCH_MIN_LOCAL_SCORE", 0.25)
    scored_targets = [
        (_top_local_score(target, sources), index, target)
        for index, target in enumerate(eligible_targets)
    ]
    scored_targets.sort(key=lambda item: (-item[0], item[1]))
    ranked_targets = [target for score, _index, target in scored_targets if score >= min_local_score]
    fallback_targets = [target for score, _index, target in scored_targets if score < min_local_score]
    request_targets = (ranked_targets + fallback_targets)[: max(target_limit, 1)]
    if not request_targets:
        return []
    valid_targets = {target.target_id for target in request_targets}
    valid_sources = {source.file_name for source in sources}

    client, model = build_openai_client(root, operation="source_match")
    slide_records = _slide_match_records(request_targets, sources, candidates_per_target)
    payload = {
        "task": (
            "Para cada slide, escolha o datasource XLSX mais provavel para cada objeto PPT pendente ou de baixa confianca "
            "e sugira apenas a receita estrutural necessaria para encaixar o datasource no Editar dados."
        ),
        "format": "jsonl_por_slide",
        "slide_records_jsonl": "\n".join(
            json.dumps(record, ensure_ascii=False, separators=(",", ":")) for record in slide_records
        ),
        "cost_control": {
            "targets_sent": len(request_targets),
            "total_unmatched_targets": len(eligible_targets),
            "total_review_targets": len(eligible_targets),
            "candidate_filtered_targets": len(ranked_targets),
            "fallback_targets_sent": sum(1 for target in request_targets if target in fallback_targets),
            "candidates_per_target": candidates_per_target,
            "rule": (
                "Cada target recebe os melhores candidatos locais. "
                "Quando nenhum target passa no score minimo, ainda enviamos os pendentes para a IA revisar."
            ),
        },
        "rules": [
            "Escolha no maximo um datasource por target.",
            "datasource deve ser exatamente um file_name listado nos datasources do mesmo slide.",
            "Use colunas, linhas e titulos/contextos do Editar dados e do XLSX.",
            "Use o texto proximo do slide apenas como apoio semantico.",
            "Se nenhum candidato fizer sentido, omita o target da resposta.",
            "Nao invente valores e nao devolva matriz final.",
            "recipe_suggestion deve ser uma receita estrutural pequena e validavel, nao os dados finais.",
        ],
        "expected_response": {
            "matches": [
                {
                    "target": "target_id",
                    "datasource": "datasources/arquivo.xlsx",
                    "confidence": 0.9,
                    "reason": "explicacao curta",
                    "recipe_suggestion": {
                        "operation": "keep|transpose|drop_and_keep|drop_and_transpose|unknown",
                        "drop_top_rows": 0,
                        "drop_left_cols": 0,
                        "header_row": 1,
                        "label_column": 1,
                        "orientation_after": "series_rows_categories_columns|categories_rows_series_columns|single_series_row_categories_columns|key_value_rows|table_cells|unknown",
                        "confidence": 0.9,
                        "reason": "por que esta estrutura encaixa no Editar dados",
                    },
                }
            ]
        },
    }
    schema = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "matches": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "target": {"type": "string"},
                        "datasource": {"type": "string"},
                        "confidence": {"type": "number", "minimum": 0, "maximum": 1},
                        "reason": {"type": "string"},
                        "recipe_suggestion": _recipe_schema(),
                    },
                    "required": ["target", "datasource", "confidence", "reason", "recipe_suggestion"],
                },
            }
        },
        "required": ["matches"],
    }
    request_kwargs = {
        "model": model,
        "store": False,
        "reasoning": reasoning_for_operation("source_match"),
        "input": [
            {
                "role": "system",
                "content": (
                    "Voce e um analista de automacao de PowerPoint. "
                    "Seu trabalho e mapear objetos PPT para arquivos XLSX e sugerir receitas estruturais simples. "
                    "Voce nunca monta a matriz final e nunca inventa valores."
                ),
            },
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
        ],
        "text": {
            "format": {
                "type": "json_schema",
                "name": "ppt_target_datasource_match",
                "schema": schema,
                "strict": True,
            }
        },
    }
    log_ai_request("source_match", request_kwargs)
    started = time.perf_counter()
    try:
        response = client.responses.create(**request_kwargs)
    except Exception as exc:
        log_debug_event(
            "ai_error",
            {
                "operation": "source_match",
                "elapsed_ms": round((time.perf_counter() - started) * 1000),
                "error": repr(exc),
            },
        )
        raise
    text = getattr(response, "output_text", "") or _response_text_fallback(response)
    log_ai_response("source_match", text, round((time.perf_counter() - started) * 1000), response)
    data = json.loads(text)
    output: list[AiSourceMatchSuggestion] = []
    returned_items = data.get("matches", data.get("suggestions", []))
    for item in returned_items:
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
                recipe_suggestion=_sanitize_recipe(item.get("recipe_suggestion") or {}),
            )
        )
    return output


def build_ai_mapping_payload(targets: list[PptTarget], sources: list[ParsedXlsxTable]) -> dict[str, Any]:
    return {
        "targets": [_target_payload(target) for target in targets],
        "datasources": [_source_payload(source) for source in sources],
        "expected_response": {
            "target": "target_id",
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


def _recipe_schema() -> dict[str, Any]:
    return {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "operation": {
                "type": "string",
                "enum": ["keep", "transpose", "drop_and_keep", "drop_and_transpose", "unknown"],
            },
            "drop_top_rows": {"type": "integer", "minimum": 0},
            "drop_left_cols": {"type": "integer", "minimum": 0},
            "header_row": {"type": "integer", "minimum": 0},
            "label_column": {"type": "integer", "minimum": 0},
            "orientation_after": {
                "type": "string",
                "enum": [
                    "series_rows_categories_columns",
                    "categories_rows_series_columns",
                    "single_series_row_categories_columns",
                    "key_value_rows",
                    "table_cells",
                    "unknown",
                ],
            },
            "confidence": {"type": "number", "minimum": 0, "maximum": 1},
            "reason": {"type": "string"},
        },
        "required": [
            "operation",
            "drop_top_rows",
            "drop_left_cols",
            "header_row",
            "label_column",
            "orientation_after",
            "confidence",
            "reason",
        ],
    }


def _slide_match_records(
    targets: list[PptTarget],
    sources: list[ParsedXlsxTable],
    candidates_per_target: int,
) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for slide in sorted({target.slide_number for target in targets}):
        slide_targets = [target for target in targets if target.slide_number == slide]
        candidates_by_target = {
            target.target_id: source_match_candidates(target, sources, limit=max(candidates_per_target, 1))
            for target in slide_targets
        }
        source_names = {
            candidate.source.file_name
            for candidates in candidates_by_target.values()
            for candidate in candidates
        }
        record_sources = [source for source in sources if source.file_name in source_names]
        records.append(
            {
                "slide": slide,
                "objects": [
                    _slide_object_payload(target, candidates_by_target.get(target.target_id, []))
                    for target in slide_targets
                ],
                "datasources": [_slide_source_payload(source) for source in record_sources],
            }
        )
    return records


def _slide_object_payload(target: PptTarget, candidates: list[SourceMatchCandidate]) -> dict[str, Any]:
    return {
        "target_id": target.target_id,
        "shape_name": target.shape_name,
        "object_type": target.object_type,
        "titles": _target_titles(target),
        "edit_data": _target_edit_data_profile(target),
        "candidate_datasources": [
            {
                "file_name": candidate.source.file_name,
                "local_score": round(candidate.score, 4),
                "local_reason": candidate.reason,
            }
            for candidate in candidates
        ],
    }


def _slide_source_payload(source: ParsedXlsxTable) -> dict[str, Any]:
    manifest = xlsx_source_manifest(source)
    return {
        "file_name": source.file_name,
        "sheet_name": source.sheet_name,
        "titles": manifest.get("semantic_context", {}),
        "structure": manifest.get("structural_profile", {}),
        "detected_data": _source_detected_data_profile(source),
        "preview_rows": _take_rows(source.preview_rows, 8, 10),
    }


def _target_titles(target: PptTarget) -> dict[str, Any]:
    # 'title' e o titulo estruturado do proprio objeto (caixa imediatamente acima),
    # o sinal mais forte de contexto. nearby_text/slide_text ficam como apoio.
    return {
        "title": _short(getattr(target, "title", "") or "", 160),
        "nearby_text": _short(target.nearby_text, 300),
        "slide_text": _short(target.slide_text, 400),
    }


def _target_edit_data_profile(target: PptTarget) -> dict[str, Any]:
    if target.object_type == "table":
        return {
            "orientation": "table_cells",
            "columns": _take(target.table_cells[0] if target.table_cells else [], 18),
            "rows": _take_rows(target.table_cells[1:] if len(target.table_cells) > 1 else target.table_cells, 12, 12),
            "shape": [len(target.table_cells), max((len(row) for row in target.table_cells), default=0)],
        }
    if target.expected_orientation == "series_rows_categories_columns":
        columns = target.expected_categories
        rows = target.expected_series
    else:
        columns = target.expected_series
        rows = target.expected_categories
    return {
        "orientation": target.expected_orientation or "unknown",
        "columns": _take(columns, 18),
        "rows": _take(rows, 18),
        "shape": [len(rows), len(columns)],
    }


def _source_detected_data_profile(source: ParsedXlsxTable) -> dict[str, Any]:
    if source.orientation in {"series_rows_categories_columns", "single_series_row_categories_columns"}:
        columns = source.categories
        rows = source.series
    elif source.orientation == "key_value_rows":
        columns = source.series or ["Valor"]
        rows = source.categories
    else:
        columns = source.series
        rows = source.categories
    return {
        "orientation": source.orientation or "unknown",
        "columns": _take(columns, 18),
        "rows": _take(rows, 18),
        "shape": [len(rows), len(columns)],
    }


def _sanitize_recipe(value: dict[str, Any]) -> dict[str, Any]:
    operations = {"keep", "transpose", "drop_and_keep", "drop_and_transpose", "unknown"}
    orientations = {
        "series_rows_categories_columns",
        "categories_rows_series_columns",
        "single_series_row_categories_columns",
        "key_value_rows",
        "table_cells",
        "unknown",
    }
    operation = str(value.get("operation") or "unknown")
    orientation_after = str(value.get("orientation_after") or "unknown")
    return {
        "operation": operation if operation in operations else "unknown",
        "drop_top_rows": _non_negative_int(value.get("drop_top_rows")),
        "drop_left_cols": _non_negative_int(value.get("drop_left_cols")),
        "header_row": _non_negative_int(value.get("header_row")),
        "label_column": _non_negative_int(value.get("label_column")),
        "orientation_after": orientation_after if orientation_after in orientations else "unknown",
        "confidence": max(0.0, min(float(value.get("confidence") or 0), 1.0)),
        "reason": str(value.get("reason") or ""),
    }


def _non_negative_int(value: Any) -> int:
    try:
        return max(int(value), 0)
    except (TypeError, ValueError):
        return 0


def _target_payload(target: PptTarget) -> dict[str, Any]:
    return {
        "slide_index": target.slide_index,
        "slide_number": target.slide_number,
        "target_id": target.target_id,
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
    return xlsx_source_manifest(source)


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
        "target": target.target_id,
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
    payload = xlsx_source_manifest(source)
    payload["categories"] = _take(source.categories, 16)
    payload["series"] = _take(source.series, 16)
    payload["preview_rows"] = _take_rows(source.preview_rows, 10, 12)
    return payload


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
