from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any
import re

from ppt_automator import analyze_update_package, generate_updated_pptx
from ppt_automator.ai_mapper import AiSourceMatchSuggestion
from ppt_automator.ppt_discovery import PptTarget
from ppt_automator.preview_model import PreviewTarget, build_preview
from ppt_automator.table_normalizer import TransformPlan, normalize_to_target
from ppt_automator.xlsx_parser import ParsedXlsxTable, parse_xlsx_table


ManualSourcePayload = tuple[str, bytes] | tuple[str, bytes, str]
ManualSourceMap = dict[str, ManualSourcePayload]
SavedSourceMatchMap = dict[str, dict[str, Any] | str]


@dataclass(frozen=True)
class AnalysisResult:
    plans: list[TransformPlan]
    preview: list[PreviewTarget]
    targets: list[PptTarget]
    sources: list[ParsedXlsxTable]
    target_count: int
    source_count: int
    warnings: list[str]


def analyze_files(
    pptx_bytes: bytes,
    datasource_zip_bytes: bytes,
    manual_sources: ManualSourceMap | None = None,
    slide_numbers: list[int] | set[int] | None = None,
) -> AnalysisResult:
    targets, sources, plans = analyze_update_package(pptx_bytes, datasource_zip_bytes)
    targets, plans = _select_slides(targets, plans, slide_numbers)
    plans, sources = _apply_manual_sources(targets, sources, plans, manual_sources or {})
    return AnalysisResult(
        plans=plans,
        preview=build_preview(plans),
        targets=targets,
        sources=sources,
        target_count=len(targets),
        source_count=len(sources),
        warnings=_analysis_warnings(targets, sources, plans),
    )


def generate_file(
    pptx_bytes: bytes,
    datasource_zip_bytes: bytes,
    manual_sources: ManualSourceMap | None = None,
    slide_numbers: list[int] | set[int] | None = None,
) -> bytes:
    targets, sources, plans = analyze_update_package(pptx_bytes, datasource_zip_bytes)
    targets, plans = _select_slides(targets, plans, slide_numbers)
    plans, _sources = _apply_manual_sources(targets, sources, plans, manual_sources or {})
    return generate_updated_pptx(pptx_bytes, plans)


def parse_slide_selection(value: str) -> list[int]:
    text = (value or "").strip()
    if not text:
        return []
    output: set[int] = set()
    for part in re.split(r"[,;\s]+", text):
        part = part.strip()
        if not part:
            continue
        if re.fullmatch(r"\d+[-:]\d+", part):
            left_text, right_text = re.split(r"[-:]", part, maxsplit=1)
            left = int(left_text)
            right = int(right_text)
            if left <= 0 or right <= 0:
                raise ValueError("Informe apenas números de slides maiores que zero.")
            start, end = sorted((left, right))
            output.update(range(start, end + 1))
            continue
        if re.fullmatch(r"\d+", part):
            slide = int(part)
            if slide <= 0:
                raise ValueError("Informe apenas números de slides maiores que zero.")
            output.add(slide)
            continue
        raise ValueError("Seleção de slides inválida. Use algo como: 3, 5, 8-10.")
    return sorted(output)


def apply_ai_recommendations_to_analysis(
    analysis: AnalysisResult,
    ai_diagnostics: dict[str, dict],
) -> AnalysisResult:
    plans = apply_ai_recommendations(analysis.plans, ai_diagnostics)
    return AnalysisResult(
        plans=plans,
        preview=build_preview(plans),
        targets=analysis.targets,
        sources=analysis.sources,
        target_count=analysis.target_count,
        source_count=analysis.source_count,
        warnings=analysis.warnings,
    )


def apply_ai_source_matches_to_analysis(
    analysis: AnalysisResult,
    ai_matches: dict[str, dict[str, Any]] | list[AiSourceMatchSuggestion],
) -> AnalysisResult:
    if not ai_matches:
        return analysis
    normalized_matches = _normalize_ai_matches(ai_matches)
    if not normalized_matches:
        return analysis

    targets_by_id = {target.target_id: target for target in analysis.targets}
    sources_by_file = {source.file_name: source for source in analysis.sources}
    plans_by_id = {plan.target_id: plan for plan in analysis.plans}

    for target_id, match in normalized_matches.items():
        if target_id in plans_by_id:
            continue
        target = targets_by_id.get(target_id)
        source = sources_by_file.get(str(match.get("datasource") or ""))
        if target is None or source is None:
            continue
        confidence = float(match.get("confidence") or 0)
        reason = str(match.get("reason") or "IA escolheu este datasource por compatibilidade.")
        plans_by_id[target_id] = normalize_to_target(
            target,
            source,
            confidence=confidence,
            match_reason=f"IA sugeriu {source.file_name}: {reason}",
        )

    ordered_ids = [plan.target_id for plan in analysis.plans]
    for target in analysis.targets:
        if target.target_id in plans_by_id and target.target_id not in ordered_ids:
            ordered_ids.append(target.target_id)
    plans = [plans_by_id[target_id] for target_id in ordered_ids]
    return AnalysisResult(
        plans=plans,
        preview=build_preview(plans),
        targets=analysis.targets,
        sources=analysis.sources,
        target_count=analysis.target_count,
        source_count=analysis.source_count,
        warnings=_analysis_warnings(analysis.targets, analysis.sources, plans),
    )


def apply_saved_source_matches_to_analysis(
    analysis: AnalysisResult,
    saved_matches: SavedSourceMatchMap,
) -> AnalysisResult:
    if not saved_matches:
        return analysis
    normalized_matches = _normalize_saved_matches(saved_matches)
    if not normalized_matches:
        return analysis

    targets_by_id = {target.target_id: target for target in analysis.targets}
    sources_by_name = _sources_by_match_name(analysis.sources)
    plans_by_id = {plan.target_id: plan for plan in analysis.plans}

    for target_id, match in normalized_matches.items():
        target = targets_by_id.get(target_id)
        source = sources_by_name.get(_source_match_key(match.get("datasource")))
        if target is None or source is None:
            continue
        confidence = float(match.get("confidence") or 1.0)
        reason = str(match.get("reason") or "Mapeamento salvo aplicado para este target.")
        plans_by_id[target_id] = normalize_to_target(
            target,
            source,
            confidence=confidence,
            match_reason=reason,
        )

    ordered_ids = [plan.target_id for plan in analysis.plans]
    for target in analysis.targets:
        if target.target_id in plans_by_id and target.target_id not in ordered_ids:
            ordered_ids.append(target.target_id)
    plans = [plans_by_id[target_id] for target_id in ordered_ids]
    return AnalysisResult(
        plans=plans,
        preview=build_preview(plans),
        targets=analysis.targets,
        sources=analysis.sources,
        target_count=analysis.target_count,
        source_count=analysis.source_count,
        warnings=_analysis_warnings(analysis.targets, analysis.sources, plans),
    )


def apply_ai_recommendations(
    plans: list[TransformPlan],
    ai_diagnostics: dict[str, dict],
) -> list[TransformPlan]:
    if not ai_diagnostics:
        return plans
    output: list[TransformPlan] = []
    for plan in plans:
        diagnostic = ai_diagnostics.get(plan.target_id) or {}
        corrected = _plan_from_ai_edit_data(plan, diagnostic)
        output.append(corrected or plan)
    return output


def _select_slides(
    targets: list[PptTarget],
    plans: list[TransformPlan],
    slide_numbers: list[int] | set[int] | None,
) -> tuple[list[PptTarget], list[TransformPlan]]:
    selected = {int(slide) for slide in (slide_numbers or []) if int(slide) > 0}
    if not selected:
        return targets, plans
    chosen_targets = [target for target in targets if target.slide_number in selected]
    allowed_targets = {target.target_id for target in chosen_targets}
    chosen_plans = [plan for plan in plans if plan.target_id in allowed_targets]
    return chosen_targets, chosen_plans


def _apply_manual_sources(
    targets: list[PptTarget],
    sources: list[ParsedXlsxTable],
    plans: list[TransformPlan],
    manual_sources: ManualSourceMap,
) -> tuple[list[TransformPlan], list[ParsedXlsxTable]]:
    if not manual_sources:
        return plans, sources

    targets_by_id = {target.target_id: target for target in targets}
    plans_by_id = {plan.target_id: plan for plan in plans}
    source_output = list(sources)

    for target_id, payload in manual_sources.items():
        filename, data, cell_range = _manual_payload(payload)
        target = targets_by_id.get(target_id)
        if target is None:
            continue
        source = parse_xlsx_table(
            data,
            file_name=f"upload_manual/{target_id}_{filename}",
            formula_mode="auto",
            cell_range=cell_range,
        )
        source_output.append(source)
        range_text = f" usando o range {cell_range}" if cell_range else ""
        plans_by_id[target_id] = normalize_to_target(
            target,
            source,
            confidence=1.0,
            match_reason=f"Datasource enviado manualmente para o target {target_id}{range_text}.",
        )

    ordered_ids = [plan.target_id for plan in plans]
    for target_id in plans_by_id:
        if target_id not in ordered_ids:
            ordered_ids.append(target_id)
    return [plans_by_id[target_id] for target_id in ordered_ids], source_output


def _manual_payload(payload: ManualSourcePayload) -> tuple[str, bytes, str]:
    if len(payload) == 2:
        filename, data = payload
        return filename, data, ""
    filename, data, cell_range = payload
    return filename, data, str(cell_range or "").strip()


def _normalize_ai_matches(
    ai_matches: dict[str, dict[str, Any]] | list[AiSourceMatchSuggestion],
) -> dict[str, dict[str, Any]]:
    if isinstance(ai_matches, dict):
        return ai_matches
    return {
        item.target: {
            "datasource": item.datasource,
            "confidence": item.confidence,
            "reason": item.reason,
        }
        for item in ai_matches
    }


def _normalize_saved_matches(saved_matches: SavedSourceMatchMap) -> dict[str, dict[str, Any]]:
    output: dict[str, dict[str, Any]] = {}
    for target_id, match in saved_matches.items():
        clean_target_id = str(target_id or "").strip()
        if not clean_target_id:
            continue
        if isinstance(match, dict):
            datasource = str(match.get("datasource") or match.get("file_name") or "").strip()
            confidence = match.get("confidence", 1.0)
            reason = str(match.get("reason") or "Mapeamento salvo aplicado para este target.")
        else:
            datasource = str(match or "").strip()
            confidence = 1.0
            reason = "Mapeamento salvo aplicado para este target."
        if not datasource:
            continue
        output[clean_target_id] = {
            "datasource": datasource,
            "confidence": confidence,
            "reason": reason,
        }
    return output


def _sources_by_match_name(sources: list[ParsedXlsxTable]) -> dict[str, ParsedXlsxTable]:
    output: dict[str, ParsedXlsxTable] = {}
    for source in sources:
        keys = {
            _source_match_key(source.file_name),
            _source_match_key(Path(source.file_name).name),
        }
        for key in keys:
            if key and key not in output:
                output[key] = source
    return output


def _source_match_key(value: Any) -> str:
    text = str(value or "").strip().replace("\\", "/")
    if not text:
        return ""
    return Path(text).name.lower()


def _analysis_warnings(
    targets: list[PptTarget],
    sources: list[ParsedXlsxTable],
    plans: list[TransformPlan],
) -> list[str]:
    warnings: list[str] = []
    updatable_targets = [target for target in targets if target.object_type in {"chart", "table"}]
    target_counts = Counter(target.target_id for target in updatable_targets)
    duplicate_targets = sorted(target_id for target_id, count in target_counts.items() if count > 1)
    if duplicate_targets:
        warnings.append(
            "Existem targets chart/table com identificador repetido no PPT: "
            + ", ".join(duplicate_targets[:12])
            + ". Renomeie os shapes no PowerPoint ou use override por target para evitar ambiguidade."
        )

    source_counts = Counter(source.source_id for source in sources if source.source_id)
    duplicate_sources = sorted(source_id for source_id, count in source_counts.items() if count > 1)
    if duplicate_sources:
        warnings.append(
            "Existem XLSX com o mesmo identificador numerico: "
            + ", ".join(duplicate_sources[:12])
            + ". O sistema usa estrutura e contexto como desempate, mas vale revisar o preview."
        )

    mapped_ids = {plan.target_id for plan in plans}
    missing_count = sum(1 for target in updatable_targets if target.target_id not in mapped_ids)
    if updatable_targets and not plans:
        warnings.append(
            "Nenhum match automatico foi aceito. Ative a IA ou use Trocar XLSX deste target com range manual."
        )
    elif missing_count:
        warnings.append(
            f"{missing_count} target(s) chart/table ainda estao sem datasource. Revise os cards sem datasource automatico."
        )
    return warnings


def _plan_from_ai_edit_data(plan: TransformPlan, diagnostic: dict[str, Any]) -> TransformPlan | None:
    if plan.object_type != "chart":
        return None
    if str(diagnostic.get("status") or "").lower() != "review":
        return None
    edit_data = diagnostic.get("recommended_edit_data") or {}
    if not isinstance(edit_data, dict):
        return None
    headers = [str(item).strip() for item in edit_data.get("headers") or []]
    raw_rows = edit_data.get("rows") or []
    rows = [[str(cell).strip() for cell in row] for row in raw_rows if isinstance(row, list) and row]
    if not rows:
        return None

    value_cols = max((len(row) - 1 for row in rows), default=0)
    if plan.orientation_ppt == "series_rows_categories_columns":
        categories = _axis_headers(headers, value_cols, plan.categories)
        series = [row[0] for row in rows]
        values = [_parse_row_values(row[1:], percentage_hint=plan.preserve_percentage_decimal) for row in rows]
        if not _valid_matrix(values, expected_rows=len(series), expected_cols=len(categories)):
            return None
        return replace(
            plan,
            categories=categories,
            series=series,
            values=values,
            action=f"{plan.action}_ai_corrected",
            reason=f"{plan.reason} Matriz substituída pela tabela correta sugerida pela IA.",
            preserve_percentage_decimal=plan.preserve_percentage_decimal or _rows_have_percent(raw_rows),
        )

    series = _axis_headers(headers, value_cols, plan.series)
    categories = [row[0] for row in rows]
    values = [_parse_row_values(row[1:], percentage_hint=plan.preserve_percentage_decimal) for row in rows]
    if not _valid_matrix(values, expected_rows=len(categories), expected_cols=len(series)):
        return None
    return replace(
        plan,
        categories=categories,
        series=series,
        values=values,
        action=f"{plan.action}_ai_corrected",
        reason=f"{plan.reason} Matriz substituída pela tabela correta sugerida pela IA.",
        preserve_percentage_decimal=plan.preserve_percentage_decimal or _rows_have_percent(raw_rows),
    )


def _axis_headers(headers: list[str], expected_count: int, fallback: list[str]) -> list[str]:
    clean = [header for header in headers if header]
    if len(clean) == expected_count:
        return clean
    if len(headers) == expected_count + 1:
        return headers[1:]
    if len(fallback) == expected_count:
        return list(fallback)
    return clean or list(fallback)


def _parse_row_values(values: list[str], percentage_hint: bool) -> list[Any]:
    return [_parse_recommended_value(value, percentage_hint=percentage_hint) for value in values]


def _parse_recommended_value(value: Any, percentage_hint: bool = False) -> Any:
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    is_percent = "%" in text
    normalized = text.replace("%", "").strip()
    normalized = re.sub(r"^,", "0,", normalized)
    normalized = re.sub(r"^-,", "-0,", normalized)
    normalized = normalized.replace(" ", "")
    if "," in normalized and "." in normalized:
        normalized = normalized.replace(".", "").replace(",", ".")
    else:
        normalized = normalized.replace(",", ".")
    try:
        number = float(normalized)
    except ValueError:
        return value
    if is_percent:
        return number / 100
    if percentage_hint and abs(number) > 1:
        return number / 100
    return number


def _valid_matrix(values: list[list[Any]], expected_rows: int, expected_cols: int) -> bool:
    if expected_rows <= 0 or expected_cols <= 0:
        return False
    if len(values) != expected_rows:
        return False
    return all(len(row) == expected_cols for row in values)


def _rows_have_percent(rows: Any) -> bool:
    if not isinstance(rows, list):
        return False
    return any("%" in str(cell) for row in rows if isinstance(row, list) for cell in row)
