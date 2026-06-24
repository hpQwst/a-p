from __future__ import annotations

from dataclasses import dataclass, replace
from typing import Any
import re

from ppt_automator import analyze_update_package, generate_updated_pptx
from ppt_automator.ppt_discovery import PptTarget
from ppt_automator.preview_model import PreviewTarget, build_preview
from ppt_automator.table_normalizer import TransformPlan, normalize_to_target
from ppt_automator.xlsx_parser import ParsedXlsxTable, parse_xlsx_table


ManualSourceMap = dict[str, tuple[str, bytes]]


@dataclass(frozen=True)
class AnalysisResult:
    plans: list[TransformPlan]
    preview: list[PreviewTarget]
    targets: list[PptTarget]
    sources: list[ParsedXlsxTable]
    target_count: int
    source_count: int


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
    allowed_targets = {target.shape_name for target in chosen_targets}
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

    targets_by_id = {target.shape_name: target for target in targets}
    plans_by_id = {plan.target_id: plan for plan in plans}
    source_output = list(sources)

    for target_id, (filename, data) in manual_sources.items():
        target = targets_by_id.get(target_id)
        if target is None:
            continue
        source = parse_xlsx_table(
            data,
            file_name=f"upload_manual/{target_id}_{filename}",
            formula_mode="auto",
        )
        source_output.append(source)
        plans_by_id[target_id] = normalize_to_target(
            target,
            source,
            confidence=1.0,
            match_reason=f"Datasource enviado manualmente para o target {target_id}.",
        )

    ordered_ids = [plan.target_id for plan in plans]
    for target_id in plans_by_id:
        if target_id not in ordered_ids:
            ordered_ids.append(target_id)
    return [plans_by_id[target_id] for target_id in ordered_ids], source_output


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
