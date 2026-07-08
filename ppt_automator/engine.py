from __future__ import annotations

from dataclasses import replace
from io import BytesIO
import os
from pathlib import Path
from typing import Any, BinaryIO
from zipfile import ZipFile

from .embedded_workbook_writer import resolve_default_sheet_name, update_embedded_workbook
from .openxml_zip import replace_zip_parts_preserving_structure
from .ppt_chart_writer import chart_replacements
from .ppt_discovery import PptTarget, discover_ppt_targets, read_bytes
from .ppt_target_renamer import rename_targets_in_slide_xml
from .ppt_table_writer import update_table_slide_xml
from .preview_model import PreviewTarget, build_preview
from .slide_datasources import collect_datasource_entries, entries_for_slide
from .table_normalizer import TransformPlan, build_transform_plans
from .typed_matrix import typed_matrix_to_rows
from .xlsx_parser import ParsedXlsxTable, parse_datasource_zip



InputFile = str | Path | bytes | bytearray | BinaryIO


def analyze_update_package(
    pptx_file: InputFile,
    datasources_zip: InputFile,
    formula_mode: str = "auto",
    slide_numbers: list[int] | set[int] | None = None,
) -> tuple[list[PptTarget], list[ParsedXlsxTable], list[TransformPlan]]:
    targets = discover_ppt_targets(pptx_file, numeric_only=False, include_text_shapes=False)
    selected_slides = {int(slide) for slide in (slide_numbers or []) if int(slide) > 0}
    if selected_slides:
        targets = [target for target in targets if target.slide_number in selected_slides]
    sources = parse_datasource_zip(
        datasources_zip,
        formula_mode=formula_mode,
        include_names=_datasource_names_for_slides(datasources_zip, selected_slides) if selected_slides else None,
    )
    plans = _build_slide_aware_plans(targets, sources, datasources_zip)
    return targets, sources, plans


def preview_update_package(
    pptx_file: InputFile,
    datasources_zip: InputFile,
    formula_mode: str = "auto",
) -> list[PreviewTarget]:
    _targets, _sources, plans = analyze_update_package(pptx_file, datasources_zip, formula_mode=formula_mode)
    return build_preview(plans)


def generate_updated_pptx(
    pptx_file: InputFile,
    plans: list[TransformPlan],
    targets: list[PptTarget] | None = None,
) -> bytes:
    ppt_bytes = read_bytes(pptx_file)
    replacements: dict[str, bytes] = {}
    table_plans_by_slide: dict[str, list[TransformPlan]] = {}
    rename_targets_by_slide: dict[str, list[PptTarget]] = {}
    if _write_internal_target_ids_enabled():
        rename_targets = targets or [plan.target for plan in plans]
        for target in rename_targets:
            rename_targets_by_slide.setdefault(target.slide_path, []).append(target)

    with ZipFile(BytesIO(ppt_bytes)) as zf:
        for plan in plans:
            if plan.object_type == "chart":
                chart_target = plan.target
                if not chart_target.sheet_name and chart_target.workbook_embedded and chart_target.workbook_embedded in zf.namelist():
                    # A descoberta nao extraiu a aba a partir das formulas do chart.xml
                    # (ex.: chart sem cache preenchido). Em vez de deixar o writer gravar
                    # uma aba adivinhada no cache do grafico, resolvemos aqui a aba real
                    # do workbook embutido que sera de fato escrita logo abaixo.
                    resolved_sheet_name = resolve_default_sheet_name(zf.read(chart_target.workbook_embedded))
                    if resolved_sheet_name:
                        chart_target = replace(chart_target, sheet_name=resolved_sheet_name)

                updates = chart_replacements(zf, chart_target, plan)

                if chart_target.workbook_embedded:
                    updates[chart_target.workbook_embedded] = update_embedded_workbook(
                        zf.read(chart_target.workbook_embedded),
                        chart_target.sheet_name,
                        _workbook_matrix(plan),
                    )

                replacements.update(updates)

            elif plan.object_type == "table":
                table_plans_by_slide.setdefault(plan.target.slide_path, []).append(plan)

        for slide_path, slide_plans in table_plans_by_slide.items():
            slide_xml = replacements.get(slide_path, zf.read(slide_path))
            for plan in slide_plans:
                slide_xml = update_table_slide_xml(slide_xml, plan.target, plan)
            replacements[slide_path] = slide_xml

        for slide_path, slide_targets in rename_targets_by_slide.items():
            if slide_path not in zf.namelist():
                continue
            slide_xml = replacements.get(slide_path, zf.read(slide_path))
            replacements[slide_path] = rename_targets_in_slide_xml(slide_xml, slide_targets)

    return replace_zip_parts_preserving_structure(ppt_bytes, replacements)


def _datasource_names_for_slides(datasources_zip: InputFile, slide_numbers: set[int]) -> set[str] | None:
    try:
        entries = collect_datasource_entries(datasources_zip)
    except Exception:
        return None
    if not entries or not any(entry.slide_number for entry in entries):
        return None
    selected = {
        entry.zip_path
        for entry in entries
        if entry.is_general or (entry.slide_number is not None and entry.slide_number in slide_numbers)
    }
    return selected or None


def _write_internal_target_ids_enabled() -> bool:
    value = os.getenv("AUTO_PPT_WRITE_INTERNAL_TARGET_IDS", "").strip().lower()
    return value in {"1", "true", "yes", "on"}


def _workbook_matrix(plan: TransformPlan) -> list[list[Any]]:
    if plan.typed_edit_data:
        return typed_matrix_to_rows(plan.typed_edit_data)
    if plan.orientation_ppt == "series_rows_categories_columns":
        matrix = [[" ", *plan.categories]]
        for index, series_name in enumerate(plan.series):
            values = plan.values[index] if index < len(plan.values) else []
            matrix.append([series_name, *values])
        return matrix

    matrix = [[" ", *plan.series]]
    for index, category in enumerate(plan.categories):
        values = plan.values[index] if index < len(plan.values) else []
        matrix.append([category, *values])
    return matrix


def _build_slide_aware_plans(
    targets: list[PptTarget],
    sources: list[ParsedXlsxTable],
    datasources_zip: InputFile,
) -> list[TransformPlan]:
    try:
        entries = collect_datasource_entries(datasources_zip)
    except Exception:
        return build_transform_plans(targets, sources)
    if not any(entry.slide_number for entry in entries):
        return build_transform_plans(targets, sources)
    entry_by_name = {entry.zip_path: entry for entry in entries}
    plans: list[TransformPlan] = []
    for target in targets:
        if target.object_type not in {"chart", "table"}:
            continue
        selected_entries, _warnings = entries_for_slide(entries, target.slide_number)
        selected_names = {entry.zip_path for entry in selected_entries}
        relevant_sources = [source for source in sources if source.file_name in selected_names]
        if not relevant_sources:
            relevant_sources = [
                source
                for source in sources
                if entry_by_name.get(source.file_name) is None or entry_by_name[source.file_name].is_general
            ]
        if not relevant_sources:
            relevant_sources = sources
        plans.extend(build_transform_plans([target], relevant_sources))
    return plans
