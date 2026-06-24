from __future__ import annotations

from dataclasses import dataclass

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
) -> AnalysisResult:
    targets, sources, plans = analyze_update_package(pptx_bytes, datasource_zip_bytes)
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
) -> bytes:
    _targets, _sources, plans = analyze_update_package(pptx_bytes, datasource_zip_bytes)
    plans, _sources = _apply_manual_sources(_targets, _sources, plans, manual_sources or {})
    return generate_updated_pptx(pptx_bytes, plans)


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
