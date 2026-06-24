from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import BinaryIO
from zipfile import ZIP_DEFLATED, ZipFile

from .ppt_chart_writer import chart_replacements
from .ppt_discovery import PptTarget, discover_ppt_targets, read_bytes
from .ppt_table_writer import update_table_slide_xml
from .preview_model import PreviewTarget, build_preview
from .table_normalizer import TransformPlan, build_transform_plans
from .xlsx_parser import ParsedXlsxTable, parse_datasource_zip


InputFile = str | Path | bytes | bytearray | BinaryIO


def analyze_update_package(
    pptx_file: InputFile,
    datasources_zip: InputFile,
    formula_mode: str = "auto",
) -> tuple[list[PptTarget], list[ParsedXlsxTable], list[TransformPlan]]:
    targets = discover_ppt_targets(pptx_file)
    sources = parse_datasource_zip(datasources_zip, formula_mode=formula_mode)
    plans = build_transform_plans(targets, sources)
    return targets, sources, plans


def preview_update_package(
    pptx_file: InputFile,
    datasources_zip: InputFile,
    formula_mode: str = "auto",
) -> list[PreviewTarget]:
    _targets, _sources, plans = analyze_update_package(pptx_file, datasources_zip, formula_mode=formula_mode)
    return build_preview(plans)


def generate_updated_pptx(pptx_file: InputFile, plans: list[TransformPlan]) -> bytes:
    ppt_bytes = read_bytes(pptx_file)
    replacements: dict[str, bytes] = {}
    table_plans_by_slide: dict[str, list[TransformPlan]] = {}

    with ZipFile(BytesIO(ppt_bytes)) as zf:
        for plan in plans:
            if plan.object_type == "chart":
                replacements.update(chart_replacements(zf, plan.target, plan))
            elif plan.object_type == "table":
                table_plans_by_slide.setdefault(plan.target.slide_path, []).append(plan)

        for slide_path, slide_plans in table_plans_by_slide.items():
            slide_xml = replacements.get(slide_path, zf.read(slide_path))
            for plan in slide_plans:
                slide_xml = update_table_slide_xml(slide_xml, plan.target, plan)
            replacements[slide_path] = slide_xml

        output = BytesIO()
        with ZipFile(output, "w", ZIP_DEFLATED) as zout:
            for item in zf.infolist():
                data = replacements.get(item.filename, zf.read(item.filename))
                zout.writestr(item, data)
    return output.getvalue()
