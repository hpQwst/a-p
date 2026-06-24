from __future__ import annotations

from typing import Any

from .ppt_discovery import PptTarget
from .table_normalizer import TransformPlan
from .xlsx_parser import ParsedXlsxTable


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
