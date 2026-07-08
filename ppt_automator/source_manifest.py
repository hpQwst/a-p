from __future__ import annotations

from typing import Any
import hashlib
import json

from .xlsx_parser import ParsedXlsxTable


def xlsx_source_manifest(source: ParsedXlsxTable) -> dict[str, Any]:
    metadata = dict(source.metadata or {})
    return {
        "source_id": source.source_id,
        "file_name": source.file_name,
        "sheet_name": source.sheet_name,
        "used_range": source.used_range,
        "orientation": source.orientation,
        "semantic_context": {
            "table_title": metadata.get("table_title", ""),
            "row_group_label": metadata.get("row_group_label", ""),
            "context_text": metadata.get("context_text", ""),
            "context_tokens": metadata.get("context_tokens", ""),
            "graph_id": metadata.get("graph_id", ""),
            "ppt_tag": metadata.get("ppt_tag", ""),
            "variable": metadata.get("variable", ""),
        },
        "structural_profile": _structural_profile(source),
        "categories": _take(source.categories, 18),
        "series": _take(source.series, 18),
        "preview_rows": _take_rows(source.preview_rows, 8, 10),
    }


def _structural_profile(source: ParsedXlsxTable) -> dict[str, Any]:
    row_count = len(source.preview_rows)
    col_count = max((len(row) for row in source.preview_rows), default=0)
    metadata = source.metadata or {}
    profile = {
        "name": _structure_name(source),
        "orientation": source.orientation,
        "row_count_preview": row_count,
        "col_count_preview": col_count,
        "category_count": len(source.categories),
        "series_count": len(source.series),
        "has_table_title": bool(metadata.get("table_title")),
        "has_row_group_label": bool(metadata.get("row_group_label")),
    }
    profile["signature"] = _signature(profile)
    profile["recipe_hint"] = _recipe_hint(source)
    return profile


def _structure_name(source: ParsedXlsxTable) -> str:
    metadata = source.metadata or {}
    prefix = "contextual" if metadata.get("table_title") or metadata.get("row_group_label") else "plain"
    if source.orientation == "key_value_rows":
        return f"{prefix}_key_value_rows"
    if source.orientation == "single_series_row_categories_columns":
        return f"{prefix}_single_series_by_columns"
    if source.orientation == "series_rows_categories_columns":
        return f"{prefix}_series_rows_categories_columns"
    if source.orientation == "categories_rows_series_columns":
        return f"{prefix}_categories_rows_series_columns"
    return f"{prefix}_{source.orientation or 'unknown'}"


def _recipe_hint(source: ParsedXlsxTable) -> dict[str, Any]:
    return {
        "operation": "keep",
        "drop_top_rows": 0,
        "drop_left_cols": 0,
        "header_row": 1 if source.preview_rows else 0,
        "label_column": 1 if source.categories else 0,
        "orientation_after": source.orientation or "unknown",
    }


def _signature(profile: dict[str, Any]) -> str:
    raw = json.dumps(profile, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha1(raw).hexdigest()[:12]


def _take(values: list[Any], limit: int) -> list[Any]:
    return list(values[:limit])


def _take_rows(rows: list[list[Any]], row_limit: int, col_limit: int) -> list[list[Any]]:
    return [list(row[:col_limit]) for row in rows[:row_limit]]
