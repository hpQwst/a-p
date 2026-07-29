from __future__ import annotations

from dataclasses import dataclass
from typing import Any
import re

from .ppt_chart_writer import resolved_series_number_formats
from .table_normalizer import TransformPlan


@dataclass(frozen=True)
class PreviewTarget:
    slide: int
    target: str
    object_type: str
    datasource: str
    action: str
    reason: str
    confidence: float
    headers: list[str]
    rows: list[list[Any]]


def build_preview(plans: list[TransformPlan]) -> list[PreviewTarget]:
    previews: list[PreviewTarget] = []
    for plan in plans:
        headers, rows = _matrix_for_preview(plan)
        previews.append(
            PreviewTarget(
                slide=plan.target.slide_number,
                target=plan.target_id,
                object_type=plan.object_type,
                datasource=plan.datasource.file_name,
                action=plan.action,
                reason=plan.reason,
                confidence=round(plan.confidence * 100, 1),
                headers=headers,
                rows=rows,
            )
        )
    return previews


def _matrix_for_preview(plan: TransformPlan) -> tuple[list[str], list[list[Any]]]:
    if plan.object_type == "chart" and plan.orientation_ppt == "categories_rows_series_columns":
        formats = resolved_series_number_formats(plan.target, plan)
        return ["", *plan.series], [
            [
                plan.categories[i],
                *[
                    _display_chart_value(
                        value,
                        formats[column] if column < len(formats) else "0.0",
                    )
                    for column, value in enumerate(row)
                ],
            ]
            for i, row in enumerate(plan.values)
        ]
    if plan.object_type == "chart":
        formats = resolved_series_number_formats(plan.target, plan)
        return ["", *plan.categories], [
            [
                plan.series[i],
                *[
                    _display_chart_value(
                        value,
                        formats[i] if i < len(formats) else "0.0",
                    )
                    for value in row
                ],
            ]
            for i, row in enumerate(plan.values)
        ]
    if plan.object_type == "table" and plan.table_matrix:
        formatted = [
            [_display_table_value(value, plan.number_format) for value in row]
            for row in plan.table_matrix
        ]
        if plan.table_header_rows:
            return [str(value) for value in formatted[0]], formatted[plan.table_header_rows :]
        return [], formatted
    if plan.values:
        return plan.categories, [
            [_display_table_value(value, plan.number_format) for value in row]
            for row in plan.values
        ]
    return [], []


def _display_chart_value(value: Any, number_format: str) -> Any:
    if value is None or value == "":
        return ""
    parsed = _to_number(value)
    if parsed is None:
        return value
    is_percent = "%" in (number_format or "")
    decimals = _decimal_places(number_format)
    displayed = parsed * 100 if is_percent else parsed
    suffix = "%" if is_percent else ""
    return _format_pt_number(
        displayed,
        decimals=decimals,
        thousands="," in (number_format or ""),
    ) + suffix


def _display_table_value(value: Any, number_format: str) -> Any:
    if value is None:
        return ""
    if number_format == "thousands_pt_br":
        try:
            number = float(value)
            if number.is_integer():
                return _format_pt_number(number, decimals=0, thousands=True)
        except (TypeError, ValueError):
            pass
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value).replace(".", ",")
    return value


def _decimal_places(number_format: str) -> int:
    primary = str(number_format or "0.0").split(";", 1)[0]
    primary = re.sub(r'"[^"]*"', "", primary)
    match = re.search(r"\.([0#]+)", primary)
    return len(match.group(1)) if match else 0


def _format_pt_number(value: float, *, decimals: int, thousands: bool) -> str:
    text = f"{value:,.{decimals}f}" if thousands else f"{value:.{decimals}f}"
    return text.replace(",", "\0").replace(".", ",").replace("\0", ".")


def _to_number(value: Any) -> float | None:
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace("%", "").strip()
    text = re.sub(r"^,", "0,", text)
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None
