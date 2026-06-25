from __future__ import annotations

from dataclasses import dataclass
from typing import Any
import re

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
    percentage = _plan_is_percentage(plan)
    if plan.object_type == "chart" and plan.orientation_ppt == "categories_rows_series_columns":
        return ["", *plan.series], [[plan.categories[i], *[_display_value(value, percentage) for value in row]] for i, row in enumerate(plan.values)]
    if plan.object_type == "chart":
        return ["", *plan.categories], [[plan.series[i], *[_display_value(value, percentage) for value in row]] for i, row in enumerate(plan.values)]
    if plan.values:
        return plan.categories, [[_display_value(value, percentage) for value in row] for row in plan.values]
    return [], []


def _plan_is_percentage(plan: TransformPlan) -> bool:
    if plan.preserve_percentage_decimal:
        return True
    values = [value for row in plan.values for value in row]
    if any(isinstance(value, str) and "%" in value for value in values):
        return True
    numeric = [_to_number(value) for value in values]
    numeric = [value for value in numeric if value is not None]
    if not numeric:
        return False
    decimal_like = sum(1 for value in numeric if 0 <= abs(value) <= 1 and value not in {0, 1})
    return decimal_like >= max(2, len(numeric) * 0.5)


def _display_value(value: Any, percentage: bool) -> Any:
    if value is None:
        return ""
    if not percentage:
        return value
    parsed = _to_number(value)
    if parsed is None:
        return value
    if isinstance(value, str) and "%" in value:
        parsed = parsed / 100 if abs(parsed) > 1 else parsed
    return _format_pt_percent(parsed)


def _format_pt_percent(value: float) -> str:
    percent = value * 100
    text = f"{percent:.1f}%".replace(".", ",")
    if text.startswith("0,"):
        return text[1:]
    if text.startswith("-0,"):
        return "-" + text[2:]
    return text


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
