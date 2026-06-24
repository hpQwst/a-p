from __future__ import annotations

from dataclasses import dataclass
from typing import Any

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
        return ["", *plan.series], [[plan.categories[i], *row] for i, row in enumerate(plan.values)]
    if plan.object_type == "chart":
        return ["", *plan.categories], [[plan.series[i], *row] for i, row in enumerate(plan.values)]
    if plan.values:
        return plan.categories, plan.values
    return [], []
