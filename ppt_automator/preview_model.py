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
    fmt = _display_format(plan)
    if plan.object_type == "chart" and plan.orientation_ppt == "categories_rows_series_columns":
        return ["", *plan.series], [[plan.categories[i], *[_display_value(value, fmt) for value in row]] for i, row in enumerate(plan.values)]
    if plan.object_type == "chart":
        return ["", *plan.categories], [[plan.series[i], *[_display_value(value, fmt) for value in row]] for i, row in enumerate(plan.values)]
    if plan.values:
        return plan.categories, [[_display_value(value, fmt) for value in row] for row in plan.values]
    return [], []


def _display_format(plan: TransformPlan) -> str:
    """Espelha como o PowerPoint vai exibir os valores gravados (verbatim):
    1 casa decimal, com '%' apenas se o template do grafico ja usar percentual."""
    template = plan.target.value_format if plan.object_type == "chart" else plan.number_format
    return "percent" if "%" in (template or "") else "decimal"


def _display_value(value: Any, fmt: str) -> Any:
    if value is None or value == "":
        return ""
    parsed = _to_number(value)
    if parsed is None:
        # Texto (rotulos de tabela, etc.) permanece verbatim.
        return value
    if fmt == "percent":
        # Formato percentual do template: o PPT multiplica por 100 na exibicao.
        return _format_pt_number(parsed * 100) + "%"
    return _format_pt_number(parsed)


def _format_pt_number(value: float) -> str:
    return f"{value:.1f}".replace(".", ",")


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
