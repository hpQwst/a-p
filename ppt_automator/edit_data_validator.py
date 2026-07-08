from __future__ import annotations

from typing import Any

from .typed_matrix import normalize_typed_edit_data, numeric_value


class EditDataValidationError(ValueError):
    pass


def validate_typed_edit_data(edit_data: dict[str, Any], object_type: str = "chart", target: Any = None) -> list[str]:
    errors: list[str] = []
    normalized = normalize_typed_edit_data(edit_data)
    headers = normalized.get("headers") or []
    rows = normalized.get("rows") or []
    if not rows:
        errors.append("final_edit_data precisa ter ao menos uma linha.")
    expected_cols = len(headers) if headers else (len(rows[0]) if rows else 0)
    if expected_cols <= 0:
        errors.append("final_edit_data precisa ter ao menos uma coluna.")
    for row_index, row in enumerate(rows, start=1):
        if len(row) != expected_cols:
            errors.append(f"Linha {row_index} tem {len(row)} coluna(s), esperado {expected_cols}.")
        for col_index, cell in enumerate(row, start=1):
            if cell.get("type") == "number" and numeric_value(cell) is None:
                errors.append(f"Celula numerica invalida em row={row_index}, col={col_index}.")
    for col_index, cell in enumerate(headers, start=1):
        if cell.get("type") == "number" and numeric_value(cell) is None:
            errors.append(f"Header numerico invalido em col={col_index}.")
    if object_type == "chart" and expected_cols < 2:
        errors.append("Grafico precisa ter pelo menos uma coluna de label e uma coluna de valor.")
    if object_type == "chart" and target is not None and rows:
        errors.extend(_chart_series_capacity_errors(target, headers, rows))
    return errors


def _chart_series_capacity_errors(target: Any, headers: list[Any], rows: list[Any]) -> list[str]:
    """O writer (ppt_chart_writer.py) so atualiza os <c:ser> que ja existem no
    template e descarta silenciosamente series excedentes. target.expected_series
    tem um item por <c:ser> encontrado no chart original, entao seu tamanho e a
    capacidade real de series que o template suporta. Aqui barramos ANTES da
    escrita uma matriz que perderia dados nesse truncamento silencioso."""
    template_series_slots = len(getattr(target, "expected_series", None) or [])
    if template_series_slots <= 0:
        return []
    orientation = getattr(target, "expected_orientation", "") or "categories_rows_series_columns"
    if orientation == "series_rows_categories_columns":
        actual_series_count = len(rows)
    else:
        actual_series_count = max(len(headers) - 1, 0)
    if actual_series_count > template_series_slots:
        return [
            f"A matriz proposta tem {actual_series_count} serie(s), mas o grafico do template so "
            f"comporta {template_series_slots} serie(s) existentes; as series excedentes seriam "
            "descartadas ao gravar o PowerPoint. Revise a orientacao ou o template."
        ]
    return []


def assert_valid_typed_edit_data(edit_data: dict[str, Any], object_type: str = "chart", target: Any = None) -> None:
    errors = validate_typed_edit_data(edit_data, object_type=object_type, target=target)
    if errors:
        raise EditDataValidationError(" ".join(errors))
