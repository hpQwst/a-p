from __future__ import annotations

from typing import Any

from .typed_matrix import normalize_typed_edit_data, numeric_value


class EditDataValidationError(ValueError):
    pass


def validate_typed_edit_data(edit_data: dict[str, Any], object_type: str = "chart") -> list[str]:
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
    return errors


def assert_valid_typed_edit_data(edit_data: dict[str, Any], object_type: str = "chart") -> None:
    errors = validate_typed_edit_data(edit_data, object_type=object_type)
    if errors:
        raise EditDataValidationError(" ".join(errors))
