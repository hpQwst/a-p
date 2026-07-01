from __future__ import annotations

from decimal import Decimal, InvalidOperation
from typing import Any
import re


TypedCell = dict[str, Any]
TypedEditData = dict[str, Any]


TEXT_LIKE_RE = re.compile(
    r"^("
    r"\d{1,2}[/-]\d{2,4}|"
    r"\d{1,2}-\d{1,2}|"
    r"\d+Q\d{2,4}|"
    r"Q\d{1,2}\s*\d{2,4}|"
    r"[A-Za-zÀ-ÿ]{3}/\d{2,4}|"
    r"\d{2,}"
    r")$",
    flags=re.IGNORECASE,
)


def normalize_typed_edit_data(edit_data: dict[str, Any]) -> TypedEditData:
    headers = [
        normalize_typed_cell(cell, force_text_default=True)
        for cell in (edit_data.get("headers") or [])
    ]
    rows: list[list[TypedCell]] = []
    for raw_row in edit_data.get("rows") or []:
        if not isinstance(raw_row, list):
            continue
        row: list[TypedCell] = []
        for col_index, cell in enumerate(raw_row):
            row.append(normalize_typed_cell(cell, force_text_default=col_index == 0))
        rows.append(row)
    return {
        "headers": headers,
        "rows": rows,
        "matrix_preview": matrix_preview(headers, rows),
    }


def typed_edit_data_from_matrix(headers: list[Any], rows: list[list[Any]]) -> TypedEditData:
    return normalize_typed_edit_data({"headers": headers, "rows": rows})


def normalize_typed_cell(cell: Any, force_text_default: bool = False) -> TypedCell:
    if isinstance(cell, dict):
        value = cell.get("value", cell.get("raw", ""))
        cell_type = str(cell.get("type") or "").lower()
        force_text = bool(cell.get("force_text")) or force_text_default
        source_raw = str(cell.get("source_raw") or cell.get("raw") or value or "")
        if cell_type not in {"text", "number"}:
            cell_type = "text" if force_text or _looks_like_forced_text(value) else ("number" if _to_decimal(value) is not None else "text")
        if cell_type == "text":
            force_text = True
        return {
            "value": "" if value is None else str(value),
            "type": cell_type,
            "force_text": force_text,
            "source_raw": source_raw if cell_type == "number" else str(value or ""),
        }

    if force_text_default or _looks_like_forced_text(cell):
        return {"value": "" if cell is None else str(cell), "type": "text", "force_text": True, "source_raw": ""}
    raw = "" if cell is None else str(cell)
    if _to_decimal(raw) is not None:
        return {"value": raw, "type": "number", "force_text": False, "source_raw": raw}
    return {"value": raw, "type": "text", "force_text": True, "source_raw": ""}


def cell_to_python_value(cell: Any) -> Any:
    typed = normalize_typed_cell(cell)
    if typed["type"] == "number":
        raw = str(typed.get("source_raw") or typed.get("value") or "")
        number = _to_decimal(raw)
        if number is None:
            return typed.get("value", "")
        return float(number)
    return str(typed.get("value") or "")


def cell_to_excel_value(cell: Any) -> Any:
    return cell_to_python_value(cell)


def typed_matrix_to_rows(edit_data: dict[str, Any]) -> list[list[TypedCell]]:
    normalized = normalize_typed_edit_data(edit_data)
    headers = normalized.get("headers") or []
    rows = normalized.get("rows") or []
    return [headers, *rows] if headers else rows


def matrix_preview(headers: list[TypedCell], rows: list[list[TypedCell]]) -> list[list[str]]:
    return [
        [str(cell.get("value") or "") for cell in headers],
        *[[str(cell.get("value") or "") for cell in row] for row in rows],
    ]


def numeric_value(cell: Any) -> Decimal | None:
    typed = normalize_typed_cell(cell)
    if typed["type"] != "number":
        return None
    return _to_decimal(str(typed.get("source_raw") or typed.get("value") or ""))


def _looks_like_forced_text(value: Any) -> bool:
    text = "" if value is None else str(value).strip()
    if not text:
        return True
    if text.startswith("'"):
        return True
    if re.fullmatch(r"0\d+", text):
        return True
    return bool(TEXT_LIKE_RE.fullmatch(text))


def _to_decimal(value: Any) -> Decimal | None:
    text = "" if value is None else str(value).strip()
    if not text:
        return None
    text = text.replace("%", "").strip()
    text = re.sub(r"^,", "0,", text)
    text = re.sub(r"^-,", "-0,", text)
    text = text.replace(" ", "")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", ".")
    try:
        return Decimal(text)
    except (InvalidOperation, ValueError):
        return None
