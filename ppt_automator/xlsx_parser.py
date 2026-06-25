from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Any, BinaryIO
from zipfile import ZipFile
import re
import unicodedata

import openpyxl
from openpyxl.utils.cell import range_boundaries

from .core import prepare_workbook_values, read_bytes


InputFile = str | Path | bytes | bytearray | BinaryIO


@dataclass(frozen=True)
class ParsedXlsxTable:
    source_id: str
    file_name: str
    sheet_name: str
    orientation: str
    categories: list[str]
    series: list[str]
    values: list[list[Any]]
    used_range: tuple[int, int, int, int] | None = None
    metadata: dict[str, str] = field(default_factory=dict)
    preview_rows: list[list[Any]] = field(default_factory=list)

    @property
    def graph_id(self) -> str:
        return self.source_id


def parse_datasource_zip(datasources_zip: InputFile, formula_mode: str = "auto") -> list[ParsedXlsxTable]:
    output: list[ParsedXlsxTable] = []
    with ZipFile(BytesIO(read_bytes(datasources_zip))) as zf:
        for name in zf.namelist():
            if name.lower().endswith(".xlsx"):
                output.append(
                    parse_xlsx_table(
                        zf.read(name),
                        file_name=name,
                        formula_mode=formula_mode,
                    )
                )
    return output


def parse_xlsx_table(
    workbook_file: InputFile,
    file_name: str = "",
    formula_mode: str = "auto",
    cell_range: str = "",
) -> ParsedXlsxTable:
    original_bytes = read_bytes(workbook_file)
    calculated_bytes = prepare_workbook_values(original_bytes, formula_mode=formula_mode)
    data_wb = openpyxl.load_workbook(BytesIO(calculated_bytes), data_only=True, read_only=True)
    formula_wb = openpyxl.load_workbook(BytesIO(original_bytes), data_only=False, read_only=True)
    data_ws, range_ref = _select_worksheet(data_wb, cell_range)
    formula_ws = formula_wb[data_ws.title] if data_ws.title in formula_wb.sheetnames else formula_wb.worksheets[0]
    if range_ref:
        trimmed_rows, used_range = _rows_from_range(data_ws, range_ref)
        formula_rows, _formula_used_range = _rows_from_range(formula_ws, range_ref)
    else:
        raw_rows = [list(row) for row in data_ws.iter_rows(values_only=True)]
        formula_all_rows = [list(row) for row in formula_ws.iter_rows(values_only=True)]
        trimmed_rows, used_range = _trim_table(raw_rows)
        formula_rows = _slice_rows(formula_all_rows, used_range)
    metadata = _extract_metadata(trimmed_rows)
    parsed = _parse_rectangular_table(
        trimmed_rows,
        formula_rows,
        file_name=file_name,
        sheet_name=data_ws.title,
        used_range=used_range,
        metadata=metadata,
    )
    data_wb.close()
    formula_wb.close()
    return parsed


def _select_worksheet(workbook: Any, cell_range: str) -> tuple[Any, str]:
    sheet_name, range_ref = _split_range_ref(cell_range)
    if sheet_name:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"Aba '{sheet_name}' nao encontrada no XLSX.")
        return workbook[sheet_name], range_ref
    return workbook.worksheets[0], range_ref


def _split_range_ref(cell_range: str) -> tuple[str, str]:
    text = (cell_range or "").strip().replace("$", "")
    if not text:
        return "", ""
    if "!" in text:
        sheet_name, range_ref = text.split("!", 1)
        sheet_name = sheet_name.strip().strip("'")
    else:
        sheet_name, range_ref = "", text
    try:
        range_boundaries(range_ref)
    except Exception as exc:
        raise ValueError("Range invalido. Use algo como D5:G12 ou Planilha1!D5:G12.") from exc
    return sheet_name, range_ref


def _rows_from_range(worksheet: Any, range_ref: str) -> tuple[list[list[Any]], tuple[int, int, int, int]]:
    min_col, min_row, max_col, max_row = range_boundaries(range_ref)
    rows = [
        list(row)
        for row in worksheet.iter_rows(
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            values_only=True,
        )
    ]
    trimmed_rows, trimmed_range = _trim_table(rows)
    if trimmed_range is None:
        return [], (min_row, min_col, max_row, max_col)
    trim_min_row, trim_min_col, trim_max_row, trim_max_col = trimmed_range
    absolute_range = (
        min_row + trim_min_row - 1,
        min_col + trim_min_col - 1,
        min_row + trim_max_row - 1,
        min_col + trim_max_col - 1,
    )
    return trimmed_rows, absolute_range


def _slice_rows(rows: list[list[Any]], used_range: tuple[int, int, int, int] | None) -> list[list[Any]]:
    if used_range is None:
        return []
    min_row, min_col, max_row, max_col = used_range
    output = []
    for row in rows[min_row - 1 : max_row]:
        output.append(list(row[min_col - 1 : max_col]))
    return output


def _parse_rectangular_table(
    rows: list[list[Any]],
    formula_rows: list[list[Any]],
    file_name: str,
    sheet_name: str,
    used_range: tuple[int, int, int, int] | None,
    metadata: dict[str, str],
) -> ParsedXlsxTable:
    if not rows:
        return ParsedXlsxTable(
            source_id=_graph_id(Path(file_name).stem),
            file_name=file_name,
            sheet_name=sheet_name,
            orientation="empty",
            categories=[],
            series=[],
            values=[],
            used_range=used_range,
            metadata=metadata,
        )

    header_row_index, header_start_col, header_end_col = _find_header_row(rows)
    if header_row_index is None:
        key_value = _parse_key_value_rows(rows)
        if key_value:
            categories, values = key_value
            return ParsedXlsxTable(
                source_id=_graph_id(Path(file_name).stem),
                file_name=file_name,
                sheet_name=sheet_name,
                orientation="key_value_rows",
                categories=categories,
                series=["Valor"],
                values=values,
                used_range=used_range,
                metadata=metadata,
                preview_rows=_preview_rows(categories, ["Valor"], values, "categories_rows_series_columns"),
            )
        return ParsedXlsxTable(
            source_id=_graph_id(Path(file_name).stem),
            file_name=file_name,
            sheet_name=sheet_name,
            orientation="unknown",
            categories=[],
            series=[],
            values=[],
            used_range=used_range,
            metadata=metadata,
            preview_rows=rows[:10],
        )

    label_col = _find_label_col(rows, header_row_index, header_start_col)
    header_values = [_text(value) for value in rows[header_row_index][header_start_col : header_end_col + 1]]
    data_items: list[tuple[str, list[Any], list[Any]]] = []
    for row_index in range(header_row_index + 1, len(rows)):
        row = rows[row_index]
        values = list(row[header_start_col : header_end_col + 1])
        if not _has_numeric_or_text(values):
            continue
        label = _text(row[label_col] if label_col is not None and label_col < len(row) else "")
        formula_values = []
        if row_index < len(formula_rows):
            formula_row = formula_rows[row_index]
            formula_values = list(formula_row[header_start_col : header_end_col + 1])
        if not label and _looks_like_nps_row(values, formula_values, data_items):
            label = "NPS"
        data_items.append((label, values, formula_values))

    categories_in_header = _period_score(header_values) >= 0.45
    labels = [item[0] for item in data_items]
    labels_are_categories = _period_score(labels) >= 0.45

    if categories_in_header and len(data_items) == 1:
        orientation = "single_series_row_categories_columns"
        series = [labels[0] or "Valor"]
        values = [data_items[0][1]]
        categories = header_values
    elif categories_in_header and not labels_are_categories:
        orientation = "series_rows_categories_columns"
        categories = header_values
        series = [_series_label(label, values, formulas, index) for index, (label, values, formulas) in enumerate(data_items)]
        values = [item[1] for item in data_items]
    else:
        orientation = "categories_rows_series_columns"
        categories = [label or f"Linha {index + 1}" for index, (label, _values, _formulas) in enumerate(data_items)]
        series = header_values
        values = [item[1] for item in data_items]

    return ParsedXlsxTable(
        source_id=_graph_id(Path(file_name).stem),
        file_name=file_name,
        sheet_name=sheet_name,
        orientation=orientation,
        categories=categories,
        series=series,
        values=values,
        used_range=used_range,
        metadata=metadata,
        preview_rows=_preview_rows(categories, series, values, orientation),
    )


def _parse_key_value_rows(rows: list[list[Any]]) -> tuple[list[str], list[list[Any]]] | None:
    if len(rows) < 2:
        return None
    max_cols = max((len(row) for row in rows), default=0)
    if max_cols < 2:
        return None

    best: tuple[int, int, list[str], list[list[Any]]] | None = None
    for label_col in range(max_cols - 1):
        value_col = label_col + 1
        categories: list[str] = []
        values: list[list[Any]] = []
        for row in rows:
            label = _text(row[label_col] if label_col < len(row) else "")
            if not label or _to_number(label) is not None:
                continue
            value = row[value_col] if value_col < len(row) else None
            categories.append(label)
            values.append([value])
        value_count = sum(1 for row in values if _text(row[0]) or _to_number(row[0]) is not None)
        if len(categories) >= 2 and value_count >= 1:
            score = len(categories) + value_count
            if best is None or score > best[0]:
                best = (score, value_col, categories, values)
    if best is None:
        return None
    return best[2], best[3]


def _find_header_row(rows: list[list[Any]]) -> tuple[int | None, int, int]:
    best: tuple[int | None, int, int, int] = (None, 0, 0, -1)
    for row_index, row in enumerate(rows):
        header_cols = [col for col, value in enumerate(row) if _is_header_cell(value)]
        if len(header_cols) < 2:
            continue
        start = header_cols[0]
        end = header_cols[-1]
        if start == 0 and len(header_cols) > 1:
            if not _looks_like_period(row[0]) and not _is_likely_header_label(row[0]):
                start = header_cols[1]
        data_score = 0
        for next_row in rows[row_index + 1 : row_index + 5]:
            data_score += sum(2 for value in next_row[start : end + 1] if _to_number(value) is not None)
            data_score += sum(1 for value in next_row[start : end + 1] if _is_value(value))
        header_numeric_penalty = sum(3 for value in row[start : end + 1] if _to_number(value) is not None)
        score = len(header_cols) + data_score - header_numeric_penalty
        if score > best[3]:
            best = (row_index, start, end, score)
    return best[0], best[1], best[2]


def _find_label_col(rows: list[list[Any]], header_row_index: int, header_start_col: int) -> int | None:
    best_col = None
    best_score = 0
    for col in range(header_start_col):
        score = 0
        for row in rows[header_row_index + 1 :]:
            value = row[col] if col < len(row) else None
            if _text(value) and _to_number(value) is None:
                score += 1
        if score > best_score:
            best_col = col
            best_score = score
    if best_col is not None:
        return best_col
    return header_start_col - 1 if header_start_col > 0 else None


def _trim_table(rows: list[list[Any]]) -> tuple[list[list[Any]], tuple[int, int, int, int] | None]:
    positions: list[tuple[int, int]] = []
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            if _text(value) or _is_value(value):
                positions.append((r, c))
    if not positions:
        return [], None
    min_row = min(r for r, _c in positions)
    max_row = max(r for r, _c in positions)
    min_col = min(c for _r, c in positions)
    max_col = max(c for _r, c in positions)
    trimmed = []
    for row in rows[min_row : max_row + 1]:
        trimmed.append(list(row[min_col : max_col + 1]))
    return trimmed, (min_row + 1, min_col + 1, max_row + 1, max_col + 1)


def _preview_rows(categories: list[str], series: list[str], values: list[list[Any]], orientation: str) -> list[list[Any]]:
    if orientation in {"series_rows_categories_columns", "single_series_row_categories_columns"}:
        return [["", *categories], *[[series[i] if i < len(series) else "", *row] for i, row in enumerate(values)]]
    return [["", *series], *[[categories[i] if i < len(categories) else "", *row] for i, row in enumerate(values)]]


def _series_label(label: str, values: list[Any], formulas: list[Any], index: int) -> str:
    if label:
        return label
    if _looks_like_nps_row(values, formulas, []):
        return "NPS"
    return f"Serie {index + 1}"


def _looks_like_nps_row(values: list[Any], formulas: list[Any], previous: list[tuple[str, list[Any], list[Any]]]) -> bool:
    formula_text = " ".join(_text(value) for value in formulas)
    if re.search(r"\([A-Z]+\d+-[A-Z]+\d+\)\*100", formula_text, flags=re.IGNORECASE):
        return True
    numeric = [_to_number(value) for value in values if _to_number(value) is not None]
    if not numeric:
        return False
    previous_values = [
        number
        for _label, row_values, _formula_values in previous
        for number in (_to_number(value) for value in row_values)
        if number is not None
    ]
    previous_percent_like = previous_values and sum(1 for value in previous_values if -1 <= value <= 1) >= len(previous_values) * 0.65
    return previous_percent_like and sum(1 for value in numeric if abs(value) > 1) >= len(numeric) * 0.65


def _extract_metadata(rows: list[list[Any]]) -> dict[str, str]:
    metadata: dict[str, str] = {}
    aliases = {
        "GRAPH ID": "graph_id",
        "GRAPHID": "graph_id",
        "ID GRAFICO": "graph_id",
        "PPT TAG": "ppt_tag",
        "TAG PPT": "ppt_tag",
        "VARIAVEL": "variable",
        "VAR ANALISE": "variable",
    }
    for row in rows[:20]:
        cells = [_text(value) for value in row]
        for index, cell in enumerate(cells):
            if not cell:
                continue
            key_text = cell
            value_text = cells[index + 1] if index + 1 < len(cells) else ""
            if ":" in cell:
                key_text, value_text = [part.strip() for part in cell.split(":", 1)]
            key = aliases.get(_norm(key_text))
            if key and value_text:
                metadata[key] = value_text
    return metadata


def _period_score(values: list[Any]) -> float:
    clean = [_text(value) for value in values if _text(value)]
    if not clean:
        return 0.0
    return sum(1 for value in clean if _looks_like_period(value)) / len(clean)


def _looks_like_period(value: Any) -> bool:
    text = _norm(value)
    if re.fullmatch(r"(JAN|FEV|FEB|MAR|ABR|APR|MAI|MAY|JUN|JUL|AGO|AUG|SET|SEP|OUT|OCT|NOV|DEZ|DEC)\s*\d{2,4}", text):
        return True
    if re.fullmatch(r"\d{1,2}[/-]\d{2,4}", text):
        return True
    return False


def _is_likely_header_label(value: Any) -> bool:
    text = _norm(value)
    return text in {"SERIE", "SERIES", "CATEGORIA", "CATEGORIAS", "MES", "MESES"}


def _is_header_cell(value: Any) -> bool:
    text = _text(value)
    if not text:
        return False
    return _to_number(text) is None


def _has_numeric_or_text(values: list[Any]) -> bool:
    return any(_is_value(value) for value in values)


def _is_value(value: Any) -> bool:
    if isinstance(value, (int, float)):
        return True
    text = _text(value)
    if not text:
        return False
    return _to_number(text) is not None or len(text) > 0


def _to_number(value: Any) -> float | None:
    if isinstance(value, (int, float)):
        return float(value)
    text = _text(value)
    if not text:
        return None
    text = text.replace("%", "")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def _graph_id(value: Any) -> str:
    text = _text(value)
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return re.sub(r"\D", "", text)


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _norm(value: Any) -> str:
    text = _text(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.upper()
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()
