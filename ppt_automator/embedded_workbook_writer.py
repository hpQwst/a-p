from __future__ import annotations

from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path
from typing import Any
from zipfile import ZipFile
import copy
import xml.etree.ElementTree as ET

from .openxml_zip import replace_zip_parts_preserving_structure
from .typed_matrix import normalize_typed_cell


class EmbeddedWorkbookWriterUnavailable(RuntimeError):
    pass


SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
X14AC_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
XR_NS = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
XR2_NS = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
XR3_NS = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"
XR6_NS = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6"
XR10_NS = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"
X15_NS = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
XLWCV_NS = "http://schemas.microsoft.com/office/spreadsheetml/2024/workbookCompatibilityVersion"


for prefix, uri in {
    "": SHEET_NS,
    "r": DOC_REL_NS,
    "mc": MC_NS,
    "x14ac": X14AC_NS,
    "xr": XR_NS,
    "xr2": XR2_NS,
    "xr3": XR3_NS,
    "xr6": XR6_NS,
    "xr10": XR10_NS,
    "x15": X15_NS,
    "xlwcv": XLWCV_NS,
}.items():
    ET.register_namespace(prefix, uri)


def update_embedded_workbook(workbook_bytes: bytes, sheet_name: str, matrix: list[list[Any]]) -> bytes:
    row_count = len(matrix)
    col_count = max((len(row) for row in matrix), default=0)
    if row_count == 0 or col_count == 0:
        raise EmbeddedWorkbookWriterUnavailable("Matriz vazia; nada para gravar no workbook embutido.")

    replacements: dict[str, bytes] = {}
    with ZipFile(BytesIO(workbook_bytes)) as workbook_zip:
        sheet_path = _worksheet_path_for_sheet(workbook_zip, sheet_name)
        shared_strings_path = "xl/sharedStrings.xml"
        use_shared_strings = shared_strings_path in workbook_zip.namelist()

        replacements[sheet_path] = _updated_sheet_xml(
            workbook_zip.read(sheet_path),
            matrix,
            use_shared_strings=use_shared_strings,
        )
        if use_shared_strings:
            replacements[shared_strings_path] = _shared_strings_xml(_matrix_strings(matrix))

        table_path = _first_table_path(workbook_zip, sheet_path)
        if table_path:
            replacements[table_path] = _updated_table_xml(workbook_zip.read(table_path), matrix)

    return replace_zip_parts_preserving_structure(workbook_bytes, replacements)


def _worksheet_path_for_sheet(workbook_zip: ZipFile, sheet_name: str) -> str:
    workbook_root = ET.fromstring(workbook_zip.read("xl/workbook.xml"))
    rels_root = ET.fromstring(workbook_zip.read("xl/_rels/workbook.xml.rels"))
    rels = {
        rel.attrib.get("Id"): rel.attrib.get("Target", "")
        for rel in rels_root.findall(f"{{{REL_NS}}}Relationship")
    }
    sheets = workbook_root.find(f"{{{SHEET_NS}}}sheets")
    if sheets is None:
        raise EmbeddedWorkbookWriterUnavailable("Workbook embutido nao tem lista de abas.")

    selected_sheet = None
    for sheet in sheets.findall(f"{{{SHEET_NS}}}sheet"):
        if sheet_name and sheet.attrib.get("name") == sheet_name:
            selected_sheet = sheet
            break
        if selected_sheet is None:
            selected_sheet = sheet
    if selected_sheet is None:
        raise EmbeddedWorkbookWriterUnavailable("Workbook embutido nao tem abas.")

    rel_id = selected_sheet.attrib.get(f"{{{DOC_REL_NS}}}id")
    target = rels.get(rel_id or "")
    if not target:
        raise EmbeddedWorkbookWriterUnavailable(f"Nao encontrei relationship da aba '{selected_sheet.attrib.get('name')}'.")
    return _join_xlsx_path("xl/workbook.xml", target)


def _first_table_path(workbook_zip: ZipFile, sheet_path: str) -> str:
    rels_path = _rels_path(sheet_path)
    if rels_path not in workbook_zip.namelist():
        return ""
    rels_root = ET.fromstring(workbook_zip.read(rels_path))
    for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
        if str(rel.attrib.get("Type") or "").endswith("/table"):
            return _join_xlsx_path(sheet_path, str(rel.attrib.get("Target") or ""))
    return ""


def _updated_sheet_xml(sheet_xml: bytes, matrix: list[list[Any]], *, use_shared_strings: bool) -> bytes:
    root = ET.fromstring(sheet_xml)
    row_count = len(matrix)
    col_count = max((len(row) for row in matrix), default=0)
    dimension = root.find(f"{{{SHEET_NS}}}dimension")
    if dimension is not None:
        dimension.attrib["ref"] = f"A1:{_excel_col(col_count)}{row_count}"

    sheet_data = root.find(f"{{{SHEET_NS}}}sheetData")
    if sheet_data is None:
        sheet_data = ET.Element(f"{{{SHEET_NS}}}sheetData")
        root.append(sheet_data)
    for child in list(sheet_data):
        sheet_data.remove(child)

    string_index = {value: index for index, value in enumerate(_matrix_strings(matrix))} if use_shared_strings else {}
    for row_index, row_values in enumerate(matrix, start=1):
        row = ET.Element(
            f"{{{SHEET_NS}}}row",
            {
                "r": str(row_index),
                "spans": f"1:{col_count}",
                f"{{{X14AC_NS}}}dyDescent": "0.25",
            },
        )
        for col_index in range(1, col_count + 1):
            value = row_values[col_index - 1] if col_index - 1 < len(row_values) else ""
            cell_ref = f"{_excel_col(col_index)}{row_index}"
            cell = ET.Element(f"{{{SHEET_NS}}}c", {"r": cell_ref})
            if _matrix_cell_is_text(value, row_index=row_index, col_index=col_index):
                _write_text_cell(cell, _matrix_cell_text(value), string_index, use_shared_strings=use_shared_strings)
            else:
                v = ET.SubElement(cell, f"{{{SHEET_NS}}}v")
                v.text = _xml_number_text(_matrix_cell_number(value))
            row.append(cell)
        sheet_data.append(row)
    return _serialize_xml(root)


def _write_text_cell(cell: ET.Element, text: str, string_index: dict[str, int], *, use_shared_strings: bool) -> None:
    if use_shared_strings:
        cell.attrib["t"] = "s"
        v = ET.SubElement(cell, f"{{{SHEET_NS}}}v")
        v.text = str(string_index[text])
        return

    cell.attrib["t"] = "inlineStr"
    inline = ET.SubElement(cell, f"{{{SHEET_NS}}}is")
    text_node = ET.SubElement(inline, f"{{{SHEET_NS}}}t")
    if text != text.strip() or text == "":
        text_node.attrib["{http://www.w3.org/XML/1998/namespace}space"] = "preserve"
    text_node.text = text


def _updated_table_xml(table_xml: bytes, matrix: list[list[Any]]) -> bytes:
    root = ET.fromstring(table_xml)
    row_count = len(matrix)
    col_count = max((len(row) for row in matrix), default=0)
    root.attrib["ref"] = f"A1:{_excel_col(col_count)}{row_count}"
    auto_filter = root.find(f"{{{SHEET_NS}}}autoFilter")
    if auto_filter is not None:
        auto_filter.attrib["ref"] = root.attrib["ref"]

    table_columns = root.find(f"{{{SHEET_NS}}}tableColumns")
    if table_columns is None:
        table_columns = ET.SubElement(root, f"{{{SHEET_NS}}}tableColumns")
    old_columns = list(table_columns.findall(f"{{{SHEET_NS}}}tableColumn"))
    for child in old_columns:
        table_columns.remove(child)
    table_columns.attrib["count"] = str(col_count)
    header = matrix[0] if matrix else []
    for index in range(1, col_count + 1):
        template = copy.deepcopy(old_columns[index - 1]) if index - 1 < len(old_columns) else ET.Element(f"{{{SHEET_NS}}}tableColumn")
        template.attrib["id"] = str(index)
        template.attrib["name"] = _matrix_cell_text(header[index - 1] if index - 1 < len(header) else f"Column{index}") or " "
        table_columns.append(template)
    return _serialize_xml(root)


def _shared_strings_xml(strings: list[str]) -> bytes:
    root = ET.Element(f"{{{SHEET_NS}}}sst", {"count": str(len(strings)), "uniqueCount": str(len(strings))})
    for value in strings:
        si = ET.SubElement(root, f"{{{SHEET_NS}}}si")
        text = ET.SubElement(si, f"{{{SHEET_NS}}}t")
        if value != value.strip() or value == "":
            text.attrib["{http://www.w3.org/XML/1998/namespace}space"] = "preserve"
        text.text = value
    return _serialize_xml(root)


def _serialize_xml(root: ET.Element) -> bytes:
    _prune_ignorable_prefixes(root)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _prune_ignorable_prefixes(root: ET.Element) -> None:
    ignorable_key = f"{{{MC_NS}}}Ignorable"
    ignorable = root.attrib.get(ignorable_key)
    if not ignorable:
        return
    prefix_map = {
        "x14ac": X14AC_NS,
        "xr": XR_NS,
        "xr2": XR2_NS,
        "xr3": XR3_NS,
        "xr6": XR6_NS,
        "xr10": XR10_NS,
        "x15": X15_NS,
        "xlwcv": XLWCV_NS,
    }
    used_namespaces = _used_namespaces(root)
    kept = [prefix for prefix in ignorable.split() if prefix_map.get(prefix) in used_namespaces]
    if kept:
        root.attrib[ignorable_key] = " ".join(kept)
    else:
        root.attrib.pop(ignorable_key, None)


def _used_namespaces(root: ET.Element) -> set[str]:
    namespaces: set[str] = set()
    for element in root.iter():
        _collect_namespace(element.tag, namespaces)
        for attr_name in element.attrib:
            _collect_namespace(attr_name, namespaces)
    return namespaces


def _collect_namespace(name: str, namespaces: set[str]) -> None:
    if name.startswith("{") and "}" in name:
        namespaces.add(name[1:].split("}", 1)[0])


def _matrix_strings(matrix: list[list[Any]]) -> list[str]:
    output: list[str] = []
    seen: set[str] = set()
    for row_index, row in enumerate(matrix, start=1):
        for col_index, value in enumerate(row, start=1):
            if not _matrix_cell_is_text(value, row_index=row_index, col_index=col_index):
                continue
            text = _matrix_cell_text(value)
            if text not in seen:
                seen.add(text)
                output.append(text)
    return output


def _matrix_cell_is_text(value: Any, *, row_index: int, col_index: int) -> bool:
    if isinstance(value, dict):
        typed = normalize_typed_cell(value, force_text_default=row_index == 1 or col_index == 1)
        return typed["type"] == "text" or bool(typed.get("force_text"))
    return row_index == 1 or col_index == 1


def _matrix_cell_text(value: Any) -> str:
    if isinstance(value, dict):
        typed = normalize_typed_cell(value)
        return str(typed.get("value") or "")
    return "" if value is None else str(value)


def _matrix_cell_number(value: Any) -> Any:
    if isinstance(value, dict):
        typed = normalize_typed_cell(value)
        return typed.get("source_raw") or typed.get("value")
    return value


def _xml_number_text(value: Any) -> str:
    if value is None or value == "":
        return "0"
    number = _number_or_none(value)
    if number is None:
        return "0"
    return f"{float(number):.15g}"


def _number_or_none(value: Any) -> Decimal | None:
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    text = str(value).strip()
    if not text:
        return None
    text = text.replace("%", "").replace(" ", "")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", ".")
    try:
        return Decimal(text)
    except (InvalidOperation, ValueError):
        return None


def _excel_col(index: int) -> str:
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _join_xlsx_path(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base_dir = str(Path(base_part).parent).replace("\\", "/")
    parts: list[str] = []
    for part in f"{base_dir}/{target}".split("/"):
        if part in {"", "."}:
            continue
        if part == "..":
            if parts:
                parts.pop()
            continue
        parts.append(part)
    return "/".join(parts)


def _rels_path(part_name: str) -> str:
    path = Path(part_name)
    return str(path.parent / "_rels" / f"{path.name}.rels").replace("\\", "/")
