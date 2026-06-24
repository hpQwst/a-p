from __future__ import annotations

from io import BytesIO
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile
import posixpath
import re
import xml.etree.ElementTree as ET


SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"s": SHEET_NS}

ET.register_namespace("", SHEET_NS)


def update_embedded_workbook(workbook_bytes: bytes, sheet_name: str, matrix: list[list[Any]]) -> bytes:
    sheet_path = _worksheet_path(workbook_bytes, sheet_name)
    if not sheet_path:
        return workbook_bytes

    with ZipFile(BytesIO(workbook_bytes)) as zin:
        if sheet_path not in zin.namelist():
            return workbook_bytes

        sheet_xml = _updated_sheet_xml(zin.read(sheet_path), matrix)

        output = BytesIO()
        with ZipFile(output, "w", ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = sheet_xml if item.filename == sheet_path else zin.read(item.filename)
                zout.writestr(item, data)

        return output.getvalue()


def _worksheet_path(workbook_bytes: bytes, sheet_name: str) -> str:
    with ZipFile(BytesIO(workbook_bytes)) as zin:
        names = set(zin.namelist())
        if "xl/workbook.xml" not in names or "xl/_rels/workbook.xml.rels" not in names:
            return ""

        workbook = ET.fromstring(zin.read("xl/workbook.xml"))
        rels = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))

        sheets = workbook.findall("./s:sheets/s:sheet", NS)
        selected = next((sheet for sheet in sheets if sheet.attrib.get("name") == sheet_name), None)

        if selected is None and sheets:
            selected = sheets[0]

        if selected is None:
            return ""

        rel_id = selected.attrib.get(f"{{{R_NS}}}id")
        if not rel_id:
            return ""

        for rel in rels.findall(f"{{{REL_NS}}}Relationship"):
            if rel.attrib.get("Id") == rel_id:
                target = rel.attrib.get("Target", "")
                if target.startswith("/"):
                    return target.lstrip("/")
                return posixpath.normpath(posixpath.join("xl", target))

    return ""


def _updated_sheet_xml(sheet_xml: bytes, matrix: list[list[Any]]) -> bytes:
    root = ET.fromstring(sheet_xml)

    sheet_data = root.find("./s:sheetData", NS)
    if sheet_data is None:
        sheet_data = ET.Element(f"{{{SHEET_NS}}}sheetData")
        root.append(sheet_data)

    for row in list(sheet_data):
        sheet_data.remove(row)

    for row_idx, values in enumerate(matrix, 1):
        row_el = ET.Element(f"{{{SHEET_NS}}}row", {"r": str(row_idx)})
        for col_idx, value in enumerate(values, 1):
            row_el.append(_cell(row_idx, col_idx, value))
        sheet_data.append(row_el)

    _set_dimension(root, len(matrix), max((len(row) for row in matrix), default=1))

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _cell(row_idx: int, col_idx: int, value: Any) -> ET.Element:
    cell = ET.Element(f"{{{SHEET_NS}}}c", {"r": f"{_excel_col(col_idx)}{row_idx}"})

    if value is None:
        return cell

    number = _number(value)
    if number is not None and (not isinstance(value, str) or _numeric_text(value)):
        ET.SubElement(cell, f"{{{SHEET_NS}}}v").text = f"{number:.12g}"
        return cell

    cell.attrib["t"] = "inlineStr"
    inline = ET.SubElement(cell, f"{{{SHEET_NS}}}is")
    text = ET.SubElement(inline, f"{{{SHEET_NS}}}t")
    text.text = str(value)
    return cell


def _set_dimension(root: ET.Element, row_count: int, col_count: int) -> None:
    dimension = root.find("./s:dimension", NS)
    if dimension is None:
        dimension = ET.Element(f"{{{SHEET_NS}}}dimension")
        root.insert(0, dimension)

    dimension.attrib["ref"] = f"A1:{_excel_col(max(col_count, 1))}{max(row_count, 1)}"


def _number(value: Any) -> float | None:
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    is_percent = "%" in text
    text = text.replace("%", "").strip()
    text = re.sub(r"^,", "0,", text).replace(" ", "")

    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", ".")

    try:
        number = float(text)
    except ValueError:
        return None

    return number / 100 if is_percent else number


def _numeric_text(value: str) -> bool:
    return bool(re.fullmatch(r"-?[0-9.,% ]+", value.strip()))


def _excel_col(index: int) -> str:
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters