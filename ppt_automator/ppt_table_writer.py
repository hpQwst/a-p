from __future__ import annotations

from typing import Any
import xml.etree.ElementTree as ET

from .ppt_discovery import DML_NS, NS, PML_NS, PptTarget
from .table_normalizer import TransformPlan


def update_table_slide_xml(slide_xml: bytes, target: PptTarget, plan: TransformPlan) -> bytes:
    root = ET.fromstring(slide_xml)
    frame = _find_graphic_frame(root, target.shape_name)
    if frame is None:
        return slide_xml
    table = frame.find(".//a:tbl", NS)
    if table is None:
        return slide_xml

    table_rows = table.findall("./a:tr", NS)
    cells_by_row = [row.findall("./a:tc", NS) for row in table_rows]
    values = plan.values
    if not values:
        return slide_xml

    if len(cells_by_row) == len(values) and all(len(cells_by_row[i]) >= len(values[i]) for i in range(len(values))):
        for row_index, row_values in enumerate(values):
            _write_row(cells_by_row[row_index], row_values, plan.number_format)
    elif len(cells_by_row) == 1:
        _write_row(cells_by_row[0], values[0], plan.number_format)
    elif len(cells_by_row) >= 2 and _matches_categories(cells_by_row[0], plan.categories):
        _write_row(cells_by_row[1], values[0], plan.number_format)
    else:
        flat_values = [value for row in values for value in row]
        flat_cells = [cell for row in cells_by_row for cell in row]
        _write_row(flat_cells, flat_values, plan.number_format)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _find_graphic_frame(root: ET.Element, shape_name: str) -> ET.Element | None:
    for frame in root.findall(".//p:graphicFrame", NS):
        cnv = frame.find("./p:nvGraphicFramePr/p:cNvPr", NS)
        if cnv is not None and cnv.attrib.get("name") == shape_name:
            return frame
    return None


def _write_row(cells: list[ET.Element], values: list[Any], number_format: str) -> None:
    for cell, value in zip(cells, values):
        _set_cell_text(cell, _format_value(value, number_format))


def _set_cell_text(cell: ET.Element, text: str) -> None:
    text_nodes = cell.findall(".//a:t", NS)
    if text_nodes:
        text_nodes[0].text = text
        for extra in text_nodes[1:]:
            extra.text = ""
        return

    tx_body = cell.find("./a:txBody", NS)
    if tx_body is None:
        tx_body = ET.SubElement(cell, f"{{{DML_NS}}}txBody")
        ET.SubElement(tx_body, f"{{{DML_NS}}}bodyPr")
        ET.SubElement(tx_body, f"{{{DML_NS}}}lstStyle")
    paragraph = tx_body.find("./a:p", NS)
    if paragraph is None:
        paragraph = ET.SubElement(tx_body, f"{{{DML_NS}}}p")
    run = paragraph.find("./a:r", NS)
    if run is None:
        run = ET.SubElement(paragraph, f"{{{DML_NS}}}r")
    text_node = run.find("./a:t", NS)
    if text_node is None:
        text_node = ET.SubElement(run, f"{{{DML_NS}}}t")
    text_node.text = text


def _matches_categories(cells: list[ET.Element], categories: list[str]) -> bool:
    cell_texts = [_norm("".join(node.text or "" for node in cell.findall(".//a:t", NS))) for cell in cells]
    category_texts = [_norm(category) for category in categories]
    if not cell_texts or not category_texts:
        return False
    matches = sum(1 for category in category_texts if category in cell_texts)
    return matches >= max(1, len(category_texts) // 2)


def _format_value(value: Any, number_format: str) -> str:
    if value is None:
        return ""
    if number_format == "thousands_pt_br":
        try:
            number = float(value)
            if number.is_integer():
                return f"{int(number):,}".replace(",", ".")
        except (TypeError, ValueError):
            pass
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value).replace(".", ",")
    return str(value)


def _norm(value: Any) -> str:
    text = "" if value is None else str(value).strip().upper()
    return text
