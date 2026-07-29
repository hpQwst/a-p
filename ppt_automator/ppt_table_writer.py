from __future__ import annotations

import copy
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

    if plan.table_matrix:
        _write_matrix(
            table,
            frame,
            plan.table_matrix,
            plan.number_format,
        )
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

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


def _write_matrix(
    table: ET.Element,
    frame: ET.Element,
    matrix: list[list[Any]],
    number_format: str,
) -> None:
    required_rows = len(matrix)
    required_cols = max((len(row) for row in matrix), default=0)
    if not required_rows or not required_cols:
        return
    _ensure_table_dimensions(table, frame, required_rows, required_cols)
    table_rows = table.findall("./a:tr", NS)
    for row_index, table_row in enumerate(table_rows):
        cells = table_row.findall("./a:tc", NS)
        values = matrix[row_index] if row_index < required_rows else []
        for column_index, cell in enumerate(cells):
            value = values[column_index] if column_index < len(values) else None
            _set_cell_text(cell, _format_value(value, number_format))


def _ensure_table_dimensions(
    table: ET.Element,
    frame: ET.Element,
    required_rows: int,
    required_cols: int,
) -> None:
    table_rows = table.findall("./a:tr", NS)
    if not table_rows:
        return

    original_row_count = len(table_rows)
    original_row_height = sum(_positive_int(row.attrib.get("h")) for row in table_rows)
    while len(table_rows) < required_rows:
        new_row = copy.deepcopy(table_rows[-1])
        for cell in new_row.findall("./a:tc", NS):
            _set_cell_text(cell, "")
        table.append(new_row)
        table_rows.append(new_row)

    existing_cols = max(
        (len(row.findall("./a:tc", NS)) for row in table_rows),
        default=0,
    )
    for row in table_rows:
        cells = row.findall("./a:tc", NS)
        if not cells:
            continue
        while len(cells) < required_cols:
            new_cell = copy.deepcopy(cells[-1])
            _set_cell_text(new_cell, "")
            _insert_before_ext_list(row, new_cell)
            cells.append(new_cell)

    grid = table.find("./a:tblGrid", NS)
    if grid is None:
        grid = ET.Element(f"{{{DML_NS}}}tblGrid")
        first_row_index = next(
            (index for index, child in enumerate(list(table)) if child.tag == f"{{{DML_NS}}}tr"),
            len(list(table)),
        )
        table.insert(first_row_index, grid)
    grid_columns = grid.findall("./a:gridCol", NS)
    original_grid_width = sum(_positive_int(column.attrib.get("w")) for column in grid_columns)
    desired_grid_cols = max(existing_cols, required_cols)
    while len(grid_columns) < desired_grid_cols:
        new_column = (
            copy.deepcopy(grid_columns[-1])
            if grid_columns
            else ET.Element(f"{{{DML_NS}}}gridCol")
        )
        grid.append(new_column)
        grid_columns.append(new_column)

    if required_cols > existing_cols:
        total_width = original_grid_width or _frame_extent(frame, "cx")
        _distribute_dimension(grid_columns[:required_cols], "w", total_width)
    if required_rows > original_row_count:
        total_height = original_row_height or _frame_extent(frame, "cy")
        _distribute_dimension(table_rows[:required_rows], "h", total_height)


def _insert_before_ext_list(parent: ET.Element, child: ET.Element) -> None:
    ext_tag = f"{{{DML_NS}}}extLst"
    children = list(parent)
    ext_index = next(
        (index for index, existing in enumerate(children) if existing.tag == ext_tag),
        len(children),
    )
    parent.insert(ext_index, child)


def _frame_extent(frame: ET.Element, attribute: str) -> int:
    extent = frame.find("./p:xfrm/a:ext", NS)
    return _positive_int(extent.attrib.get(attribute)) if extent is not None else 0


def _distribute_dimension(elements: list[ET.Element], attribute: str, total: int) -> None:
    if not elements or total <= 0:
        return
    base, remainder = divmod(total, len(elements))
    for index, element in enumerate(elements):
        element.attrib[attribute] = str(base + (1 if index < remainder else 0))


def _positive_int(value: Any) -> int:
    try:
        return max(0, int(value or 0))
    except (TypeError, ValueError):
        return 0


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
        if text == "":
            return
        run = ET.Element(f"{{{DML_NS}}}r")
        end_para = paragraph.find("./a:endParaRPr", NS)
        if end_para is not None:
            run_pr = copy.deepcopy(end_para)
            run_pr.tag = f"{{{DML_NS}}}rPr"
            run.append(run_pr)
            paragraph.insert(list(paragraph).index(end_para), run)
        else:
            paragraph.append(run)
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
