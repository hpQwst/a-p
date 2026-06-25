from __future__ import annotations

from io import BytesIO
import unittest
import xml.etree.ElementTree as ET

import openpyxl

from ppt_automator.ppt_discovery import PptTarget
from ppt_automator.ppt_table_writer import update_table_slide_xml
from ppt_automator.table_normalizer import normalize_to_target
from ppt_automator.xlsx_parser import parse_xlsx_table


NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


class TableKeyValueUpdateTests(unittest.TestCase):
    def test_key_value_xlsx_updates_second_column_without_none_text(self) -> None:
        source = parse_xlsx_table(_key_value_workbook(), file_name="t.xlsx")
        target = PptTarget(
            slide_index=0,
            slide_number=1,
            slide_path="ppt/slides/slide1.xml",
            shape_name="8282462966",
            shape_id="1",
            object_type="table",
            left_in=0,
            top_in=0,
            width_in=1,
            height_in=1,
            table_cells=[["Base:", ""], ["Total", ""], ["Natura", ""], ["Avon", ""]],
        )

        plan = normalize_to_target(target, source)
        updated = update_table_slide_xml(_slide_with_table(), target, plan)
        values = _table_values(updated)

        self.assertEqual(source.orientation, "key_value_rows")
        self.assertEqual(plan.values, [["Base:", ""], ["Total", 50], ["Natura", 20], ["Avon", 30]])
        self.assertEqual(values, [["Base:", ""], ["Total", "50"], ["Natura", "20"], ["Avon", "30"]])
        self.assertNotIn("None", updated.decode("utf-8"))


def _key_value_workbook() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Base:"
    ws["B1"] = None
    ws["A2"] = "Total"
    ws["B2"] = 50
    ws["A3"] = "Natura"
    ws["B3"] = 20
    ws["A4"] = "Avon"
    ws["B4"] = 30
    data = BytesIO()
    wb.save(data)
    return data.getvalue()


def _slide_with_table() -> bytes:
    rows = "\n".join(
        f"""
        <a:tr h="370840">
          <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{label}</a:t></a:r></a:p></a:txBody></a:tc>
          <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t></a:t></a:r></a:p></a:txBody></a:tc>
        </a:tr>
        """
        for label in ["Base:", "Total", "Natura", "Avon"]
    )
    return f"""
    <p:sld xmlns:p="{NS['p']}" xmlns:a="{NS['a']}">
      <p:cSld>
        <p:spTree>
          <p:graphicFrame>
            <p:nvGraphicFramePr><p:cNvPr id="1" name="8282462966"/></p:nvGraphicFramePr>
            <a:graphic><a:graphicData><a:tbl>{rows}</a:tbl></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree>
      </p:cSld>
    </p:sld>
    """.encode("utf-8")


def _table_values(slide_xml: bytes) -> list[list[str]]:
    root = ET.fromstring(slide_xml)
    return [
        ["".join(text.text or "" for text in cell.findall(".//a:t", NS)) for cell in row.findall("./a:tc", NS)]
        for row in root.findall(".//a:tbl/a:tr", NS)
    ]


if __name__ == "__main__":
    unittest.main()
