from __future__ import annotations

from io import BytesIO
import unittest
import xml.etree.ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

import openpyxl

from ppt_automator.xlsx_parser import _openpyxl_readable_copy, parse_xlsx_table


class XlsxPivotCompatibilityTests(unittest.TestCase):
    def test_parser_uses_cell_values_after_stripping_pivot_metadata_from_copy(self) -> None:
        original = _workbook_with_pivot_metadata()
        readable = _openpyxl_readable_copy(original)

        with ZipFile(BytesIO(readable)) as archive:
            self.assertFalse(any(name.startswith("xl/pivotCache/") for name in archive.namelist()))
            self.assertFalse(any(name.startswith("xl/pivotTables/") for name in archive.namelist()))
            self.assertNotIn(b"pivotCaches", archive.read("xl/workbook.xml"))
            self.assertNotIn(b"pivotCacheDefinition", archive.read("xl/_rels/workbook.xml.rels"))

        parsed = parse_xlsx_table(original, file_name="com-pivot.xlsx")
        self.assertEqual(parsed.categories, ["Total"])
        self.assertEqual(parsed.values, [[42]])


def _workbook_with_pivot_metadata() -> bytes:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["Indicador", "Valor"])
    sheet.append(["Total", 42])
    base = BytesIO()
    workbook.save(base)
    workbook.close()

    source = ZipFile(BytesIO(base.getvalue()))
    output = BytesIO()
    with source, ZipFile(output, "w", ZIP_DEFLATED) as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "xl/workbook.xml":
                root = ET.fromstring(data)
                namespace = root.tag.split("}", 1)[0].strip("{")
                rel_namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                caches = ET.SubElement(root, f"{{{namespace}}}pivotCaches")
                cache = ET.SubElement(caches, f"{{{namespace}}}pivotCache", {"cacheId": "1"})
                cache.set(f"{{{rel_namespace}}}id", "rId999")
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            elif info.filename == "xl/_rels/workbook.xml.rels":
                root = ET.fromstring(data)
                namespace = root.tag.split("}", 1)[0].strip("{")
                ET.SubElement(
                    root,
                    f"{{{namespace}}}Relationship",
                    {
                        "Id": "rId999",
                        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition",
                        "Target": "pivotCache/pivotCacheDefinition1.xml",
                    },
                )
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            target.writestr(info, data)
        target.writestr(
            "xl/pivotCache/pivotCacheDefinition1.xml",
            b'<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
        )
        target.writestr(
            "xl/pivotTables/pivotTable1.xml",
            b'<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
        )
    return output.getvalue()


if __name__ == "__main__":
    unittest.main()
