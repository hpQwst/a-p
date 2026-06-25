from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
import os
import re
import unittest
import xml.etree.ElementTree as ET

from fastapi.testclient import TestClient

from ppt_automator import analyze_update_package, generate_updated_pptx
from web.main import app


ANDRE_DIR = Path(os.getenv("AUTO_PPT_ANDRE_TEST_DIR", r"C:\Users\HugoRocha\Documents\automatizador-ppt-arquivos\andre"))
PPT = ANDRE_DIR / "Natura_2Q26_RelacionalCB_modelo_mapeado.pptx"
DATASOURCES = ANDRE_DIR / "datasources.zip"
MAPPING = ANDRE_DIR / "Natura_2Q26_RelacionalCB_modelo_mapeamento.xlsx"
SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"s": SHEET_NS}


def excel_com_available() -> bool:
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.Quit()
        pythoncom.CoUninitialize()
        return True
    except Exception:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        return False


@unittest.skipUnless(PPT.exists() and DATASOURCES.exists(), "Arquivos Andre de regressao nao encontrados.")
class AndreWebFlowTests(unittest.TestCase):
    def test_andre_package_analyzes_without_none_comparison_error(self) -> None:
        targets, sources, plans = analyze_update_package(PPT, DATASOURCES)
        self.assertGreaterEqual(len(targets), 12)
        self.assertEqual(len(sources), 12)
        self.assertEqual(len(plans), 12)
        if not excel_com_available():
            self.skipTest("Excel COM indisponivel neste ambiente.")
        output = generate_updated_pptx(PPT, plans)
        self.assertGreater(len(output), 1_000_000)

    def test_generated_chart_workbook_keeps_edit_data_package_valid(self) -> None:
        if not excel_com_available():
            self.skipTest("Excel COM indisponivel neste ambiente.")
        _targets, _sources, plans = analyze_update_package(PPT, DATASOURCES)
        plan = next(item for item in plans if item.target_id == "7792738590")
        output = generate_updated_pptx(PPT, plans)
        with ZipFile(BytesIO(output)) as pptx:
            workbook_bytes = pptx.read(plan.target.workbook_embedded)

        with ZipFile(BytesIO(workbook_bytes)) as workbook:
            sheet_xml = workbook.read("xl/worksheets/sheet1.xml")
            table_xml = workbook.read("xl/tables/table1.xml")

        sheet = ET.fromstring(sheet_xml)
        self.assertEqual(sheet.find("./s:dimension", NS).attrib["ref"], "A1:D6")
        self.assertEqual(sheet.find("./s:sheetData/s:row/s:c/s:is/s:t", NS).text, " ")
        self.assertNoBrokenIgnorablePrefixes(sheet_xml)

        table = ET.fromstring(table_xml)
        self.assertEqual(table.attrib["ref"], "A1:D6")
        auto_filter = table.find("./s:autoFilter", NS)
        if auto_filter is not None:
            self.assertEqual(auto_filter.attrib["ref"], "A1:D6")
        self.assertEqual(table.find("./s:tableColumns", NS).attrib["count"], "4")
        self.assertNoBrokenIgnorablePrefixes(table_xml)

    def test_fastapi_preview_download_and_target_override(self) -> None:
        client = TestClient(app)
        files = {
            "pptx": (
                PPT.name,
                PPT.read_bytes(),
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ),
            "datasources": (
                DATASOURCES.name,
                DATASOURCES.read_bytes(),
                "application/zip",
            ),
        }
        if MAPPING.exists():
            files["mapping"] = (
                MAPPING.name,
                MAPPING.read_bytes(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        response = client.post(
            "/preview",
            data={
                "project_ref": "",
                "squad": "squad1",
                "project_name": "Andre regressao",
                "confirm_large_deck": "1",
            },
            files=files,
        )
        self.assertEqual(response.status_code, 200)
        self.assertIn("7792738590", response.text)
        self.assertIn("Sem datasource automático", response.text)
        match = re.search(r"/jobs/([a-f0-9]+)/download", response.text)
        self.assertIsNotNone(match)
        job_id = match.group(1)

        with ZipFile(DATASOURCES) as zf:
            manual_data = zf.read("datasources/7792738590.xlsx")
        override = client.post(
            f"/jobs/{job_id}/targets/6889461846/override",
            files={
                "datasource": (
                    "7792738590.xlsx",
                    manual_data,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            },
        )
        self.assertEqual(override.status_code, 200)
        self.assertIn("Correção manual", override.text)

        download = client.get(f"/jobs/{job_id}/download")
        if excel_com_available():
            self.assertEqual(download.status_code, 200)
            self.assertGreater(len(download.content), 1_000_000)
        else:
            self.assertEqual(download.status_code, 500)
            self.assertIn("Excel", download.text)

    def assertNoBrokenIgnorablePrefixes(self, xml_bytes: bytes) -> None:
        prefix_names = {item[0] for _event, item in ET.iterparse(BytesIO(xml_bytes), events=("start-ns",))}
        root = ET.fromstring(xml_bytes)
        ignorable = root.attrib.get(f"{{{MC_NS}}}Ignorable", "")
        missing = [prefix for prefix in ignorable.split() if prefix not in prefix_names]
        self.assertEqual(missing, [])


if __name__ == "__main__":
    unittest.main()
