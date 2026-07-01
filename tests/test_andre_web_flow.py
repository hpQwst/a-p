from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
import os
import re
import unittest
import xml.etree.ElementTree as ET
from unittest.mock import patch

import openpyxl
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


@unittest.skipUnless(PPT.exists() and DATASOURCES.exists(), "Arquivos Andre de regressao nao encontrados.")
class AndreWebFlowTests(unittest.TestCase):
    def test_andre_package_analyzes_without_none_comparison_error(self) -> None:
        targets, sources, plans = analyze_update_package(PPT, DATASOURCES)
        self.assertGreaterEqual(len(targets), 12)
        self.assertEqual(len(sources), 12)
        self.assertEqual(len(plans), 12)
        output = generate_updated_pptx(PPT, plans)
        self.assertGreater(len(output), 1_000_000)

    def test_generated_chart_workbook_keeps_edit_data_package_valid(self) -> None:
        _targets, _sources, plans = analyze_update_package(PPT, DATASOURCES)
        plan = next(item for item in plans if item.target.shape_name == "7792738590")
        output = generate_updated_pptx(PPT, plans)
        with ZipFile(BytesIO(output)) as pptx:
            workbook_bytes = pptx.read(plan.target.workbook_embedded)

        with ZipFile(BytesIO(workbook_bytes)) as workbook:
            sheet_xml = workbook.read("xl/worksheets/sheet1.xml")
            table_xml = workbook.read("xl/tables/table1.xml")

        workbook = openpyxl.load_workbook(BytesIO(workbook_bytes), data_only=True)
        worksheet = workbook.worksheets[0]
        self.assertEqual(worksheet["A1"].value, " ")
        workbook.close()

        sheet = ET.fromstring(sheet_xml)
        self.assertEqual(sheet.find("./s:dimension", NS).attrib["ref"], "A1:D6")
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
        with patch.dict(os.environ, {"AUTO_PPT_AUTO_SLIDE_AI": "0"}), patch("web.main.ai_configured", return_value=False):
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
            self.assertIn("Sem datasource automático", response.text)
            match = re.search(r"/jobs/([a-f0-9]+)/download", response.text)
            self.assertIsNotNone(match)
            job_id = match.group(1)
            slide_three = client.get(f"/jobs/{job_id}/preview?slide=3")
            self.assertEqual(slide_three.status_code, 200)
            self.assertIn("7792738590", slide_three.text)

            with ZipFile(DATASOURCES) as zf:
                manual_source = next(
                    name
                    for name in zf.namelist()
                    if name.endswith("7792738590.xlsx") or name.endswith("tab7590_slide3.xlsx")
                )
                manual_data = zf.read(manual_source)
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
            self.assertEqual(download.status_code, 200)
            self.assertGreater(len(download.content), 1_000_000)

    def assertNoBrokenIgnorablePrefixes(self, xml_bytes: bytes) -> None:
        prefix_names = {item[0] for _event, item in ET.iterparse(BytesIO(xml_bytes), events=("start-ns",))}
        root = ET.fromstring(xml_bytes)
        ignorable = root.attrib.get(f"{{{MC_NS}}}Ignorable", "")
        missing = [prefix for prefix in ignorable.split() if prefix not in prefix_names]
        self.assertEqual(missing, [])


if __name__ == "__main__":
    unittest.main()
