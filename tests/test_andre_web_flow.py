from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile
import os
import re
import unittest

from fastapi.testclient import TestClient

from ppt_automator import analyze_update_package, generate_updated_pptx
from web.main import app


ANDRE_DIR = Path(os.getenv("AUTO_PPT_ANDRE_TEST_DIR", r"C:\Users\HugoRocha\Documents\automatizador-ppt-arquivos\andre"))
PPT = ANDRE_DIR / "Natura_2Q26_RelacionalCB_modelo_mapeado.pptx"
DATASOURCES = ANDRE_DIR / "datasources.zip"
MAPPING = ANDRE_DIR / "Natura_2Q26_RelacionalCB_modelo_mapeamento.xlsx"


@unittest.skipUnless(PPT.exists() and DATASOURCES.exists(), "Arquivos Andre de regressao nao encontrados.")
class AndreWebFlowTests(unittest.TestCase):
    def test_andre_package_analyzes_without_none_comparison_error(self) -> None:
        targets, sources, plans = analyze_update_package(PPT, DATASOURCES)
        self.assertGreaterEqual(len(targets), 12)
        self.assertEqual(len(sources), 12)
        self.assertEqual(len(plans), 12)
        output = generate_updated_pptx(PPT, plans)
        self.assertGreater(len(output), 1_000_000)

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
            data={"project_ref": "", "squad": "squad1", "project_name": "Andre regressao"},
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
        self.assertEqual(download.status_code, 200)
        self.assertGreater(len(download.content), 1_000_000)


if __name__ == "__main__":
    unittest.main()
