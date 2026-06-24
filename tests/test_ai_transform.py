from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import json
import os
import unittest

from ppt_automator import analyze_update_package
from ppt_automator.ai_transform import suggest_transform_diagnostics


ANDRE_DIR = Path(os.getenv("AUTO_PPT_ANDRE_TEST_DIR", r"C:\Users\HugoRocha\Documents\automatizador-ppt-arquivos\andre"))
PPT = ANDRE_DIR / "Natura_2Q26_RelacionalCB_modelo_mapeado.pptx"
DATASOURCES = ANDRE_DIR / "datasources.zip"


@unittest.skipUnless(PPT.exists() and DATASOURCES.exists(), "Arquivos Andre de regressao nao encontrados.")
class AiTransformTests(unittest.TestCase):
    def test_ai_receives_ppt_edit_data_contract_and_xlsx_structure(self) -> None:
        _targets, _sources, plans = analyze_update_package(PPT, DATASOURCES)
        plan = next(item for item in plans if item.target_id == "7792738590")
        captured: dict[str, object] = {}

        class FakeResponses:
            def create(self, **kwargs):
                captured.update(kwargs)
                return SimpleNamespace(
                    output_text=json.dumps(
                        {
                            "diagnostics": [
                                {
                                    "target": "7792738590",
                                    "status": "ok",
                                    "confidence": 0.91,
                                    "action": "align",
                                    "reason": "Contrato e datasource usam os mesmos eixos.",
                                    "row_mapping": [
                                        {
                                            "ppt": "Não tem auto-consumo",
                                            "xlsx": "Não compro produtos para uso próprio",
                                        }
                                    ],
                                    "column_mapping": [{"ppt": "Total", "xlsx": "TOTAL"}],
                                }
                            ]
                        },
                        ensure_ascii=False,
                    )
                )

        class FakeOpenAI:
            def __init__(self, **_kwargs):
                self.responses = FakeResponses()

        with patch.dict(os.environ, {"OPENAI_API_KEY": "test-key", "OPENAI_MODEL": "test-model"}):
            with patch.dict("sys.modules", {"openai": SimpleNamespace(OpenAI=FakeOpenAI)}):
                diagnostics = suggest_transform_diagnostics([plan])

        self.assertEqual(diagnostics[0].target, "7792738590")
        self.assertEqual(diagnostics[0].action, "align")
        request_payload = json.loads(captured["input"][1]["content"])
        ai_plan = request_payload["plans"][0]
        self.assertEqual(ai_plan["ppt_edit_data_contract"]["orientation"], "series_rows_categories_columns")
        self.assertEqual(ai_plan["ppt_edit_data_contract"]["rows"], plan.series)
        self.assertEqual(ai_plan["ppt_edit_data_contract"]["columns"], plan.categories)
        self.assertEqual(ai_plan["xlsx_detected"]["orientation"], plan.datasource.orientation)
        self.assertEqual(ai_plan["proposed_transform"]["action"], "align")
        diagnostic_schema = captured["text"]["format"]["schema"]["properties"]["diagnostics"]["items"]["properties"]
        self.assertEqual(diagnostic_schema["row_mapping"]["type"], "array")
        self.assertEqual(diagnostic_schema["column_mapping"]["type"], "array")


if __name__ == "__main__":
    unittest.main()
