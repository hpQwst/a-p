from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import patch
import json
import os
import unittest

from ppt_automator.ai_mapper import suggest_source_matches_with_ai
from ppt_automator.ppt_discovery import PptTarget
from ppt_automator.xlsx_parser import ParsedXlsxTable


class AiMapperTests(unittest.TestCase):
    def test_ai_matcher_sends_compact_candidates_and_accepts_valid_suggestion(self) -> None:
        target = PptTarget(
            slide_index=0,
            slide_number=1,
            slide_path="ppt/slides/slide1.xml",
            shape_name="123456",
            shape_id="7",
            object_type="chart",
            left_in=0,
            top_in=0,
            width_in=1,
            height_in=1,
            nearby_text="Auto-consumo por rede",
            expected_orientation="categories_rows_series_columns",
            expected_categories=["Total", "Rede 1", "Rede 2"],
            expected_series=["Ate 25%", "Entre 26% e 50%"],
            expected_values=[[0.1, 0.2], [0.3, 0.4], [0.5, 0.6]],
        )
        sources = [
            ParsedXlsxTable(
                source_id="",
                file_name="auto_consumo.xlsx",
                sheet_name="Sheet1",
                orientation="categories_rows_series_columns",
                categories=["TOTAL", "REDE 1", "REDE 2"],
                series=["Ate 25% das compras", "Entre 26% e 50%"],
                values=[[0.1, 0.2], [0.3, 0.4], [0.5, 0.6]],
                preview_rows=[["", "Ate 25% das compras", "Entre 26% e 50%"], ["TOTAL", 0.1, 0.2]],
            ),
            ParsedXlsxTable(
                source_id="",
                file_name="status.xlsx",
                sheet_name="Sheet1",
                orientation="categories_rows_series_columns",
                categories=["Ativa", "Disponivel"],
                series=["Total"],
                values=[[1], [2]],
                preview_rows=[["", "Total"], ["Ativa", 1]],
            ),
        ]
        captured: dict[str, object] = {}

        class FakeResponses:
            def create(self, **kwargs):
                captured.update(kwargs)
                return SimpleNamespace(
                    output_text=json.dumps(
                        {
                            "suggestions": [
                                {
                                    "target": "123456",
                                    "datasource": "auto_consumo.xlsx",
                                    "confidence": 0.82,
                                    "reason": "Categorias e series batem com o contrato do grafico.",
                                }
                            ]
                        }
                    )
                )

        class FakeOpenAI:
            def __init__(self, **_kwargs):
                self.responses = FakeResponses()

        with patch.dict(os.environ, {"OPENAI_API_KEY": "test-key", "OPENAI_MODEL": "test-model"}):
            with patch.dict("sys.modules", {"openai": SimpleNamespace(OpenAI=FakeOpenAI)}):
                suggestions = suggest_source_matches_with_ai([target], sources, candidates_per_target=1)

        self.assertEqual(len(suggestions), 1)
        self.assertEqual(suggestions[0].datasource, "auto_consumo.xlsx")
        request_payload = json.loads(captured["input"][1]["content"])
        self.assertEqual(len(request_payload["targets"]), 1)
        self.assertEqual(len(request_payload["targets"][0]["candidates"]), 1)
        self.assertNotIn("pptx_bytes", request_payload)


if __name__ == "__main__":
    unittest.main()
