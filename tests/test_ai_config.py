from __future__ import annotations

from pathlib import Path
from unittest.mock import patch
import tempfile
import unittest

from ppt_automator.ai import _model_for_operation, ai_configured, load_local_env, reasoning_for_operation


class AiConfigTests(unittest.TestCase):
    def test_local_env_overrides_blank_openai_key(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / ".env").write_text("OPENAI_API_KEY=test-key\nOPENAI_MODEL=test-model\n", encoding="utf-8")
            with patch.dict("os.environ", {"OPENAI_API_KEY": ""}, clear=False):
                load_local_env(root)
                self.assertTrue(ai_configured(root))

    def test_operation_defaults_route_simple_and_complex_work(self) -> None:
        with patch.dict(
            "os.environ",
            {
                "OPENAI_MODEL": "",
                "OPENAI_MODEL_SOURCE_MATCH": "",
                "OPENAI_MODEL_SLIDE_MATRIX_BUILDER": "",
                "OPENAI_REASONING_EFFORT": "",
                "OPENAI_REASONING_EFFORT_SOURCE_MATCH": "",
                "OPENAI_REASONING_EFFORT_SLIDE_MATRIX_BUILDER": "",
            },
            clear=False,
        ):
            self.assertEqual(_model_for_operation("source_match"), "gpt-5.6-luna")
            self.assertEqual(_model_for_operation("slide_matrix_builder"), "gpt-5.6-terra")
            self.assertEqual(reasoning_for_operation("source_match"), {"effort": "none"})
            self.assertEqual(reasoning_for_operation("slide_matrix_builder"), {"effort": "low"})


if __name__ == "__main__":
    unittest.main()
