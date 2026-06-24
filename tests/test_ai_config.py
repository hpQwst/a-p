from __future__ import annotations

from pathlib import Path
from unittest.mock import patch
import tempfile
import unittest

from ppt_automator.ai import ai_configured, load_local_env


class AiConfigTests(unittest.TestCase):
    def test_local_env_overrides_blank_openai_key(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / ".env").write_text("OPENAI_API_KEY=test-key\nOPENAI_MODEL=test-model\n", encoding="utf-8")
            with patch.dict("os.environ", {"OPENAI_API_KEY": ""}, clear=False):
                load_local_env(root)
                self.assertTrue(ai_configured(root))


if __name__ == "__main__":
    unittest.main()
