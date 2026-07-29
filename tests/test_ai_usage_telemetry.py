from __future__ import annotations

import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from ppt_automator.ai_debug import (
    log_ai_request,
    log_ai_response,
    reset_ai_usage_log_path,
    set_ai_usage_log_path,
)


class _FakeResponse:
    model = "gpt-5.6-luna"

    def model_dump(self) -> dict:
        return {
            "model": self.model,
            "usage": {
                "input_tokens": 100,
                "input_tokens_details": {"cached_tokens": 20},
                "output_tokens": 10,
                "output_tokens_details": {"reasoning_tokens": 2},
                "total_tokens": 110,
            },
        }


class AiUsageTelemetryTests(unittest.TestCase):
    def test_compact_log_records_cost_without_prompt_or_output(self) -> None:
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "ai_usage.jsonl"
            token = set_ai_usage_log_path(path)
            try:
                log_ai_request(
                    "source_match",
                    {
                        "model": "gpt-5.6-luna",
                        "reasoning": {"effort": "none"},
                        "input": [{"role": "user", "content": "conteudo sensivel"}],
                    },
                )
                log_ai_response("source_match", "saida sensivel", 321, _FakeResponse())
            finally:
                reset_ai_usage_log_path(token)

            record = json.loads(path.read_text(encoding="utf-8"))

        self.assertEqual(record["model_returned"], "gpt-5.6-luna")
        self.assertEqual(record["reasoning_effort"], "none")
        self.assertEqual(record["total_tokens"], 110)
        self.assertEqual(record["cached_input_tokens"], 20)
        self.assertGreater(record["estimated_cost_usd"], 0)
        serialized = json.dumps(record, ensure_ascii=False)
        self.assertNotIn("conteudo sensivel", serialized)
        self.assertNotIn("saida sensivel", serialized)


if __name__ == "__main__":
    unittest.main()
