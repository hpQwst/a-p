from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
import os
import unittest

from worker import aws_runtime


class AwsRuntimeTests(unittest.TestCase):
    def test_fargate_backend_is_explicit(self) -> None:
        with patch.dict(os.environ, {"AUTO_PPT_EXECUTION_BACKEND": "fargate"}, clear=False):
            self.assertTrue(aws_runtime.uses_fargate_workers())
        with patch.dict(os.environ, {"AUTO_PPT_EXECUTION_BACKEND": "local"}, clear=False):
            self.assertFalse(aws_runtime.uses_fargate_workers())

    def test_rejects_path_escape_when_hydrating(self) -> None:
        with TemporaryDirectory() as tmp:
            with self.assertRaises(aws_runtime.AwsExecutionError):
                aws_runtime._safe_destination(Path(tmp), "../../outside")

    def test_job_prefix_is_namespaced(self) -> None:
        with patch.dict(os.environ, {"AUTO_PPT_S3_PREFIX": "company/auto-ppt"}, clear=False):
            self.assertEqual(aws_runtime.job_prefix("abc", "preview_processing.json"), "company/auto-ppt/jobs/abc/preview_processing.json")

    def test_client_token_changes_between_launches_but_is_stable_per_launch(self) -> None:
        token_a1 = aws_runtime._client_token("job123", "preview", "2026-07-10T12:00:00")
        token_a2 = aws_runtime._client_token("job123", "preview", "2026-07-10T12:00:00")
        token_b = aws_runtime._client_token("job123", "preview", "2026-07-10T12:05:00")
        self.assertEqual(token_a1, token_a2)
        self.assertNotEqual(token_a1, token_b)


if __name__ == "__main__":
    unittest.main()
