from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
import json
import unittest

from ppt_automator.ai_mapper import AiSourceMatchSuggestion
from ppt_automator.ai_transform import AiTransformDiagnostic
from ppt_automator.ppt_discovery import PptTarget
from ppt_automator.table_normalizer import TransformPlan
from ppt_automator.xlsx_parser import ParsedXlsxTable
from web import main as web_main
from worker.processor import AnalysisResult


def _target(target_id: str) -> PptTarget:
    return PptTarget(
        slide_index=0,
        slide_number=1,
        slide_path="ppt/slides/slide1.xml",
        shape_name=target_id,
        shape_id=target_id,
        object_type="chart",
        left_in=0,
        top_in=0,
        width_in=1,
        height_in=1,
        expected_orientation="series_rows_categories_columns",
        expected_categories=["Total"],
        expected_series=["Serie"],
        expected_values=[[1]],
    )


def _source(file_name: str) -> ParsedXlsxTable:
    return ParsedXlsxTable(
        source_id=Path(file_name).stem,
        file_name=file_name,
        sheet_name="Sheet1",
        orientation="series_rows_categories_columns",
        categories=["Total"],
        series=["Serie"],
        values=[[1]],
        preview_rows=[["", "Total"], ["Serie", 1]],
    )


def _plan(target_id: str) -> TransformPlan:
    target = _target(target_id)
    source = _source(f"{target_id}.xlsx")
    return TransformPlan(
        target=target,
        datasource=source,
        action="align",
        orientation_xlsx=source.orientation,
        orientation_ppt=target.expected_orientation,
        categories=["Total"],
        series=["Serie"],
        values=[[1]],
        confidence=1,
        reason="match",
    )


def _analysis(plans: list[TransformPlan]) -> AnalysisResult:
    targets = [plan.target for plan in plans]
    sources = [plan.datasource for plan in plans]
    return AnalysisResult(
        plans=plans,
        preview=[],
        targets=targets,
        sources=sources,
        target_count=len(targets),
        source_count=len(sources),
        warnings=[],
    )


class AiScopeTests(unittest.TestCase):
    def test_manual_override_diagnostics_only_send_changed_target(self) -> None:
        analysis = _analysis([_plan("111"), _plan("222"), _plan("333")])
        calls: list[list[str]] = []

        def fake_diagnostics(plans, root=None):
            calls.append([plan.target_id for plan in plans])
            return [
                AiTransformDiagnostic(
                    target=plans[0].target_id,
                    status="ok",
                    confidence=0.94,
                    action="align",
                    reason="ok",
                    row_mapping={},
                    column_mapping={},
                    recommended_edit_data={"orientation": "", "headers": [], "rows": []},
                )
            ]

        with TemporaryDirectory() as tmp:
            job_dir = Path(tmp)
            with patch.object(web_main, "ai_configured", return_value=True):
                with patch.object(web_main, "suggest_transform_diagnostics", side_effect=fake_diagnostics):
                    payload, status = web_main._ai_diagnostics_for_job(
                        job_dir,
                        analysis,
                        allow_ai=True,
                        target_ids={"222"},
                    )

            self.assertEqual(calls, [["222"]])
            self.assertEqual(set(payload), {"222"})
            self.assertEqual(status["state"], "ok")

            cache = json.loads((job_dir / "ai_diagnostics.json").read_text(encoding="utf-8"))
            self.assertEqual(set(cache), {"222"})
            log = json.loads((job_dir / "logs" / "ai_usage.jsonl").read_text(encoding="utf-8").splitlines()[-1])
            self.assertEqual(log["operation"], "transform_diagnostics")
            self.assertEqual(log["target_count"], 1)
            self.assertEqual(log["payload_summary"]["targets"][0]["target"], "222")

    def test_saved_preview_mode_does_not_call_ai_for_missing_diagnostics(self) -> None:
        analysis = _analysis([_plan("111"), _plan("222")])

        with TemporaryDirectory() as tmp:
            with patch.object(web_main, "ai_configured", return_value=True):
                with patch.object(web_main, "suggest_transform_diagnostics") as fake_diagnostics:
                    payload, status = web_main._ai_diagnostics_for_job(
                        Path(tmp),
                        analysis,
                        allow_ai=False,
                    )

            fake_diagnostics.assert_not_called()
            self.assertEqual(payload, {})
            self.assertEqual(status["state"], "warn")
            self.assertIn("apenas dados salvos", status["message"])

    def test_source_match_cache_update_accepts_ai_suggestions(self) -> None:
        target = _target("444")
        source = _source("444.xlsx")
        analysis = AnalysisResult(
            plans=[],
            preview=[],
            targets=[target],
            sources=[source],
            target_count=1,
            source_count=1,
            warnings=[],
        )

        def fake_matches(targets, sources, existing_plan_ids=None, root=None):
            return [AiSourceMatchSuggestion("444", "444.xlsx", 0.88, "id compativel")]

        with TemporaryDirectory() as tmp:
            with patch.object(web_main, "ai_configured", return_value=True):
                with patch.object(web_main, "suggest_source_matches_with_ai", side_effect=fake_matches):
                    payload, status = web_main._ai_source_matches_for_job(Path(tmp), analysis, allow_ai=True)

            self.assertEqual(status["state"], "ok")
            self.assertEqual(payload["444"]["datasource"], "444.xlsx")
            self.assertEqual(payload["444"]["status"], "matched")


if __name__ == "__main__":
    unittest.main()
