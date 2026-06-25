from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
import os
import unittest

from ppt_automator.project_store import (
    create_project,
    ensure_store,
    list_mapping_templates,
    load_mapping_template,
    save_mapping_template,
)
from ppt_automator.ppt_discovery import PptTarget
from ppt_automator.table_normalizer import TransformPlan
from ppt_automator.xlsx_parser import ParsedXlsxTable
from worker.processor import AnalysisResult, apply_saved_source_matches_to_analysis


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


def _source(file_name: str, value: int) -> ParsedXlsxTable:
    return ParsedXlsxTable(
        source_id=Path(file_name).stem,
        file_name=file_name,
        sheet_name="Sheet1",
        orientation="series_rows_categories_columns",
        categories=["Total"],
        series=["Serie"],
        values=[[value]],
        preview_rows=[["", "Total"], ["Serie", value]],
    )


class MappingTemplateTests(unittest.TestCase):
    def test_mapping_templates_are_isolated_by_squad(self) -> None:
        with TemporaryDirectory() as tmp:
            env = {
                **os.environ,
                "AUTO_PPT_STORAGE_BACKEND": "local",
                "AUTO_PPT_DATA_ROOT": tmp,
            }
            with patch.dict(os.environ, env, clear=True):
                ensure_store()
                squad1_project = create_project("squad1", "Projeto Squad 1")
                squad2_project = create_project("squad2", "Projeto Squad 2")

                save_mapping_template(
                    squad2_project,
                    "Modelo Squad 2",
                    {"111": {"datasource": "a.xlsx"}},
                )

                self.assertEqual(list_mapping_templates("squad1"), [])
                self.assertEqual([item.name for item in list_mapping_templates("squad2")], ["Modelo Squad 2"])
                self.assertIsNone(load_mapping_template(squad1_project.squad, "modelo-squad-2"))

    def test_saved_mapping_overrides_plan_by_datasource_basename(self) -> None:
        target = _target("111")
        wrong_source = _source("wrong.xlsx", 1)
        right_source = _source("datasources/right.xlsx", 2)
        wrong_plan = TransformPlan(
            target=target,
            datasource=wrong_source,
            action="align",
            orientation_xlsx=wrong_source.orientation,
            orientation_ppt=target.expected_orientation,
            categories=["Total"],
            series=["Serie"],
            values=[[1]],
            confidence=0.5,
            reason="match automatico",
        )
        analysis = AnalysisResult(
            plans=[wrong_plan],
            preview=[],
            targets=[target],
            sources=[wrong_source, right_source],
            target_count=1,
            source_count=2,
            warnings=[],
        )

        updated = apply_saved_source_matches_to_analysis(
            analysis,
            {"111": {"datasource": "right.xlsx", "reason": "Mapeamento salvo"}},
        )

        self.assertEqual(updated.plans[0].datasource.file_name, "datasources/right.xlsx")
        self.assertEqual(updated.plans[0].values, [[2]])
        self.assertIn("Mapeamento salvo", updated.plans[0].reason)


if __name__ == "__main__":
    unittest.main()
