from __future__ import annotations

import unittest

from ppt_automator.ppt_discovery import PptTarget
from ppt_automator.preview_model import build_preview
from ppt_automator.table_normalizer import TransformPlan, normalize_to_target
from ppt_automator.xlsx_parser import ParsedXlsxTable


class PreviewModelTests(unittest.TestCase):
    def test_manual_table_plan_shows_after_matrix(self) -> None:
        target = _target("table")
        source = ParsedXlsxTable(
            source_id="",
            file_name="upload_manual/S001_T001_TABLE_y.xlsx",
            sheet_name="Planilha1",
            orientation="single_series_row_categories_columns",
            categories=["2025", "2026"],
            series=["a"],
            values=[[50, 20]],
        )

        preview = build_preview([normalize_to_target(target, source)])[0]

        self.assertEqual(preview.headers, ["2025", "2026"])
        self.assertEqual(preview.rows, [[50, 20]])

    def test_mixed_chart_preview_formats_each_series_separately(self) -> None:
        target = _target("chart")
        source = ParsedXlsxTable(
            source_id="dados",
            file_name="dados.xlsx",
            sheet_name="Sheet1",
            orientation="categories_rows_series_columns",
            categories=["Nov/25"],
            series=["Detrator", "Neutro", "Promotor", "NPS"],
            values=[[0.200899, 0.151455, 0.647646, 44.6747]],
            series_number_formats=["0%", "0%", "0%", "#,##0.0"],
        )
        plan = TransformPlan(
            target=target,
            datasource=source,
            action="align",
            orientation_xlsx=source.orientation,
            orientation_ppt=source.orientation,
            categories=source.categories,
            series=source.series,
            values=source.values,
            confidence=1.0,
            reason="teste",
        )

        preview = build_preview([plan])[0]

        self.assertEqual(
            preview.rows,
            [["Nov/25", "20%", "15%", "65%", "44,7"]],
        )


def _target(object_type: str) -> PptTarget:
    return PptTarget(
        slide_index=0,
        slide_number=1,
        slide_path="ppt/slides/slide1.xml",
        shape_name="Objeto 1",
        shape_id="1",
        object_type=object_type,
        left_in=0,
        top_in=0,
        width_in=1,
        height_in=1,
        expected_orientation="categories_rows_series_columns",
        expected_categories=["Nov/25"],
        expected_series=["Detrator", "Neutro", "Promotor", "NPS"],
        series_value_formats=["0%", "0%", "0%", "#,##0.0"],
        table_cells=[["2025", "2026"], ["", ""]],
    )


if __name__ == "__main__":
    unittest.main()
