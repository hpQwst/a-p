from __future__ import annotations

import unittest

from ppt_automator.ppt_discovery import PptTarget
from ppt_automator.xlsx_parser import ParsedXlsxTable
from ppt_automator.table_normalizer import normalize_to_target


def _chart_target(orientation: str, series: list[str], categories: list[str]) -> PptTarget:
    return PptTarget(
        slide_index=0,
        slide_number=3,
        slide_path="ppt/slides/slide3.xml",
        shape_name="chart1",
        shape_id="1",
        object_type="chart",
        left_in=0.0,
        top_in=0.0,
        width_in=5.0,
        height_in=3.0,
        expected_orientation=orientation,
        expected_series=series,
        expected_categories=categories,
    )


def _source(orientation: str, categories: list[str], series: list[str], values: list[list[float]]) -> ParsedXlsxTable:
    return ParsedXlsxTable(
        source_id="",
        file_name="datasources/tab_slide3.xlsx",
        sheet_name="Plan1",
        orientation=orientation,
        categories=categories,
        series=series,
        values=values,
    )


class TransposeAlignmentTests(unittest.TestCase):
    def test_inverted_xlsx_is_transposed_to_match_edit_data_contract(self) -> None:
        # Contrato do Editar dados: series nas LINHAS, categorias nas COLUNAS.
        target = _chart_target(
            "series_rows_categories_columns",
            series=["Total", "Natura", "Avon"],
            categories=["Ativa", "Inativa", "Cessada"],
        )
        # XLSX veio INVERTIDO: categorias nas linhas, series nas colunas.
        source = _source(
            "categories_rows_series_columns",
            categories=["Ativa", "Inativa", "Cessada"],
            series=["Total", "Natura", "Avon"],
            values=[[1, 2, 3], [4, 5, 6], [7, 8, 9]],
        )
        plan = normalize_to_target(target, source, confidence=1.0)

        self.assertEqual(plan.action, "transpose")
        self.assertEqual(plan.series, ["Total", "Natura", "Avon"])
        self.assertEqual(plan.categories, ["Ativa", "Inativa", "Cessada"])
        # Matriz reordenada para a orientacao do Editar dados (transposta da origem).
        self.assertEqual(plan.values, [[1, 4, 7], [2, 5, 8], [3, 6, 9]])

    def test_aligned_xlsx_is_kept_without_transpose(self) -> None:
        target = _chart_target(
            "series_rows_categories_columns",
            series=["Total", "Natura", "Avon"],
            categories=["Ativa", "Inativa", "Cessada"],
        )
        # XLSX ja na mesma orientacao do contrato.
        source = _source(
            "series_rows_categories_columns",
            categories=["Ativa", "Inativa", "Cessada"],
            series=["Total", "Natura", "Avon"],
            values=[[1, 4, 7], [2, 5, 8], [3, 6, 9]],
        )
        plan = normalize_to_target(target, source, confidence=1.0)

        self.assertEqual(plan.action, "align")
        self.assertEqual(plan.values, [[1, 4, 7], [2, 5, 8], [3, 6, 9]])


if __name__ == "__main__":
    unittest.main()
