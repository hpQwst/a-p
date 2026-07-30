from __future__ import annotations

import unittest

from ppt_automator.ppt_discovery import PptTarget
from ppt_automator.table_normalizer import normalize_to_target
from ppt_automator.xlsx_parser import ParsedXlsxTable


class TableNormalizerLabelTests(unittest.TestCase):
    def test_chart_output_uses_xlsx_labels_instead_of_ppt_labels(self) -> None:
        target = PptTarget(
            slide_index=0,
            slide_number=1,
            slide_path="ppt/slides/slide1.xml",
            shape_name="7792738590",
            shape_id="1",
            object_type="chart",
            left_in=0,
            top_in=0,
            width_in=1,
            height_in=1,
            expected_orientation="categories_rows_series_columns",
            expected_categories=[
                "Não tem auto-consumo",
                "Até 25%",
                "Entre 26% e 50%",
                "Entre 51% e 75%",
                "Mais de 75%",
            ],
            expected_series=["Total", "Natura", "Avon"],
            expected_values=[[10.0, 10.0, 10.0] for _ in range(5)],
        )
        source = ParsedXlsxTable(
            source_id="",
            file_name="j.xlsx",
            sheet_name="Sheet1",
            orientation="categories_rows_series_columns",
            categories=[
                "Não compro produtos para uso próprio",
                "Até 25% das minhas compras são para uso próprio",
                "Entre 26% e 50%",
                "Entre 51% e 75%",
                "Mais de 75% das compras são para uso próprio",
            ],
            series=["Total", "Natura", "Avon"],
            values=[
                [9.59, 9.79, 8.81],
                [61.71, 62.62, 58.23],
                [17.47, 17.24, 18.32],
                [5.96, 5.82, 6.49],
                [5.25, 4.50, 8.12],
            ],
        )

        plan = normalize_to_target(target, source)

        self.assertEqual(plan.categories, source.categories)
        self.assertEqual(plan.series, source.series)
        self.assertEqual(plan.values[0], [9.59, 9.79, 8.81])
        self.assertNotIn("Não tem auto-consumo", plan.categories)

    def test_diamante_plus_does_not_match_respondentes(self) -> None:
        target = PptTarget(
            slide_index=0,
            slide_number=1,
            slide_path="ppt/slides/slide1.xml",
            shape_name="1130655160",
            shape_id="1",
            object_type="chart",
            left_in=0,
            top_in=0,
            width_in=1,
            height_in=1,
            expected_orientation="series_rows_categories_columns",
            expected_categories=["Total", "Rede 1", "Rede 2"],
            expected_series=["Cristal", "Bronze", "Prata", "Ouro", "Diamante", "Diamante +"],
            expected_values=[[10.0, 10.0, 10.0] for _ in range(6)],
        )
        source = ParsedXlsxTable(
            source_id="",
            file_name="e.xlsx",
            sheet_name="Sheet1",
            orientation="categories_rows_series_columns",
            categories=["Cristal", "Bronze", "Prata", "Ouro", "Diamante", "Diamante +", "Respondentes"],
            series=["TOTAL", "REDE 1", "REDE 2"],
            values=[
                [0.16, 0.14, 0.21],
                [0.30, 0.27, 0.44],
                [0.27, 0.28, 0.25],
                [0.16, 0.19, 0.06],
                [0.07, 0.09, 0.02],
                [0.008, 0.010, 0.0],
                [1584.0, 1251.0, 333.0],
            ],
        )

        plan = normalize_to_target(target, source)

        self.assertEqual(plan.series[-1], "Diamante +")
        self.assertEqual(plan.values[-1], [0.008, 0.010, 0.0])
        self.assertNotEqual(plan.values[-1], [1584.0, 1251.0, 333.0])

    def test_chart_preserves_xlsx_order_after_matching_ppt_labels(self) -> None:
        target = PptTarget(
            slide_index=0,
            slide_number=1,
            slide_path="ppt/slides/slide1.xml",
            shape_name="chart",
            shape_id="1",
            object_type="chart",
            left_in=0,
            top_in=0,
            width_in=1,
            height_in=1,
            expected_orientation="categories_rows_series_columns",
            expected_categories=["Primeiro", "Segundo", "Terceiro"],
            expected_series=["Total", "Natura"],
            expected_values=[[0, 0], [0, 0], [0, 0]],
        )
        source = ParsedXlsxTable(
            source_id="",
            file_name="reordenado.xlsx",
            sheet_name="Sheet1",
            orientation="categories_rows_series_columns",
            categories=["Terceiro completo", "Primeiro completo", "Segundo completo"],
            series=["Natura oficial", "Total oficial"],
            values=[[31, 30], [11, 10], [21, 20]],
        )

        plan = normalize_to_target(target, source)

        self.assertEqual(plan.categories, source.categories)
        self.assertEqual(plan.series, source.series)
        self.assertEqual(plan.values, source.values)

    def test_named_table_or_dynamic_range_adds_new_months_from_source(self) -> None:
        target = PptTarget(
            slide_index=0,
            slide_number=1,
            slide_path="ppt/slides/slide1.xml",
            shape_name="chart",
            shape_id="1",
            object_type="chart",
            left_in=0,
            top_in=0,
            width_in=1,
            height_in=1,
            expected_orientation="series_rows_categories_columns",
            expected_categories=["Jan/26", "Fev/26"],
            expected_series=["Detrator", "Promotor"],
            expected_values=[[0.20, 0.19], [0.65, 0.67]],
        )
        source = ParsedXlsxTable(
            source_id="",
            file_name="nps.xlsx",
            sheet_name="Dados",
            orientation="series_rows_categories_columns",
            categories=["Jan/26", "Fev/26", "Mar/26"],
            series=["Detrator", "Promotor"],
            values=[[0.20, 0.19, 0.18], [0.65, 0.67, 0.69]],
            metadata={"_allow_axis_growth": "1"},
        )

        plan = normalize_to_target(target, source)

        self.assertEqual(plan.categories, source.categories)
        self.assertEqual(plan.series, source.series)
        self.assertEqual(plan.values, source.values)

    def test_manual_orientation_forces_chart_transpose(self) -> None:
        target = PptTarget(
            slide_index=0,
            slide_number=1,
            slide_path="ppt/slides/slide1.xml",
            shape_name="chart",
            shape_id="1",
            object_type="chart",
            left_in=0,
            top_in=0,
            width_in=1,
            height_in=1,
            expected_orientation="categories_rows_series_columns",
            expected_categories=["Jan", "Fev"],
            expected_series=["Total", "Natura"],
            expected_values=[[1, 2], [3, 4]],
        )
        source = ParsedXlsxTable(
            source_id="",
            file_name="dados.xlsx",
            sheet_name="Dados",
            orientation="categories_rows_series_columns",
            categories=["Jan", "Fev"],
            series=["Total", "Natura"],
            values=[[1, 2], [3, 4]],
        )

        automatic = normalize_to_target(target, source)
        forced = normalize_to_target(target, source, axis_mode="transpor")

        self.assertEqual(automatic.action, "align")
        self.assertEqual(forced.action, "transpose")
        self.assertIn("forçou a transposicao", forced.reason)

    def test_manual_orientation_transposes_powerpoint_table_matrix(self) -> None:
        target = PptTarget(
            slide_index=0,
            slide_number=1,
            slide_path="ppt/slides/slide1.xml",
            shape_name="table",
            shape_id="2",
            object_type="table",
            left_in=0,
            top_in=0,
            width_in=1,
            height_in=1,
            table_cells=[["", "Jan", "Fev"], ["Total", 1, 2], ["Natura", 3, 4]],
        )
        source = ParsedXlsxTable(
            source_id="",
            file_name="dados.xlsx",
            sheet_name="Dados",
            orientation="series_rows_categories_columns",
            categories=["Jan", "Fev"],
            series=["Total", "Natura"],
            values=[[1, 2], [3, 4]],
        )

        normal = normalize_to_target(target, source)
        transposed = normalize_to_target(target, source, axis_mode="transpor")
        expected = [
            list(row)
            for row in zip(
                *[
                    values + [None] * (max(map(len, normal.table_matrix)) - len(values))
                    for values in normal.table_matrix
                ]
            )
        ]

        self.assertEqual(transposed.table_matrix, expected)
        self.assertIn("transpos a tabela", transposed.reason)


if __name__ == "__main__":
    unittest.main()
