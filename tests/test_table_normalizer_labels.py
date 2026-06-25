from __future__ import annotations

import unittest

from ppt_automator.ppt_discovery import PptTarget
from ppt_automator.table_normalizer import normalize_to_target
from ppt_automator.xlsx_parser import ParsedXlsxTable


class TableNormalizerLabelTests(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
