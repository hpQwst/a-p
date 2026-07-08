from __future__ import annotations

import unittest

from ppt_automator.ppt_discovery import PptTarget
from ppt_automator.xlsx_parser import ParsedXlsxTable
from ppt_automator.table_normalizer import build_transform_plans, _hungarian


def _target(shape: str, categories: list[str]) -> PptTarget:
    return PptTarget(
        slide_index=0,
        slide_number=3,
        slide_path="ppt/slides/slide3.xml",
        shape_name=shape,
        shape_id=shape,
        object_type="chart",
        left_in=0.0,
        top_in=0.0,
        width_in=5.0,
        height_in=3.0,
        expected_orientation="categories_rows_series_columns",
        expected_series=["Total"],
        expected_categories=categories,
    )


def _source(name: str, categories: list[str]) -> ParsedXlsxTable:
    return ParsedXlsxTable(
        source_id="",
        file_name=name,
        sheet_name="Plan1",
        orientation="categories_rows_series_columns",
        categories=categories,
        series=["Total"],
        values=[[float(i)] for i in range(len(categories))],
    )


class SlideAssignmentTests(unittest.TestCase):
    def test_hungarian_is_one_to_one_and_optimal(self) -> None:
        # custo: linha 0 prefere col 0; linha 1 prefere col 0 tambem, mas o otimo
        # global obriga uma a ceder -> atribuicao 1:1 sem colisao.
        cost = [[0.1, 0.9], [0.2, 0.8]]
        assignment = _hungarian(cost)
        self.assertEqual(sorted(assignment), [0, 1])
        # otimo: (0->1, 1->0) custo 0.9+0.2=1.1  vs (0->0,1->1)=0.1+0.8=0.9 -> escolhe 0.9
        self.assertEqual(assignment, [0, 1])

    def test_two_targets_never_share_the_same_datasource(self) -> None:
        targets = [
            _target("chartA", ["Norte", "Sul", "Leste"]),
            _target("chartB", ["Cristal", "Bronze", "Prata"]),
        ]
        sources = [
            _source("datasources/tabA_slide3.xlsx", ["Norte", "Sul", "Leste"]),
            _source("datasources/tabB_slide3.xlsx", ["Cristal", "Bronze", "Prata"]),
        ]
        plans = build_transform_plans(targets, sources)
        self.assertEqual(len(plans), 2)
        used = [plan.datasource.file_name for plan in plans]
        self.assertEqual(len(set(used)), 2)
        by_shape = {plan.target.shape_name: plan.datasource.file_name for plan in plans}
        self.assertEqual(by_shape["chartA"], "datasources/tabA_slide3.xlsx")
        self.assertEqual(by_shape["chartB"], "datasources/tabB_slide3.xlsx")

    def test_slide_hint_blocks_cross_slide_datasources(self) -> None:
        targets = [_target("chartA", ["Norte", "Sul", "Leste"])]
        sources = [
            _source("datasources/tabA_slide3.xlsx", ["Norte", "Sul", "Leste"]),
            _source("datasources/tabX_slide9.xlsx", ["Norte", "Sul", "Leste"]),
        ]
        plans = build_transform_plans(targets, sources)
        self.assertEqual(len(plans), 1)
        self.assertEqual(plans[0].datasource.file_name, "datasources/tabA_slide3.xlsx")


if __name__ == "__main__":
    unittest.main()
