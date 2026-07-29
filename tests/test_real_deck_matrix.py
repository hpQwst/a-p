from __future__ import annotations

from pathlib import Path
import os
import unittest

from ppt_automator import analyze_update_package, generate_updated_pptx
from ppt_automator.ppt_chart_writer import resolved_series_number_formats
from ppt_automator.regression_validation import validate_generated_pptx


FIXTURE_ROOT = Path(
    os.getenv(
        "AUTO_PPT_REAL_FIXTURE_ROOT",
        r"C:\Users\HugoRocha\Documents\automatizador-ppt-arquivos",
    )
)
ANDRE_ENXUTO_PPT = (
    FIXTURE_ROOT
    / "andre"
    / "Natura_2Q26_RelacionalCB_modelo_mapeado-enxuto.pptx"
)
ANDRE_DATASOURCES = FIXTURE_ROOT / "andre" / "datasources.zip"
MB2_PPT = FIXTURE_ROOT / "mb2" / "C Experiência 1Q26_TRIMESTRAL_v01.pptx"
MB2_DATASOURCES = FIXTURE_ROOT / "mb2" / "datasources.zip"


class RealDeckMatrixTests(unittest.TestCase):
    @unittest.skipUnless(
        ANDRE_ENXUTO_PPT.exists() and ANDRE_DATASOURCES.exists(),
        "Deck Andre enxuto de regressao nao encontrado.",
    )
    def test_andre_enxuto_generates_without_collateral_package_changes(self) -> None:
        targets, sources, plans = analyze_update_package(
            ANDRE_ENXUTO_PPT,
            ANDRE_DATASOURCES,
        )
        self.assertGreaterEqual(len(targets), 4)
        self.assertEqual(len(sources), 12)
        self.assertEqual(len(plans), 4)

        generated = generate_updated_pptx(
            ANDRE_ENXUTO_PPT,
            plans,
            targets=targets,
        )
        report = validate_generated_pptx(ANDRE_ENXUTO_PPT, generated, plans)
        self.assertEqual(report.charts_checked + report.tables_checked, len(plans))

    @unittest.skipUnless(
        (FIXTURE_ROOT / "hugo" / "Natura_2Q26_RelacionalCB_modelo_mapeado.pptx").exists()
        and (FIXTURE_ROOT / "hugo" / "datasources.zip").exists(),
        "Deck Hugo de regressao nao encontrado.",
    )
    def test_hugo_percentage_sources_keep_percentage_labels(self) -> None:
        pptx = FIXTURE_ROOT / "hugo" / "Natura_2Q26_RelacionalCB_modelo_mapeado.pptx"
        datasources = FIXTURE_ROOT / "hugo" / "datasources.zip"
        _targets, _sources, plans = analyze_update_package(pptx, datasources)
        percentage_plans = [
            plan
            for plan in plans
            if plan.target.slide_number == 3
            and any("%" in fmt for fmt in plan.datasource.series_number_formats)
        ]
        self.assertEqual(len(percentage_plans), 4)
        for plan in percentage_plans:
            with self.subTest(target=plan.target.shape_name):
                formats = resolved_series_number_formats(plan.target, plan)
                self.assertTrue(formats)
                self.assertTrue(all("%" in fmt for fmt in formats))

    @unittest.skipUnless(
        MB2_PPT.exists() and MB2_DATASOURCES.exists(),
        "Deck MB2 grande de regressao nao encontrado.",
    )
    def test_mb2_formula_deck_analyzes_and_generates(self) -> None:
        targets, sources, plans = analyze_update_package(MB2_PPT, MB2_DATASOURCES)
        self.assertGreaterEqual(len(targets), 400)
        self.assertEqual(len(sources), 2)
        self.assertGreaterEqual(len(plans), 6)
        self.assertTrue({6, 7, 8, 9, 45, 46}.issubset({plan.target.slide_number for plan in plans}))

        generated = generate_updated_pptx(MB2_PPT, plans, targets=targets)
        report = validate_generated_pptx(MB2_PPT, generated, plans)
        self.assertEqual(report.charts_checked + report.tables_checked, len(plans))


if __name__ == "__main__":
    unittest.main()
