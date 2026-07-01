from __future__ import annotations

from pathlib import Path
import os
import unittest

from ppt_automator import analyze_update_package


HUGO_DIR = Path(os.getenv("AUTO_PPT_HUGO_TEST_DIR", r"C:\Users\HugoRocha\Documents\automatizador-ppt-arquivos\hugo"))
PPT = HUGO_DIR / "Natura_2Q26_RelacionalCB_modelo_mapeado.pptx"
DATASOURCES = HUGO_DIR / "datasources.zip"


@unittest.skipUnless(PPT.exists() and DATASOURCES.exists(), "Arquivos Hugo de regressao nao encontrados.")
class HugoMatchingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.targets, cls.sources, cls.plans = analyze_update_package(PPT, DATASOURCES)

    def test_letter_named_xlsx_files_match_by_edit_data_contract(self) -> None:
        plan_by_shape = {plan.target.shape_name: plan for plan in self.plans}
        expected = {
            "7792738590": "j.xlsx",
            "5977261166": "g.xlsx",
            "3958478347": "f.xlsx",
            "1842587759": "e.xlsx",
            "9362596625": "k.xlsx",
            "7472034903": "i.xlsx",
            "1823080929": "d.xlsx",
            "1130655160": "b.xlsx",
            "9607212133": "l.xlsx",
            "6760545480": "h.xlsx",
        }
        for target_id, filename in expected.items():
            with self.subTest(target_id=target_id):
                self.assertIn(target_id, plan_by_shape)
                self.assertEqual(plan_by_shape[target_id].datasource.file_name, filename)
                self.assertGreaterEqual(plan_by_shape[target_id].confidence, 0.45)


if __name__ == "__main__":
    unittest.main()
