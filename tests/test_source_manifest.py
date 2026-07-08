from __future__ import annotations

import unittest

from ppt_automator.source_manifest import xlsx_source_manifest
from ppt_automator.xlsx_parser import ParsedXlsxTable


class SourceManifestTests(unittest.TestCase):
    def test_manifest_names_structural_profile_and_recipe_hint(self) -> None:
        source = ParsedXlsxTable(
            source_id="tab123",
            file_name="datasources/tab123_slide3.xlsx",
            sheet_name="Sheet1",
            orientation="series_rows_categories_columns",
            categories=["Total", "Rede 1"],
            series=["Cristal", "Bronze"],
            values=[[1, 2], [3, 4]],
            preview_rows=[["Plano de crescimento"], ["", "Total", "Rede 1"], ["Cristal", 1, 2]],
            metadata={"table_title": "Plano de crescimento", "row_group_label": "Brasil"},
        )

        manifest = xlsx_source_manifest(source)
        profile = manifest["structural_profile"]

        self.assertEqual(profile["name"], "contextual_series_rows_categories_columns")
        self.assertTrue(profile["signature"])
        self.assertEqual(profile["recipe_hint"]["operation"], "keep")
        self.assertEqual(profile["recipe_hint"]["orientation_after"], "series_rows_categories_columns")


if __name__ == "__main__":
    unittest.main()
