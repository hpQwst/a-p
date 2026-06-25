from __future__ import annotations

from io import BytesIO
import unittest

import openpyxl

from ppt_automator.xlsx_parser import parse_xlsx_table


class XlsxRangeParserTests(unittest.TestCase):
    def test_parse_xlsx_table_uses_manual_range_only(self) -> None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Dados"
        ws["A1"] = "ruido fora do range"
        ws["D5"] = ""
        ws["E5"] = "Jan/26"
        ws["F5"] = "Fev/26"
        ws["D6"] = "Total"
        ws["E6"] = 10
        ws["F6"] = 20

        data = BytesIO()
        wb.save(data)

        parsed = parse_xlsx_table(
            data.getvalue(),
            file_name="manual.xlsx",
            formula_mode="auto",
            cell_range="Dados!D5:F6",
        )

        self.assertEqual(parsed.sheet_name, "Dados")
        self.assertEqual(parsed.used_range, (5, 4, 6, 6))
        self.assertEqual(parsed.orientation, "single_series_row_categories_columns")
        self.assertEqual(parsed.categories, ["Jan/26", "Fev/26"])
        self.assertEqual(parsed.series, ["Total"])
        self.assertEqual(parsed.values, [[10, 20]])


if __name__ == "__main__":
    unittest.main()
