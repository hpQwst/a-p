from __future__ import annotations

from datetime import datetime
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

    def test_excel_date_period_labels_keep_month_year_text(self) -> None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Dados"
        ws.append(["Serie", datetime(2026, 11, 25), "Dez/25", datetime(2026, 1, 26)])
        ws.append(["NPS", 1, 2, 3])

        data = BytesIO()
        wb.save(data)
        wb.close()

        parsed = parse_xlsx_table(data.getvalue(), file_name="periods.xlsx")

        self.assertEqual(parsed.categories, ["Nov/25", "Dez/25", "Jan/26"])


if __name__ == "__main__":
    unittest.main()
