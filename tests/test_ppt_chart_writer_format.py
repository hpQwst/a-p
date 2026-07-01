from __future__ import annotations

import unittest

from ppt_automator.ppt_chart_writer import _chart_value_text


class PptChartWriterFormatTests(unittest.TestCase):
    def test_decimal_visual_format_does_not_scale_value_as_percentage(self) -> None:
        self.assertEqual(_chart_value_text(15.990453460620525, numeric=True, number_format="0.0"), "15.9904534606")

    def test_percent_visual_format_scales_percent_values_for_excel_chart_cache(self) -> None:
        self.assertEqual(_chart_value_text(15.990453460620525, numeric=True, number_format="0.0%"), "0.159904534606")


if __name__ == "__main__":
    unittest.main()
