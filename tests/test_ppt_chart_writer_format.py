from __future__ import annotations

import unittest

from ppt_automator.ppt_chart_writer import _chart_value_text


class PptChartWriterFormatTests(unittest.TestCase):
    def test_decimal_visual_format_does_not_scale_value_as_percentage(self) -> None:
        self.assertEqual(_chart_value_text(15.990453460620525, numeric=True, number_format="0.0"), "15.9904534606")

    def test_value_is_written_verbatim_even_with_percent_format(self) -> None:
        # Regra: o valor gravado e sempre verbatim (precisao total para o "Editar
        # dados"); um formatCode "%" so afeta a EXIBICAO no PowerPoint, nunca a
        # escala do numero armazenado. Nao dividimos mais por 100.
        self.assertEqual(_chart_value_text(15.990453460620525, numeric=True, number_format="0.0%"), "15.9904534606")

    def test_textual_percent_is_stripped_without_rescaling(self) -> None:
        self.assertEqual(_chart_value_text("30,8%", numeric=True, number_format="0.0%"), "30.8")


if __name__ == "__main__":
    unittest.main()
