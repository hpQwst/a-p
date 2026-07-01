from __future__ import annotations

import unittest

from ppt_automator.edit_data_validator import validate_typed_edit_data
from ppt_automator.typed_matrix import normalize_typed_cell, normalize_typed_edit_data, numeric_value


class TypedMatrixTests(unittest.TestCase):
    def test_period_and_code_labels_are_forced_text(self) -> None:
        for value in ["Nov/25", "1Q26", "001", "10-15"]:
            cell = normalize_typed_cell(value)
            self.assertEqual(cell["type"], "text")
            self.assertTrue(cell["force_text"])

    def test_decimal_number_is_preserved_as_numeric_raw(self) -> None:
        cell = normalize_typed_cell("45.8419")
        self.assertEqual(cell["type"], "number")
        self.assertEqual(cell["source_raw"], "45.8419")
        self.assertEqual(str(numeric_value(cell)), "45.8419")

    def test_validation_rejects_ragged_rows_and_bad_number(self) -> None:
        ragged = {
            "headers": ["", "TIM"],
            "rows": [["Nov/25", "45.8419"], ["Dez/25"]],
        }
        self.assertTrue(validate_typed_edit_data(ragged))

        bad_number = {
            "headers": [{"value": "", "type": "text", "force_text": True, "source_raw": ""}, {"value": "TIM", "type": "text", "force_text": True, "source_raw": ""}],
            "rows": [[{"value": "Nov/25", "type": "text", "force_text": True, "source_raw": ""}, {"value": "abc", "type": "number", "force_text": False, "source_raw": "abc"}]],
        }
        self.assertTrue(validate_typed_edit_data(bad_number))

    def test_accepts_valid_typed_matrix(self) -> None:
        data = normalize_typed_edit_data(
            {
                "headers": ["", "TIM"],
                "rows": [["Nov/25", {"value": "45.8419", "type": "number", "source_raw": "45.8419"}]],
            }
        )

        self.assertEqual(validate_typed_edit_data(data), [])


if __name__ == "__main__":
    unittest.main()
