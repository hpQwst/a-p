from __future__ import annotations

import unittest

from ppt_automator.slide_datasources import parse_slide_number_from_path


class SlideDatasourceTests(unittest.TestCase):
    def test_parse_slide_folder_variants(self) -> None:
        cases = {
            "slide_006/teste.xlsx": 6,
            "slide6/teste.xlsx": 6,
            "slide_6/teste.xlsx": 6,
            "s6/teste.xlsx": 6,
            "006/teste.xlsx": 6,
            "Slide 6/teste.xlsx": 6,
            "xxxxx_slide6.xlsx": 6,
            "xxxxx_s5.xlsx": 5,
            "xxxxx_s4.xlsx": 4,
            "xxxyxx_s4.xlsx": 4,
            "teste.xlsx": None,
        }

        for path, expected in cases.items():
            with self.subTest(path=path):
                self.assertEqual(parse_slide_number_from_path(path), expected)


if __name__ == "__main__":
    unittest.main()
