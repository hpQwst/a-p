from __future__ import annotations

import unittest

from ppt_automator.xlsx_parser import _graph_id


class XlsxGraphIdTests(unittest.TestCase):
    def test_slide_suffix_is_not_treated_as_graph_id(self) -> None:
        self.assertEqual(_graph_id("xxxxx_slide6"), "")
        self.assertEqual(_graph_id("xxxxx_s5"), "")
        self.assertEqual(_graph_id("7792738590_slide6"), "7792738590")
        self.assertEqual(_graph_id("7792738590_s5"), "7792738590")


if __name__ == "__main__":
    unittest.main()
