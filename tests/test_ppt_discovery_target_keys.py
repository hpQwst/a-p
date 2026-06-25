from __future__ import annotations

import unittest

from ppt_automator.ppt_discovery import PptTarget, _with_target_keys


def _target(shape_name: str, shape_id: str, slide_number: int = 1, object_type: str = "chart") -> PptTarget:
    return PptTarget(
        slide_index=slide_number - 1,
        slide_number=slide_number,
        slide_path=f"ppt/slides/slide{slide_number}.xml",
        shape_name=shape_name,
        shape_id=shape_id,
        object_type=object_type,
        left_in=0,
        top_in=0,
        width_in=1,
        height_in=1,
    )


class PptDiscoveryTargetKeyTests(unittest.TestCase):
    def test_unique_numeric_shape_name_stays_as_target_id(self) -> None:
        target = _with_target_keys([_target("7792738590", "8")])[0]

        self.assertEqual(target.target_id, "7792738590")

    def test_non_numeric_chart_gets_stable_slide_key(self) -> None:
        target = _with_target_keys([_target("Grafico 3", "5", slide_number=6)])[0]

        self.assertEqual(target.shape_name, "Grafico 3")
        self.assertEqual(target.target_id, "slide006_chart_5")

    def test_duplicate_numeric_shape_name_gets_stable_slide_key(self) -> None:
        first, second = _with_target_keys(
            [
                _target("123456", "2", slide_number=1),
                _target("123456", "7", slide_number=2),
            ]
        )

        self.assertEqual(first.target_id, "slide001_chart_2")
        self.assertEqual(second.target_id, "slide002_chart_7")


if __name__ == "__main__":
    unittest.main()
