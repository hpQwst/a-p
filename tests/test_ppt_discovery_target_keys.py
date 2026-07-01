from __future__ import annotations

import unittest

from ppt_automator.ppt_discovery import PptTarget
from ppt_automator.target_labeler import assign_slide_target_ids


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
    def test_unique_numeric_shape_name_gets_stable_target_id(self) -> None:
        target = assign_slide_target_ids([_target("7792738590", "8")])[0]

        self.assertEqual(target.target_id, "S001_T001_CHART")

    def test_non_numeric_chart_gets_stable_slide_key(self) -> None:
        target = assign_slide_target_ids([_target("Grafico 3", "5", slide_number=6)])[0]

        self.assertEqual(target.shape_name, "Grafico 3")
        self.assertEqual(target.target_id, "S006_T001_CHART")

    def test_duplicate_numeric_shape_name_gets_stable_slide_key(self) -> None:
        first, second = assign_slide_target_ids(
            [
                _target("123456", "2", slide_number=1),
                _target("123456", "7", slide_number=2),
            ]
        )

        self.assertEqual(first.target_id, "S001_T001_CHART")
        self.assertEqual(second.target_id, "S002_T001_CHART")

    def test_existing_internal_target_id_is_preserved(self) -> None:
        target = assign_slide_target_ids([_target("S006_T003_CHART", "8", slide_number=6)])[0]

        self.assertEqual(target.target_id, "S006_T003_CHART")


if __name__ == "__main__":
    unittest.main()
