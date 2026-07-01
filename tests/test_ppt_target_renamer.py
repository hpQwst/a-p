from __future__ import annotations

import unittest
import xml.etree.ElementTree as ET

from ppt_automator.ppt_discovery import NS
from ppt_automator.ppt_discovery import PptTarget
from ppt_automator.ppt_target_renamer import rename_targets_in_slide_xml
from ppt_automator.target_labeler import assign_slide_target_ids


SLIDE_XML = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <p:cSld>
    <p:spTree>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="8" name="7792738590"/></p:nvGraphicFramePr>
        <a:graphic><a:graphicData><c:chart/></a:graphicData></a:graphic>
      </p:graphicFrame>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="9" name="Tabela 1"/></p:nvGraphicFramePr>
        <a:graphic><a:graphicData><a:tbl/></a:graphicData></a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
</p:sld>
'''


def _target(shape_name: str, shape_id: str, object_type: str = "chart") -> PptTarget:
    return PptTarget(
        slide_index=0,
        slide_number=1,
        slide_path="ppt/slides/slide1.xml",
        shape_name=shape_name,
        shape_id=shape_id,
        object_type=object_type,
        left_in=0,
        top_in=0,
        width_in=1,
        height_in=1,
    )


class PptTargetRenamerTests(unittest.TestCase):
    def test_renames_chart_and_table_shapes_to_internal_ids(self) -> None:
        chart, table = assign_slide_target_ids(
            [
                _target("7792738590", "8", object_type="chart"),
                _target("Tabela 1", "9", object_type="table"),
            ]
        )

        updated = rename_targets_in_slide_xml(SLIDE_XML, [chart, table])
        root = ET.fromstring(updated)
        names = [
            cnv.attrib.get("name")
            for cnv in root.findall(".//p:nvGraphicFramePr/p:cNvPr", NS)
        ]

        self.assertEqual(names, ["S001_T001_CHART", "S001_T002_TABLE"])


if __name__ == "__main__":
    unittest.main()
