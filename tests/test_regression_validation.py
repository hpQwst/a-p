from __future__ import annotations

from io import BytesIO
from pathlib import Path
import subprocess
import sys
from zipfile import ZIP_DEFLATED, ZipFile
import unittest

from ppt_automator.ppt_discovery import PptTarget
from ppt_automator.regression_validation import (
    PptxRegressionError,
    validate_generated_pptx,
)
from ppt_automator.table_normalizer import TransformPlan
from ppt_automator.xlsx_parser import ParsedXlsxTable


class RegressionValidationTests(unittest.TestCase):
    def test_structural_validator_import_does_not_require_pillow(self) -> None:
        repository_root = Path(__file__).resolve().parents[1]
        code = """
import builtins
real_import = builtins.__import__
def guarded_import(name, *args, **kwargs):
    if name == "PIL" or name.startswith("PIL."):
        raise ModuleNotFoundError("Pillow intentionally unavailable")
    return real_import(name, *args, **kwargs)
builtins.__import__ = guarded_import
import ppt_automator.regression_validation
"""
        completed = subprocess.run(
            [sys.executable, "-c", code],
            cwd=repository_root,
            capture_output=True,
            text=True,
            timeout=30,
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)

    def test_accepts_intact_noop_package(self) -> None:
        pptx = _minimal_pptx()

        report = validate_generated_pptx(pptx, pptx, [])

        self.assertEqual(report.slide_count, 1)
        self.assertGreater(report.xml_parts, 0)
        self.assertGreater(report.geometry_objects_checked, 0)

    def test_rejects_broken_internal_relationship(self) -> None:
        pptx = _minimal_pptx(
            extra={
                "ppt/slides/_rels/slide1.xml.rels": (
                    b'<?xml version="1.0" encoding="UTF-8"?>'
                    b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    b'<Relationship Id="rId1" Type="urn:test" Target="../media/missing.png"/>'
                    b"</Relationships>"
                )
            }
        )

        with self.assertRaisesRegex(PptxRegressionError, "Relacionamento quebrado"):
            validate_generated_pptx(pptx, pptx, [])

    def test_rejects_change_outside_target_parts(self) -> None:
        original = _minimal_pptx()
        generated = _replace_part(
            original,
            "ppt/presentation.xml",
            b'<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" changed="1"/>',
        )

        with self.assertRaisesRegex(PptxRegressionError, "fora dos targets"):
            validate_generated_pptx(original, generated, [])

    def test_rejects_truncated_powerpoint_table(self) -> None:
        pptx = _minimal_pptx()
        target = PptTarget(
            slide_index=0,
            slide_number=1,
            slide_path="ppt/slides/slide1.xml",
            shape_name="Tabela 1",
            shape_id="2",
            object_type="table",
            left_in=0,
            top_in=0,
            width_in=1,
            height_in=1,
            table_cells=[[""]],
        )
        source = ParsedXlsxTable(
            source_id="teste",
            file_name="teste.xlsx",
            sheet_name="Sheet1",
            orientation="series_rows_categories_columns",
            categories=["2025", "2026"],
            series=["Total"],
            values=[[10, 20]],
        )
        plan = TransformPlan(
            target=target,
            datasource=source,
            action="fill_table_cells",
            orientation_xlsx=source.orientation,
            orientation_ppt="table_cells",
            categories=source.categories,
            series=source.series,
            values=source.values,
            confidence=1.0,
            reason="teste",
            table_matrix=[["2025", "2026"], [10, 20]],
            table_header_rows=1,
        )

        with self.assertRaisesRegex(PptxRegressionError, "truncou linhas"):
            validate_generated_pptx(pptx, pptx, [plan])


def _minimal_pptx(extra: dict[str, bytes] | None = None) -> bytes:
    parts = {
        "[Content_Types].xml": (
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
        ),
        "ppt/presentation.xml": (
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
        ),
        "ppt/slides/slide1.xml": _slide_xml(),
        **(extra or {}),
    }
    output = BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return output.getvalue()


def _slide_xml() -> bytes:
    return b"""<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="2" name="Tabela 1"/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></p:xfrm>
        <a:graphic><a:graphicData><a:tbl>
          <a:tblGrid><a:gridCol w="1000"/></a:tblGrid>
          <a:tr h="1000">
            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t></a:t></a:r></a:p></a:txBody></a:tc>
          </a:tr>
        </a:tbl></a:graphicData></a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
</p:sld>"""


def _replace_part(pptx: bytes, name: str, data: bytes) -> bytes:
    output = BytesIO()
    with ZipFile(BytesIO(pptx)) as source, ZipFile(output, "w", ZIP_DEFLATED) as target:
        for info in source.infolist():
            target.writestr(info, data if info.filename == name else source.read(info.filename))
    return output.getvalue()


if __name__ == "__main__":
    unittest.main()
