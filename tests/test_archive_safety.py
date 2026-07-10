from __future__ import annotations

from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile
import unittest

from ppt_automator.archive_safety import (
    ArchiveLimits,
    UnsafeArchiveError,
    validate_datasource_zip_bytes,
    validate_pptx_bytes,
)


class ArchiveSafetyTests(unittest.TestCase):
    def setUp(self) -> None:
        self.limits = ArchiveLimits(
            max_members=20,
            max_uncompressed_bytes=10_000,
            max_member_bytes=8_000,
            max_compression_ratio=1_000,
        )

    def test_accepts_minimal_pptx_and_datasource_zip(self) -> None:
        pptx = _zip_bytes({"[Content_Types].xml": b"types", "ppt/presentation.xml": b"ppt"})
        xlsx = _zip_bytes({"[Content_Types].xml": b"types", "xl/workbook.xml": b"book"})
        datasources = _zip_bytes({"slide1/source.xlsx": xlsx})

        validate_pptx_bytes(pptx, self.limits)
        validate_datasource_zip_bytes(datasources, self.limits)

    def test_rejects_path_traversal(self) -> None:
        payload = _zip_bytes({"../source.xlsx": b"bad"})

        with self.assertRaisesRegex(UnsafeArchiveError, "caminho interno inseguro"):
            validate_datasource_zip_bytes(payload, self.limits)

    def test_rejects_oversized_member(self) -> None:
        payload = _zip_bytes(
            {"[Content_Types].xml": b"types", "ppt/presentation.xml": b"x" * 9_000}
        )

        with self.assertRaisesRegex(UnsafeArchiveError, "item maior"):
            validate_pptx_bytes(payload, self.limits)


def _zip_bytes(files: dict[str, bytes]) -> bytes:
    output = BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        for name, data in files.items():
            archive.writestr(name, data)
    return output.getvalue()


if __name__ == "__main__":
    unittest.main()
