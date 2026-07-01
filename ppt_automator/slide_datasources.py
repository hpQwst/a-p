from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import PurePosixPath
from typing import BinaryIO
from zipfile import ZipFile
import re

from .ppt_discovery import read_bytes


InputFile = str | bytes | bytearray | BinaryIO


@dataclass(frozen=True)
class SlideDatasourceEntry:
    zip_path: str
    file_name: str
    slide_number: int | None
    is_general: bool


def parse_slide_number_from_path(path: str) -> int | None:
    parts = [part for part in PurePosixPath(path.replace("\\", "/")).parts if part and part != "."]
    if len(parts) > 1:
        folder = parts[0].strip()
        folder_slide = _parse_slide_token(folder)
        if folder_slide is not None:
            return folder_slide
    filename = PurePosixPath(parts[-1]).stem if parts else ""
    return _parse_slide_suffix(filename)


def _parse_slide_token(value: str) -> int | None:
    normalized = re.sub(r"[\s_-]+", "", value, flags=re.IGNORECASE).lower()
    patterns = [
        r"slide0*(\d+)",
        r"s0*(\d+)",
        r"0*(\d+)",
    ]
    for pattern in patterns:
        match = re.fullmatch(pattern, normalized)
        if match:
            number = int(match.group(1))
            return number if number > 0 else None
    match = re.fullmatch(r"slide\s*0*(\d+)", value, flags=re.IGNORECASE)
    if match:
        number = int(match.group(1))
        return number if number > 0 else None
    return None


def _parse_slide_suffix(filename_stem: str) -> int | None:
    match = re.search(r"(?:^|[\s_-])(?:slide|s)0*(\d+)$", filename_stem.strip(), flags=re.IGNORECASE)
    if match:
        number = int(match.group(1))
        return number if number > 0 else None
    return None


def collect_datasource_entries(datasources_zip: InputFile) -> list[SlideDatasourceEntry]:
    entries: list[SlideDatasourceEntry] = []
    with ZipFile(BytesIO(read_bytes(datasources_zip))) as zf:
        for name in zf.namelist():
            if not name.lower().endswith(".xlsx") or name.endswith("/"):
                continue
            slide_number = parse_slide_number_from_path(name)
            entries.append(
                SlideDatasourceEntry(
                    zip_path=name,
                    file_name=name,
                    slide_number=slide_number,
                    is_general=slide_number is None,
                )
            )
    return entries


def entries_for_slide(entries: list[SlideDatasourceEntry], slide_number: int) -> tuple[list[SlideDatasourceEntry], list[str]]:
    slide_entries = [entry for entry in entries if entry.slide_number == slide_number]
    if slide_entries:
        return slide_entries, []
    general = [entry for entry in entries if entry.is_general]
    if general:
        return general, [f"Slide {slide_number} nao tem pasta propria no ZIP; usando XLSX gerais como fallback."]
    return [], [f"Slide {slide_number} nao tem XLSX em pasta propria nem XLSX geral no ZIP."]


def read_entry_bytes(datasources_zip: InputFile, entry: SlideDatasourceEntry) -> bytes:
    with ZipFile(BytesIO(read_bytes(datasources_zip))) as zf:
        return zf.read(entry.zip_path)
