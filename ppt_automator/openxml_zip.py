from __future__ import annotations

from io import BytesIO
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile
import binascii
import struct
import zlib


class ZipPartReplacementError(ValueError):
    pass


def replace_zip_parts_preserving_structure(zip_bytes: bytes, replacements: dict[str, bytes]) -> bytes:
    if not replacements:
        return zip_bytes

    with ZipFile(BytesIO(zip_bytes)) as zin:
        infos = zin.infolist()
        names = {info.filename for info in infos}
        missing = sorted(set(replacements) - names)
        if missing:
            raise ZipPartReplacementError(f"Partes ZIP nao encontradas: {', '.join(missing)}")

        raw_end_by_name = _raw_local_record_ends(zip_bytes, infos, getattr(zin, "start_dir", len(zip_bytes)))
        output = BytesIO()
        central_records: list[bytes] = []

        for info in infos:
            local_offset = output.tell()
            replacement = replacements.get(info.filename)
            if replacement is None:
                start = info.header_offset
                end = raw_end_by_name[info.filename]
                output.write(zip_bytes[start:end])
                crc = info.CRC
                compress_size = info.compress_size
                file_size = info.file_size
            else:
                compressed, crc, compress_size, file_size = _compress_payload(
                    replacement,
                    info.compress_type,
                    info.flag_bits,
                )
                output.write(_local_header(info, crc, compress_size, file_size))
                output.write(_filename_bytes(info))
                output.write(info.extra)
                output.write(compressed)

            central_records.append(_central_header(info, crc, compress_size, file_size, local_offset))

        central_offset = output.tell()
        for record in central_records:
            output.write(record)
        central_size = output.tell() - central_offset
        output.write(
            struct.pack(
                "<IHHHHIIH",
                0x06054B50,
                0,
                0,
                len(infos),
                len(infos),
                central_size,
                central_offset,
                len(zin.comment),
            )
        )
        output.write(zin.comment)
    return output.getvalue()


def _raw_local_record_ends(zip_bytes: bytes, infos: list[Any], central_offset: int) -> dict[str, int]:
    sorted_infos = sorted(infos, key=lambda item: item.header_offset)
    output: dict[str, int] = {}
    for index, info in enumerate(sorted_infos):
        output[info.filename] = sorted_infos[index + 1].header_offset if index + 1 < len(sorted_infos) else central_offset
    return output


def _compress_payload(data: bytes, compress_type: int, flag_bits: int) -> tuple[bytes, int, int, int]:
    crc = binascii.crc32(data) & 0xFFFFFFFF
    if compress_type == 0:
        compressed = data
    elif compress_type == ZIP_DEFLATED:
        level = 1 if flag_bits & 0x0006 == 0x0006 else 6
        compressor = zlib.compressobj(level=level, wbits=-15)
        compressed = compressor.compress(data) + compressor.flush()
    else:
        raise ZipPartReplacementError(f"Metodo de compressao ZIP nao suportado: {compress_type}")
    return compressed, crc, len(compressed), len(data)


def _datetime(info: Any) -> tuple[int, int]:
    year, month, day, hour, minute, second = info.date_time
    dos_time = (hour << 11) | (minute << 5) | (second // 2)
    dos_date = ((year - 1980) << 9) | (month << 5) | day
    return dos_time, dos_date


def _filename_bytes(info: Any) -> bytes:
    encoding = "utf-8" if info.flag_bits & 0x800 else "cp437"
    return info.filename.encode(encoding)


def _local_header(info: Any, crc: int, compress_size: int, file_size: int) -> bytes:
    filename = _filename_bytes(info)
    dos_time, dos_date = _datetime(info)
    return struct.pack(
        "<IHHHHHIIIHH",
        0x04034B50,
        info.extract_version,
        info.flag_bits,
        info.compress_type,
        dos_time,
        dos_date,
        crc,
        compress_size,
        file_size,
        len(filename),
        len(info.extra),
    )


def _central_header(info: Any, crc: int, compress_size: int, file_size: int, local_offset: int) -> bytes:
    filename = _filename_bytes(info)
    comment = info.comment or b""
    dos_time, dos_date = _datetime(info)
    version_made_by = (info.create_system << 8) | info.create_version
    return (
        struct.pack(
            "<IHHHHHHIIIHHHHHII",
            0x02014B50,
            version_made_by,
            info.extract_version,
            info.flag_bits,
            info.compress_type,
            dos_time,
            dos_date,
            crc,
            compress_size,
            file_size,
            len(filename),
            len(info.extra),
            len(comment),
            getattr(info, "volume", 0),
            info.internal_attr,
            info.external_attr,
            local_offset,
        )
        + filename
        + info.extra
        + comment
    )
