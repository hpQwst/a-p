from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import PurePosixPath
from zipfile import BadZipFile, ZipFile, ZipInfo
import os


@dataclass(frozen=True)
class ArchiveLimits:
    max_members: int
    max_uncompressed_bytes: int
    max_member_bytes: int
    max_compression_ratio: float


class UnsafeArchiveError(ValueError):
    """Raised when an uploaded Office/ZIP package exceeds the safety policy."""


def archive_limits() -> ArchiveLimits:
    return ArchiveLimits(
        max_members=_env_int("AUTO_PPT_ARCHIVE_MAX_MEMBERS", 5_000),
        max_uncompressed_bytes=_env_mb("AUTO_PPT_ARCHIVE_MAX_UNCOMPRESSED_MB", 1_024),
        max_member_bytes=_env_mb("AUTO_PPT_ARCHIVE_MAX_MEMBER_MB", 512),
        max_compression_ratio=_env_float("AUTO_PPT_ARCHIVE_MAX_RATIO", 1_000.0),
    )


def validate_pptx_bytes(data: bytes, limits: ArchiveLimits | None = None) -> None:
    names = _validate_zip(data, limits or archive_limits(), label="PPTX")
    required = {"[Content_Types].xml", "ppt/presentation.xml"}
    missing = required.difference(names)
    if missing:
        raise UnsafeArchiveError("O arquivo PPTX nao possui a estrutura Office esperada.")


def validate_xlsx_bytes(data: bytes, limits: ArchiveLimits | None = None) -> None:
    names = _validate_zip(data, limits or archive_limits(), label="XLSX")
    required = {"[Content_Types].xml", "xl/workbook.xml"}
    missing = required.difference(names)
    if missing:
        raise UnsafeArchiveError("O arquivo XLSX nao possui a estrutura Office esperada.")


def validate_datasource_zip_bytes(data: bytes, limits: ArchiveLimits | None = None) -> None:
    policy = limits or archive_limits()
    names = _validate_zip(data, policy, label="ZIP de datasources")
    xlsx_names = sorted(name for name in names if name.lower().endswith(".xlsx"))
    if not xlsx_names:
        raise UnsafeArchiveError("O ZIP de datasources nao contem arquivos XLSX.")
    try:
        with ZipFile(BytesIO(data)) as archive:
            for name in xlsx_names:
                validate_xlsx_bytes(archive.read(name), policy)
    except BadZipFile as exc:
        raise UnsafeArchiveError("O ZIP de datasources esta corrompido.") from exc


def _validate_zip(data: bytes, limits: ArchiveLimits, label: str) -> set[str]:
    try:
        with ZipFile(BytesIO(data)) as archive:
            members = archive.infolist()
            if len(members) > limits.max_members:
                raise UnsafeArchiveError(
                    f"{label} possui arquivos internos demais ({len(members)}; limite {limits.max_members})."
                )

            total = 0
            names: set[str] = set()
            for member in members:
                _validate_member(member, limits, label)
                total += member.file_size
                if total > limits.max_uncompressed_bytes:
                    limit_mb = limits.max_uncompressed_bytes // (1024 * 1024)
                    raise UnsafeArchiveError(
                        f"{label} excede o limite descompactado de {limit_mb} MB."
                    )
                names.add(member.filename.replace("\\", "/"))
            return names
    except BadZipFile as exc:
        raise UnsafeArchiveError(f"{label} nao e um arquivo ZIP/Office valido.") from exc


def _validate_member(member: ZipInfo, limits: ArchiveLimits, label: str) -> None:
    normalized = member.filename.replace("\\", "/")
    path = PurePosixPath(normalized)
    if path.is_absolute() or ".." in path.parts:
        raise UnsafeArchiveError(f"{label} possui um caminho interno inseguro.")
    if member.flag_bits & 0x1:
        raise UnsafeArchiveError(f"{label} possui conteudo criptografado e nao suportado.")
    if member.file_size > limits.max_member_bytes:
        limit_mb = limits.max_member_bytes // (1024 * 1024)
        raise UnsafeArchiveError(f"{label} possui um item maior que {limit_mb} MB.")
    if member.file_size <= 0:
        return
    if member.compress_size <= 0:
        raise UnsafeArchiveError(f"{label} possui um item com compactacao suspeita.")
    ratio = member.file_size / member.compress_size
    if ratio > limits.max_compression_ratio:
        raise UnsafeArchiveError(f"{label} possui uma taxa de compactacao suspeita.")


def _env_int(name: str, default: int) -> int:
    try:
        return max(int(os.getenv(name, str(default))), 1)
    except ValueError:
        return default


def _env_float(name: str, default: float) -> float:
    try:
        return max(float(os.getenv(name, str(default))), 1.0)
    except ValueError:
        return default


def _env_mb(name: str, default_mb: int) -> int:
    return _env_int(name, default_mb) * 1024 * 1024
