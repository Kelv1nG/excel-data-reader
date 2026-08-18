"""Pre-parse validation for untrusted OOXML workbook uploads."""

from __future__ import annotations

import hashlib
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from zipfile import (
    ZIP_DEFLATED,
    ZIP_STORED,
    BadZipFile,
    LargeZipFile,
    ZipFile,
    ZipInfo,
    is_zipfile,
)

from excel_data_reader.diagnostics import (
    Diagnostic,
    DiagnosticCode,
    ExcelDataReaderError,
)
from excel_data_reader.model import WorkbookFormat

_MIB = 1024 * 1024
_OLE_COMPOUND_MAGIC = bytes.fromhex("D0CF11E0A1B11AE1")
_REQUIRED_OOXML_MEMBERS = frozenset({"[Content_Types].xml", "xl/workbook.xml"})
_MACRO_EXTENSIONS = frozenset({".xlsm", ".xltm"})
_SUPPORTED_COMPRESSION = frozenset({ZIP_STORED, ZIP_DEFLATED})


class WorkbookRejectedError(ExcelDataReaderError):
    """Raised when an untrusted workbook fails pre-parse policy."""


@dataclass(frozen=True)
class WorkbookPolicy:
    allowed_extensions: frozenset[str] = field(
        default_factory=lambda: frozenset({".xls", ".xlsx", ".xlsm", ".xltx", ".xltm"})
    )
    max_file_size: int = 50 * _MIB
    max_archive_entries: int = 10_000
    max_uncompressed_size: int = 250 * _MIB
    max_member_size: int = 100 * _MIB
    max_compression_ratio: float = 200.0
    max_member_name_length: int = 512
    allow_macros: bool = False
    allow_external_links: bool = False

    def __post_init__(self) -> None:
        """Normalize extensions and validate every configured resource limit."""

        extensions = frozenset(
            item.casefold() if item.startswith(".") else f".{item.casefold()}"
            for item in self.allowed_extensions
        )
        object.__setattr__(self, "allowed_extensions", extensions)
        numeric_limits = {
            "max_file_size": self.max_file_size,
            "max_archive_entries": self.max_archive_entries,
            "max_uncompressed_size": self.max_uncompressed_size,
            "max_member_size": self.max_member_size,
            "max_member_name_length": self.max_member_name_length,
        }
        invalid = [name for name, value in numeric_limits.items() if value < 1]
        if invalid:
            raise ValueError(f"workbook policy limits must be positive: {', '.join(invalid)}")
        if self.max_compression_ratio < 1:
            raise ValueError("max_compression_ratio must be at least one")
        if not extensions:
            raise ValueError("allowed_extensions cannot be empty")


@dataclass(frozen=True)
class WorkbookInspection:
    format: WorkbookFormat
    extension: str
    file_size: int
    sha256: str
    archive_entries: int | None
    compressed_size: int | None
    uncompressed_size: int | None
    largest_member_size: int | None
    maximum_compression_ratio: float | None
    has_macros: bool | None
    has_external_links: bool | None


def inspect_workbook(
    path: str | Path,
    policy: WorkbookPolicy | None = None,
    *,
    checkpoint: Callable[[], None] | None = None,
) -> WorkbookInspection:
    """Validate a supported workbook container before parsing cell data.

    Args:
        path: Filesystem path to the workbook container.
        policy: Validation limits and allowed workbook features. Uses the default
            :class:`WorkbookPolicy` when omitted.
        checkpoint: Optional callback invoked during long-running work to support
            cancellation and deadlines.
    """

    workbook_path = Path(path)
    active_policy = policy or WorkbookPolicy()
    _checkpoint(checkpoint)
    extension = workbook_path.suffix.casefold()
    if extension not in active_policy.allowed_extensions:
        _reject(
            DiagnosticCode.UNSUPPORTED_WORKBOOK_FORMAT,
            f"extension {extension or '<none>'!r} is not allowed",
        )
    try:
        stat = workbook_path.stat()
    except FileNotFoundError as error:
        raise WorkbookRejectedError(
            Diagnostic(DiagnosticCode.FILE_NOT_FOUND, "workbook file was not found")
        ) from error
    if not workbook_path.is_file():
        _reject(DiagnosticCode.FILE_NOT_FOUND, "workbook path is not a regular file")
    if stat.st_size > active_policy.max_file_size:
        _reject(
            DiagnosticCode.WORKBOOK_TOO_LARGE,
            f"file is {stat.st_size:,} bytes; limit is {active_policy.max_file_size:,}",
        )

    _checkpoint(checkpoint)

    with workbook_path.open("rb") as stream:
        signature = stream.read(len(_OLE_COMPOUND_MAGIC))
    if extension == ".xls":
        if signature != _OLE_COMPOUND_MAGIC:
            _reject(
                DiagnosticCode.INVALID_LEGACY_WORKBOOK,
                "file signature is not an Excel 97-2003 compound document",
            )
        return WorkbookInspection(
            format=WorkbookFormat.LEGACY_XLS,
            extension=extension,
            file_size=stat.st_size,
            sha256=_sha256(workbook_path, checkpoint=checkpoint),
            archive_entries=None,
            compressed_size=None,
            uncompressed_size=None,
            largest_member_size=None,
            maximum_compression_ratio=None,
            has_macros=None,
            has_external_links=None,
        )
    if signature == _OLE_COMPOUND_MAGIC:
        raise WorkbookRejectedError(
            Diagnostic(
                DiagnosticCode.ENCRYPTED_WORKBOOK,
                "compound-document input is encrypted or a legacy Excel format",
            )
        )
    if not is_zipfile(workbook_path):
        _reject(
            DiagnosticCode.INVALID_WORKBOOK_ARCHIVE,
            "file signature is not a valid ZIP-based OOXML workbook",
        )

    try:
        with ZipFile(workbook_path, mode="r", allowZip64=True) as archive:
            members = archive.infolist()
            _validate_members(members, active_policy, checkpoint=checkpoint)
            names = {item.filename for item in members}
            missing = sorted(_REQUIRED_OOXML_MEMBERS - names)
            if missing:
                _reject(
                    DiagnosticCode.INVALID_WORKBOOK_ARCHIVE,
                    "archive is missing required OOXML members: " + ", ".join(missing),
                )

            compressed_size = sum(item.compress_size for item in members)
            uncompressed_size = sum(item.file_size for item in members)
            largest_member_size = max((item.file_size for item in members), default=0)
            maximum_ratio = max((_compression_ratio(item) for item in members), default=1.0)
            has_macros = extension in _MACRO_EXTENSIONS or any(
                item.filename.casefold().endswith("/vbaproject.bin") for item in members
            )
            has_external_links = any(
                item.filename.casefold().startswith("xl/externallinks/") for item in members
            )
    except (BadZipFile, LargeZipFile) as error:
        raise WorkbookRejectedError(
            Diagnostic(
                DiagnosticCode.INVALID_WORKBOOK_ARCHIVE,
                "workbook ZIP directory is invalid",
            )
        ) from error

    if has_macros and not active_policy.allow_macros:
        _reject(DiagnosticCode.MACROS_NOT_ALLOWED, "macro-enabled workbooks are not allowed")
    if has_external_links and not active_policy.allow_external_links:
        _reject(
            DiagnosticCode.EXTERNAL_LINKS_NOT_ALLOWED,
            "workbooks containing external links are not allowed",
        )

    return WorkbookInspection(
        format=WorkbookFormat.OOXML,
        extension=extension,
        file_size=stat.st_size,
        sha256=_sha256(workbook_path, checkpoint=checkpoint),
        archive_entries=len(members),
        compressed_size=compressed_size,
        uncompressed_size=uncompressed_size,
        largest_member_size=largest_member_size,
        maximum_compression_ratio=maximum_ratio,
        has_macros=has_macros,
        has_external_links=has_external_links,
    )


def _validate_members(
    members: list[ZipInfo],
    policy: WorkbookPolicy,
    *,
    checkpoint: Callable[[], None] | None,
) -> None:
    """Reject unsafe, encrypted, oversized, or unusually compressed ZIP members."""

    if len(members) > policy.max_archive_entries:
        _reject(
            DiagnosticCode.ARCHIVE_LIMIT_EXCEEDED,
            f"archive has {len(members):,} entries; limit is {policy.max_archive_entries:,}",
        )
    names: set[str] = set()
    total_uncompressed = 0
    for member in members:
        _checkpoint(checkpoint)
        name = member.filename
        if name in names:
            _reject(
                DiagnosticCode.INVALID_WORKBOOK_ARCHIVE,
                f"archive contains duplicate member {name!r}",
            )
        names.add(name)
        path = PurePosixPath(name)
        if (
            len(name) > policy.max_member_name_length
            or path.is_absolute()
            or ".." in path.parts
            or "\\" in name
        ):
            _reject(
                DiagnosticCode.INVALID_WORKBOOK_ARCHIVE,
                f"archive contains unsafe member name {name!r}",
            )
        if member.flag_bits & 0x1:
            raise WorkbookRejectedError(
                Diagnostic(
                    DiagnosticCode.ENCRYPTED_WORKBOOK,
                    f"archive member {name!r} is encrypted",
                )
            )
        if member.compress_type not in _SUPPORTED_COMPRESSION:
            _reject(
                DiagnosticCode.INVALID_WORKBOOK_ARCHIVE,
                f"archive member {name!r} uses unsupported compression",
            )
        if member.file_size > policy.max_member_size:
            _reject(
                DiagnosticCode.ARCHIVE_LIMIT_EXCEEDED,
                f"archive member {name!r} is {member.file_size:,} bytes; "
                f"limit is {policy.max_member_size:,}",
            )
        total_uncompressed += member.file_size
        if total_uncompressed > policy.max_uncompressed_size:
            _reject(
                DiagnosticCode.ARCHIVE_LIMIT_EXCEEDED,
                f"archive expands beyond {policy.max_uncompressed_size:,} bytes",
            )
        ratio = _compression_ratio(member)
        if ratio > policy.max_compression_ratio:
            _reject(
                DiagnosticCode.ARCHIVE_LIMIT_EXCEEDED,
                f"archive member {name!r} has compression ratio {ratio:.1f}; "
                f"limit is {policy.max_compression_ratio:.1f}",
            )


def _compression_ratio(member: ZipInfo) -> float:
    """Return an archive member's expanded-to-compressed size ratio."""

    if member.file_size == 0:
        return 1.0
    if member.compress_size == 0:
        return float("inf")
    return member.file_size / member.compress_size


def _sha256(path: Path, *, checkpoint: Callable[[], None] | None) -> str:
    """Hash a workbook incrementally while honoring cooperative checkpoints."""

    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            _checkpoint(checkpoint)
            digest.update(chunk)
    return digest.hexdigest()


def _checkpoint(callback: Callable[[], None] | None) -> None:
    """Invoke a cooperative execution checkpoint when one is configured."""

    if callback is not None:
        callback()


def _reject(code: DiagnosticCode, message: str) -> None:
    """Raise a workbook-policy rejection with one stable diagnostic."""

    raise WorkbookRejectedError(Diagnostic(code, message))
