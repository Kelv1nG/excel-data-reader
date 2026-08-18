"""Stable diagnostics exposed by the reader."""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass
from enum import StrEnum


class Severity(StrEnum):
    ERROR = "error"
    WARNING = "warning"


class DiagnosticCode(StrEnum):
    TABLE_NOT_FOUND = "TABLE_NOT_FOUND"
    AMBIGUOUS_TABLE = "AMBIGUOUS_TABLE"
    NATIVE_TABLE_NOT_FOUND = "NATIVE_TABLE_NOT_FOUND"
    NAMED_RANGE_NOT_FOUND = "NAMED_RANGE_NOT_FOUND"
    DUPLICATE_HEADER = "DUPLICATE_HEADER"
    INVALID_HEADER_QUERY = "INVALID_HEADER_QUERY"
    INVALID_RANGE = "INVALID_RANGE"
    SHEET_NOT_FOUND = "SHEET_NOT_FOUND"
    NON_RECTANGULAR_NAMED_RANGE = "NON_RECTANGULAR_NAMED_RANGE"
    DYNAMIC_NAMED_RANGE_UNRESOLVED = "DYNAMIC_NAMED_RANGE_UNRESOLVED"
    SCAN_LIMIT_EXCEEDED = "SCAN_LIMIT_EXCEEDED"
    READER_CLOSED = "READER_CLOSED"
    FILE_NOT_FOUND = "FILE_NOT_FOUND"
    UNSUPPORTED_WORKBOOK_FORMAT = "UNSUPPORTED_WORKBOOK_FORMAT"
    WORKBOOK_TOO_LARGE = "WORKBOOK_TOO_LARGE"
    INVALID_WORKBOOK_ARCHIVE = "INVALID_WORKBOOK_ARCHIVE"
    ARCHIVE_LIMIT_EXCEEDED = "ARCHIVE_LIMIT_EXCEEDED"
    ENCRYPTED_WORKBOOK = "ENCRYPTED_WORKBOOK"
    MACROS_NOT_ALLOWED = "MACROS_NOT_ALLOWED"
    EXTERNAL_LINKS_NOT_ALLOWED = "EXTERNAL_LINKS_NOT_ALLOWED"
    ANALYSIS_CANCELLED = "ANALYSIS_CANCELLED"
    ANALYSIS_TIMEOUT = "ANALYSIS_TIMEOUT"
    UPLOAD_STAGING_FAILED = "UPLOAD_STAGING_FAILED"
    INVALID_LEGACY_WORKBOOK = "INVALID_LEGACY_WORKBOOK"
    LEGACY_XLS_LIMITED = "LEGACY_XLS_LIMITED"


@dataclass(frozen=True)
class Diagnostic:
    code: DiagnosticCode
    message: str
    severity: Severity = Severity.ERROR
    sheet: str | None = None
    address: str | None = None

    def __str__(self) -> str:
        location = ""
        if self.sheet is not None:
            location = self.sheet
            if self.address is not None:
                location += f"!{self.address}"
            location += ": "
        return f"{self.code}: {location}{self.message}"


class ExcelDataReaderError(ValueError):
    """Raised when a caller asks for a result that cannot be returned safely."""

    def __init__(self, diagnostics: Diagnostic | Iterable[Diagnostic]):
        normalized = (diagnostics,) if isinstance(diagnostics, Diagnostic) else tuple(diagnostics)
        if not normalized:
            raise ValueError("at least one diagnostic is required")
        self.diagnostics = normalized
        super().__init__("; ".join(str(item) for item in normalized))
