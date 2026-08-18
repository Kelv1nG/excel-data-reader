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
