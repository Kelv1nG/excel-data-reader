"""Immutable multilevel table queries and coordinate-preserving results."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from types import MappingProxyType
from typing import Any

from excel_data_reader.diagnostics import Diagnostic, DiagnosticCode, ExcelDataReaderError
from excel_data_reader.model import BodyPolicy, Coordinate, DataRow, MatchSource, Rectangle
from excel_data_reader.normalization import normalize_header


@dataclass(frozen=True)
class MultiLevelTableQuery:
    """Select authored labels or complete paths inside a declared header depth."""

    required_headers: Sequence[str | tuple[str, ...]]
    header_depth: int
    sheet: str | None = None
    within: Rectangle | str | None = None
    near: Coordinate | str | None = None
    body: BodyPolicy = field(default_factory=BodyPolicy)

    def __post_init__(self) -> None:
        """Validate the declared depth and freeze nonempty authored selectors."""

        if isinstance(self.header_depth, bool) or not isinstance(self.header_depth, int):
            raise TypeError("header_depth must be an integer")
        if self.header_depth < 1:
            raise ValueError("header_depth must be positive")
        headers = self.required_headers
        if isinstance(headers, str):
            headers = (headers,)
        frozen = tuple(headers)
        if not frozen:
            raise ValueError("at least one required header is needed")
        for selector in frozen:
            if not isinstance(selector, (str, tuple)):
                raise TypeError("selectors must be strings or tuple paths")
            labels = (selector,) if isinstance(selector, str) else selector
            if not labels or any(
                not isinstance(label, str) or not normalize_header(label) for label in labels
            ):
                raise ValueError("header labels and path components must be nonempty strings")
        if self.sheet is not None and (not isinstance(self.sheet, str) or not self.sheet):
            raise ValueError("sheet must be a nonempty string or None")
        if not isinstance(self.body, BodyPolicy):
            raise TypeError("body must be a BodyPolicy")
        object.__setattr__(self, "required_headers", frozen)


@dataclass(frozen=True)
class MultiLevelHeaderPart:
    """One authored label and its span, or a positional label with no invented source."""

    label: str
    raw_header: Any
    source_coordinate: Coordinate | None
    source_bounds: Rectangle | None

    @property
    def is_synthetic(self) -> bool:
        """Return whether the label has no authored source cell."""

        return self.source_coordinate is None


@dataclass(frozen=True)
class MultiLevelColumn:
    """A physical data column and its complete ordered header path."""

    source_column: int
    parts: tuple[MultiLevelHeaderPart, ...]

    @property
    def path(self) -> tuple[str, ...]:
        """Return authored and positional labels in root-to-leaf order."""

        return tuple(part.label for part in self.parts)

    @property
    def name(self) -> str:
        """Return the flat record key, preserving authored spelling."""

        return "__".join(self.path)


def _name_diagnostics(
    columns: tuple[MultiLevelColumn, ...], sheet: str, bounds: Rectangle
) -> tuple[Diagnostic, ...]:
    """Report flat names that collide under the reader's exact normalization.

    Args:
        columns: Projected physical columns to validate.
        sheet: Worksheet containing these columns.
        bounds: Header footprint used for diagnostic locations.
    """

    seen: dict[str, MultiLevelColumn] = {}
    diagnostics: list[Diagnostic] = []
    for column in columns:
        key = normalize_header(column.name)
        if key in seen:
            other = seen[key]
            diagnostics.append(
                Diagnostic(
                    DiagnosticCode.DUPLICATE_MULTILEVEL_HEADER,
                    f"{column.name!r} collides between physical columns "
                    f"{Coordinate(bounds.top, other.source_column).a1} and "
                    f"{Coordinate(bounds.top, column.source_column).a1}",
                    sheet=sheet,
                    address=bounds.a1,
                )
            )
        seen[key] = column
    return tuple(diagnostics)


@dataclass(frozen=True)
class MultiLevelTableMatch:
    """One header footprint and independently bounded table body."""

    sheet: str
    header_bounds: Rectangle
    columns: tuple[MultiLevelColumn, ...]
    data_start_row: int
    data_end_row: int
    diagnostics: tuple[Diagnostic, ...] = ()

    @property
    def bounds(self) -> Rectangle:
        """Return the header footprint extended to the final data row."""

        return Rectangle(
            self.header_bounds.top,
            self.header_bounds.left,
            max(self.header_bounds.bottom, self.data_end_row),
            self.header_bounds.right,
        )

    @property
    def range(self) -> str:
        """Return the full table rectangle in A1 notation."""

        return self.bounds.a1

    @property
    def is_empty(self) -> bool:
        """Return whether the match contains no data rows."""

        return self.data_end_row < self.data_start_row

    @property
    def source(self) -> MatchSource:
        """Identify the match as discovered from worksheet headers."""

        return MatchSource.HEADER


@dataclass(frozen=True)
class MultiLevelTableMatchSet:
    """Ordered matches, diagnostics, and an explicit discovery-completeness flag."""

    matches: tuple[MultiLevelTableMatch, ...]
    diagnostics: tuple[Diagnostic, ...] = ()
    is_complete: bool = True
    candidate_count: int = 0

    def require_one(self) -> MultiLevelTableMatch:
        """Return one proven match or raise incomplete, ambiguity, or missing diagnostics."""

        limits = tuple(d for d in self.diagnostics if d.code is DiagnosticCode.SCAN_LIMIT_EXCEEDED)
        if not self.is_complete or limits:
            raise ExcelDataReaderError(
                limits
                or (Diagnostic(DiagnosticCode.SCAN_LIMIT_EXCEEDED, "discovery is incomplete"),)
            )
        ambiguous = tuple(
            d for d in self.diagnostics if d.code is DiagnosticCode.AMBIGUOUS_MULTILEVEL_HEADER
        )
        if ambiguous:
            raise ExcelDataReaderError(ambiguous)
        if len(self.matches) == 1:
            return self.matches[0]
        if not self.matches:
            raise ExcelDataReaderError(
                self.diagnostics
                or (Diagnostic(DiagnosticCode.TABLE_NOT_FOUND, "no multilevel table matched"),)
            )
        locations = ", ".join(f"{m.sheet}!{m.range}" for m in self.matches)
        raise ExcelDataReaderError(
            Diagnostic(DiagnosticCode.AMBIGUOUS_TABLE, f"multiple tables matched: {locations}")
        )


@dataclass(frozen=True)
class MultiLevelTableData:
    """Typed, source-addressed data rows with flat header-path record keys."""

    match: MultiLevelTableMatch
    rows: tuple[DataRow, ...]

    @property
    def columns(self) -> tuple[MultiLevelColumn, ...]:
        """Return the selected physical columns and their header paths."""

        return self.match.columns

    @property
    def values(self) -> tuple[tuple[Any, ...], ...]:
        """Return typed row values in projection order."""

        return tuple(row.values for row in self.rows)

    def records(self) -> tuple[Mapping[str, Any], ...]:
        """Return immutable flat records, rejecting duplicate normalized keys."""

        diagnostics = _name_diagnostics(self.columns, self.match.sheet, self.match.header_bounds)
        if diagnostics:
            raise ExcelDataReaderError(diagnostics)
        names = tuple(column.name for column in self.columns)
        return tuple(
            MappingProxyType(dict(zip(names, row.values, strict=True))) for row in self.rows
        )
