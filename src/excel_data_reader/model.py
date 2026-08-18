"""Adapter-neutral immutable result models."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from enum import StrEnum
from types import MappingProxyType
from typing import Any

from excel_data_reader.diagnostics import (
    Diagnostic,
    DiagnosticCode,
    ExcelDataReaderError,
)
from excel_data_reader.normalization import normalize_header


class ValueMode(StrEnum):
    FORMULA = "formula"
    CACHED = "cached"
    BOTH = "both"


class WorkbookFormat(StrEnum):
    OOXML = "ooxml"
    LEGACY_XLS = "xls"


class MatchSource(StrEnum):
    EXPLICIT_RANGE = "explicit_range"
    NATIVE_TABLE = "native_table"
    NAMED_RANGE = "named_range"
    HEADER = "header"


class Confidence(StrEnum):
    STRUCTURAL = "structural"
    HIGH = "high"


class BodyPolicyMode(StrEnum):
    BLANK_ROWS = "blank_rows"
    LAST_POPULATED = "last_populated"
    EXPLICIT = "explicit"


class CandidateReason(StrEnum):
    OUTSIDE_WITHIN = "outside_within"
    MISSING_REQUIRED_HEADERS = "missing_required_headers"
    NON_ADJACENT_COLUMNS = "non_adjacent_columns"
    EXPLICIT_BOTTOM_BEFORE_HEADER = "explicit_bottom_before_header"
    SHADOWED_BY_NATIVE_TABLE = "shadowed_by_native_table"
    FARTHER_FROM_NEAR = "farther_from_near"
    CANDIDATE_LIMIT = "candidate_limit"


@dataclass(frozen=True)
class BodyPolicy:
    """Choose how a header-discovered table body ends.

    Native Excel Tables always keep their authored boundary. These policies are
    used only for tables inferred from a worksheet header row.
    """

    mode: BodyPolicyMode = BodyPolicyMode.BLANK_ROWS
    blank_rows: int = 2
    bottom_row: int | None = None

    def __post_init__(self) -> None:
        object.__setattr__(self, "mode", BodyPolicyMode(self.mode))
        if self.mode is BodyPolicyMode.BLANK_ROWS and self.blank_rows < 1:
            raise ValueError("blank_rows must be at least one")
        if self.mode is BodyPolicyMode.EXPLICIT:
            if self.bottom_row is None or self.bottom_row < 1:
                raise ValueError("explicit body policy requires a positive bottom_row")
        elif self.bottom_row is not None:
            raise ValueError("bottom_row is only valid for the explicit body policy")

    @classmethod
    def until_blank_rows(cls, count: int = 2) -> BodyPolicy:
        return cls(BodyPolicyMode.BLANK_ROWS, blank_rows=count)

    @classmethod
    def last_populated(cls) -> BodyPolicy:
        return cls(BodyPolicyMode.LAST_POPULATED)

    @classmethod
    def through_row(cls, bottom_row: int) -> BodyPolicy:
        return cls(BodyPolicyMode.EXPLICIT, bottom_row=bottom_row)


@dataclass(frozen=True)
class TableQuery:
    """A reusable, deterministic header-table query."""

    required_headers: Sequence[str]
    optional_headers: Sequence[str] = ()
    aliases: Mapping[str, Sequence[str]] = field(default_factory=dict)
    sheet: str | None = None
    allow_non_adjacent_columns: bool = True
    body: BodyPolicy = field(default_factory=BodyPolicy)
    near: Coordinate | str | None = None
    within: Rectangle | str | None = None

    def __post_init__(self) -> None:
        required = self._headers_tuple(self.required_headers)
        optional = self._headers_tuple(self.optional_headers)
        aliases = {
            str(header): self._headers_tuple(values) for header, values in self.aliases.items()
        }
        object.__setattr__(self, "required_headers", required)
        object.__setattr__(self, "optional_headers", optional)
        object.__setattr__(self, "aliases", MappingProxyType(aliases))

    @staticmethod
    def _headers_tuple(values: Sequence[str] | str) -> tuple[str, ...]:
        if isinstance(values, str):
            return (values,)
        return tuple(values)


def _column_letters(index: int) -> str:
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


@dataclass(frozen=True, order=True)
class Coordinate:
    row: int
    column: int

    def __post_init__(self) -> None:
        if self.row < 1 or self.column < 1:
            raise ValueError("worksheet coordinates are one-based")

    @property
    def a1(self) -> str:
        return f"{_column_letters(self.column)}{self.row}"


@dataclass(frozen=True)
class Rectangle:
    top: int
    left: int
    bottom: int
    right: int

    def __post_init__(self) -> None:
        if min(self.top, self.left) < 1 or self.bottom < self.top or self.right < self.left:
            raise ValueError("invalid worksheet rectangle")

    @property
    def height(self) -> int:
        return self.bottom - self.top + 1

    @property
    def width(self) -> int:
        return self.right - self.left + 1

    @property
    def area(self) -> int:
        return self.height * self.width

    @property
    def a1(self) -> str:
        start = Coordinate(self.top, self.left).a1
        end = Coordinate(self.bottom, self.right).a1
        return start if start == end else f"{start}:{end}"

    def contains(self, coordinate: Coordinate) -> bool:
        return (
            self.top <= coordinate.row <= self.bottom
            and self.left <= coordinate.column <= self.right
        )

    def contains_rectangle(self, rectangle: Rectangle) -> bool:
        return (
            self.top <= rectangle.top
            and self.left <= rectangle.left
            and self.bottom >= rectangle.bottom
            and self.right >= rectangle.right
        )


@dataclass(frozen=True)
class HeaderEvidence:
    requested_header: str
    required: bool
    coordinates: tuple[Coordinate, ...]
    raw_headers: tuple[str, ...]

    @property
    def matched(self) -> bool:
        return bool(self.raw_headers)


@dataclass(frozen=True)
class SheetScan:
    sheet: str
    bounds: Rectangle | None
    cells_considered: int
    completed: bool
    diagnostics: tuple[Diagnostic, ...] = ()


@dataclass(frozen=True)
class DiscoveryCandidate:
    sheet: str
    source: MatchSource
    header_row: int | None
    bounds: Rectangle | None
    evidence: tuple[HeaderEvidence, ...]
    produced_matches: int
    selected: bool
    reasons: tuple[CandidateReason, ...] = ()
    name: str | None = None
    distance_from_near: int | None = None


@dataclass(frozen=True)
class RangeReference:
    sheet: str
    bounds: Rectangle

    @property
    def a1(self) -> str:
        return f"'{self.sheet}'!{self.bounds.a1}"


@dataclass(frozen=True)
class FormulaValue:
    formula: Any
    cached: Any


@dataclass(frozen=True)
class CellData:
    sheet: str
    coordinate: Coordinate
    value: Any
    formula: Any | None
    cached_value: Any
    data_type: str
    number_format: str
    is_date: bool
    hidden_row: bool
    hidden_column: bool

    @property
    def address(self) -> str:
        return self.coordinate.a1


@dataclass(frozen=True)
class ColumnInfo:
    name: str
    source_column: int
    raw_header: Any | None = None
    requested_header: str | None = None
    header_coordinate: Coordinate | None = None


@dataclass(frozen=True)
class TableMatch:
    sheet: str
    bounds: Rectangle
    columns: tuple[ColumnInfo, ...]
    data_start_row: int
    data_end_row: int
    source: MatchSource
    confidence: Confidence
    header_row: int | None = None
    name: str | None = None
    diagnostics: tuple[Diagnostic, ...] = ()

    @property
    def range(self) -> str:
        return self.bounds.a1

    @property
    def is_empty(self) -> bool:
        return self.data_end_row < self.data_start_row


@dataclass(frozen=True)
class MatchSet:
    matches: tuple[TableMatch, ...]
    diagnostics: tuple[Diagnostic, ...] = ()

    def require_one(self) -> TableMatch:
        if len(self.matches) == 1:
            return self.matches[0]
        if not self.matches:
            diagnostics = self.diagnostics or (
                Diagnostic(DiagnosticCode.TABLE_NOT_FOUND, "no table matched the query"),
            )
            raise ExcelDataReaderError(diagnostics)
        locations = ", ".join(f"{match.sheet}!{match.range}" for match in self.matches)
        raise ExcelDataReaderError(
            Diagnostic(
                DiagnosticCode.AMBIGUOUS_TABLE,
                f"query matched {len(self.matches)} tables: {locations}",
            )
        )


@dataclass(frozen=True)
class DiscoveryReport:
    query: TableQuery
    result: MatchSet
    scans: tuple[SheetScan, ...]
    candidates: tuple[DiscoveryCandidate, ...]

    @property
    def selected_matches(self) -> tuple[TableMatch, ...]:
        return self.result.matches

    @property
    def diagnostics(self) -> tuple[Diagnostic, ...]:
        return self.result.diagnostics


@dataclass(frozen=True)
class DataRow:
    source_row: int
    cells: tuple[CellData, ...]

    @property
    def values(self) -> tuple[Any, ...]:
        return tuple(cell.value for cell in self.cells)


@dataclass(frozen=True)
class TableData:
    match: TableMatch
    rows: tuple[DataRow, ...]

    @property
    def columns(self) -> tuple[ColumnInfo, ...]:
        return self.match.columns

    @property
    def values(self) -> tuple[tuple[Any, ...], ...]:
        return tuple(row.values for row in self.rows)

    def records(self) -> tuple[Mapping[str, Any], ...]:
        names = tuple(column.name for column in self.columns)
        normalized = tuple(normalize_header(name) for name in names)
        duplicates = sorted({name for name in normalized if normalized.count(name) > 1})
        if duplicates:
            raise ExcelDataReaderError(
                Diagnostic(
                    DiagnosticCode.DUPLICATE_HEADER,
                    "normalized column names are not unique: " + ", ".join(duplicates),
                    sheet=self.match.sheet,
                    address=(
                        Coordinate(self.match.header_row, self.match.bounds.left).a1
                        if self.match.header_row is not None
                        else None
                    ),
                )
            )
        return tuple(
            MappingProxyType(dict(zip(names, row.values, strict=True))) for row in self.rows
        )


@dataclass(frozen=True)
class SheetData:
    name: str
    cells: tuple[CellData, ...]
    bounds: Rectangle | None

    def to_matrix(self, *, fill: Any = None) -> tuple[tuple[Any, ...], ...]:
        if self.bounds is None:
            return ()
        values = {(cell.coordinate.row, cell.coordinate.column): cell.value for cell in self.cells}
        return tuple(
            tuple(
                values.get((row, column), fill)
                for column in range(self.bounds.left, self.bounds.right + 1)
            )
            for row in range(self.bounds.top, self.bounds.bottom + 1)
        )


@dataclass(frozen=True)
class NativeTableInfo:
    name: str
    sheet: str
    bounds: Rectangle
    column_names: tuple[str, ...]
    header_row_count: int
    totals_row_count: int


@dataclass(frozen=True)
class NamedRangeInfo:
    name: str
    scope: str | None
    value: str
    destinations: tuple[RangeReference, ...]
    is_dynamic: bool
    is_resolvable: bool


@dataclass(frozen=True)
class SheetInfo:
    name: str
    state: str
    apparent_bounds: Rectangle | None
    dimension: str
    table_names: tuple[str, ...]
    auto_filter_ref: str | None
    merged_ranges: tuple[str, ...]


@dataclass(frozen=True)
class WorkbookInventory:
    sheets: tuple[SheetInfo, ...]
    native_tables: tuple[NativeTableInfo, ...]
    named_ranges: tuple[NamedRangeInfo, ...]
