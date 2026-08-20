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


class MatrixBoundarySource(StrEnum):
    """Evidence that determined the final row of an unpivoted matrix section."""

    MERGED_SECTION = "merged_section"
    NEXT_SECTION = "next_section"
    BLANK_ROWS = "blank_rows"
    LAST_POPULATED = "last_populated"
    EXPLICIT = "explicit"


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
        """Normalize the mode and reject inconsistent boundary settings."""

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
        """Stop a discovered body after the requested consecutive blank rows.

        Args:
            count: Number of consecutive blank rows that ends the table body.
        """

        return cls(BodyPolicyMode.BLANK_ROWS, blank_rows=count)

    @classmethod
    def last_populated(cls) -> BodyPolicy:
        """Continue a discovered body through its last populated selected cell."""

        return cls(BodyPolicyMode.LAST_POPULATED)

    @classmethod
    def through_row(cls, bottom_row: int) -> BodyPolicy:
        """Use an explicit one-based bottom row for a discovered body.

        Args:
            bottom_row: One-based worksheet row at which the table body ends.
        """

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
        """Freeze header and alias sequences into deterministic immutable values."""

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
        """Coerce one header or a header sequence into an immutable tuple.

        Args:
            values: Single header string or ordered header sequence.
        """

        if isinstance(values, str):
            return (values,)
        return tuple(values)


@dataclass(frozen=True)
class MatrixQuery:
    """A deterministic query for row-sectioned matrices with hierarchical headers.

    Attributes:
        sections: Required logical section labels matched with exact normalization.
        header_level_names: Field names assigned to the hierarchical levels in long
            records; the number of names declares the required header depth.
        aliases: Alternate exact-normalized labels keyed by a declared section.
        sheet: Optional exact worksheet name used to limit discovery.
        within: Optional finite A1 rectangle used to limit discovery and scan cost.
        body: Boundary policy used when a section label has no vertical merge.
        header_rows: Optional one-based row override with one row per header level.
        identifier_column: Optional one-based index or Excel column letters identifying
            the row-key column.
    """

    sections: Sequence[str]
    header_level_names: Sequence[str] = ("group", "attribute")
    aliases: Mapping[str, Sequence[str]] = field(default_factory=dict)
    sheet: str | None = None
    within: Rectangle | str | None = None
    body: BodyPolicy = field(default_factory=BodyPolicy)
    header_rows: Sequence[int] | None = None
    identifier_column: int | str | None = None

    def __post_init__(self) -> None:
        """Freeze sequences and reject ambiguous or inconsistent query fields."""

        sections = self._text_tuple(self.sections)
        level_names = self._text_tuple(self.header_level_names)
        normalized_sections = tuple(normalize_header(value) for value in sections)
        normalized_levels = tuple(normalize_header(value) for value in level_names)
        if not sections or any(not value for value in normalized_sections):
            raise ValueError("at least one non-empty matrix section is required")
        if len(set(normalized_sections)) != len(normalized_sections):
            raise ValueError("matrix sections must be unique after normalization")
        if not level_names or any(not value for value in normalized_levels):
            raise ValueError("at least one non-empty matrix header level name is required")
        if len(set(normalized_levels)) != len(normalized_levels):
            raise ValueError("matrix header level names must be unique after normalization")
        reserved = {
            "section",
            "identifier",
            "value",
            "source sheet",
            "source cell",
            "source row",
            "source column",
            "identifier cell",
        }
        collisions = sorted(set(normalized_levels) & reserved)
        if collisions:
            raise ValueError(
                "matrix header level names conflict with record fields: " + ", ".join(collisions)
            )

        aliases = {
            str(section): self._text_tuple(values) for section, values in self.aliases.items()
        }
        header_rows = None if self.header_rows is None else tuple(self.header_rows)
        if header_rows is not None:
            if len(header_rows) != len(level_names):
                raise ValueError("header_rows must contain one row for each header level")
            if (
                any(
                    isinstance(row, bool) or not isinstance(row, int) or row < 1
                    for row in header_rows
                )
                or tuple(sorted(header_rows)) != header_rows
                or len(set(header_rows)) != len(header_rows)
            ):
                raise ValueError("header_rows must be unique positive integers in ascending order")

        identifier_column = self.identifier_column
        if isinstance(identifier_column, str):
            identifier_column = identifier_column.strip().upper()
            if not identifier_column or not identifier_column.isalpha():
                raise ValueError("identifier_column must be a positive index or column letters")
        elif identifier_column is not None and (
            isinstance(identifier_column, bool)
            or not isinstance(identifier_column, int)
            or identifier_column < 1
        ):
            raise ValueError("identifier_column must be a positive index or column letters")
        if not isinstance(self.body, BodyPolicy):
            raise TypeError("body must be a BodyPolicy")

        object.__setattr__(self, "sections", sections)
        object.__setattr__(self, "header_level_names", level_names)
        object.__setattr__(self, "aliases", MappingProxyType(aliases))
        object.__setattr__(self, "header_rows", header_rows)
        object.__setattr__(self, "identifier_column", identifier_column)

    @staticmethod
    def _text_tuple(values: Sequence[str] | str) -> tuple[str, ...]:
        """Coerce one string or an ordered string sequence into a tuple.

        Args:
            values: Single string or ordered sequence to freeze.
        """

        if isinstance(values, str):
            return (values,)
        return tuple(str(value) for value in values)


def _column_letters(index: int) -> str:
    """Convert a positive one-based column index to Excel letters.

    Args:
        index: Positive one-based worksheet column index.
    """

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
        """Reject coordinates outside Excel's one-based address space."""

        if self.row < 1 or self.column < 1:
            raise ValueError("worksheet coordinates are one-based")

    @property
    def a1(self) -> str:
        """Return the coordinate in A1 notation."""

        return f"{_column_letters(self.column)}{self.row}"


@dataclass(frozen=True)
class Rectangle:
    top: int
    left: int
    bottom: int
    right: int

    def __post_init__(self) -> None:
        """Reject inverted or non-positive rectangle bounds."""

        if min(self.top, self.left) < 1 or self.bottom < self.top or self.right < self.left:
            raise ValueError("invalid worksheet rectangle")

    @property
    def height(self) -> int:
        """Return the inclusive row count."""

        return self.bottom - self.top + 1

    @property
    def width(self) -> int:
        """Return the inclusive column count."""

        return self.right - self.left + 1

    @property
    def area(self) -> int:
        """Return the number of cells in the rectangle."""

        return self.height * self.width

    @property
    def a1(self) -> str:
        """Return the rectangle in compact A1 notation."""

        start = Coordinate(self.top, self.left).a1
        end = Coordinate(self.bottom, self.right).a1
        return start if start == end else f"{start}:{end}"

    def contains(self, coordinate: Coordinate) -> bool:
        """Return whether a coordinate lies inside the inclusive bounds.

        Args:
            coordinate: Worksheet coordinate to test.
        """

        return (
            self.top <= coordinate.row <= self.bottom
            and self.left <= coordinate.column <= self.right
        )

    def contains_rectangle(self, rectangle: Rectangle) -> bool:
        """Return whether another rectangle lies fully inside these bounds.

        Args:
            rectangle: Worksheet rectangle to test.
        """

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
        """Return whether at least one physical header supplied evidence."""

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
        """Return the sheet-qualified range in A1 notation."""

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
        """Return the source cell coordinate in A1 notation."""

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
        """Return the matched rectangle in A1 notation."""

        return self.bounds.a1

    @property
    def is_empty(self) -> bool:
        """Return whether the match contains no data rows."""

        return self.data_end_row < self.data_start_row


@dataclass(frozen=True)
class MatchSet:
    matches: tuple[TableMatch, ...]
    diagnostics: tuple[Diagnostic, ...] = ()

    def require_one(self) -> TableMatch:
        """Return the sole match or raise structured not-found or ambiguity diagnostics."""

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
        """Return the final matches selected by discovery."""

        return self.result.matches

    @property
    def diagnostics(self) -> tuple[Diagnostic, ...]:
        """Return diagnostics attached to the discovery result."""

        return self.result.diagnostics


@dataclass(frozen=True)
class DataRow:
    source_row: int
    cells: tuple[CellData, ...]

    @property
    def values(self) -> tuple[Any, ...]:
        """Return cell values in logical column order."""

        return tuple(cell.value for cell in self.cells)


@dataclass(frozen=True)
class TableData:
    match: TableMatch
    rows: tuple[DataRow, ...]

    @property
    def columns(self) -> tuple[ColumnInfo, ...]:
        """Return the logical columns selected by the match."""

        return self.match.columns

    @property
    def values(self) -> tuple[tuple[Any, ...], ...]:
        """Return table values as immutable rows in logical column order."""

        return tuple(row.values for row in self.rows)

    def records(self) -> tuple[Mapping[str, Any], ...]:
        """Return immutable name-to-value rows when logical column names are unique."""

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
class MatrixHeader:
    """One physical value column identified by an ordered hierarchical header path.

    Attributes:
        labels: Authored header labels from outermost to innermost level.
        coordinates: Source coordinates corresponding to each header label.
        source_column: One-based physical worksheet column containing matrix values.
    """

    labels: tuple[str, ...]
    coordinates: tuple[Coordinate, ...]
    source_column: int

    def flatten(self, separator: str = "__") -> str:
        """Join the hierarchical labels into one display or wide-record name.

        Args:
            separator: Text inserted between adjacent header levels.
        """

        if not separator:
            raise ValueError("matrix header separator cannot be empty")
        return separator.join(self.labels)


@dataclass(frozen=True)
class MatrixMatch:
    """A discovered matrix row section and its shared hierarchical columns.

    Attributes:
        section: Logical requested section name.
        raw_section: Authored label value found in the workbook.
        sheet: Worksheet containing the section.
        anchor: Cell or merged rectangle containing the section label.
        header_level_names: Field names assigned to hierarchical header levels.
        header_rows: One-based authored rows supplying the header levels.
        headers: Resolved hierarchical value columns in physical order.
        identifier_column: One-based physical row-identifier column.
        data_bounds: Inclusive rectangle from identifiers through matrix values.
        boundary_source: Evidence that determined the final section row.
        diagnostics: Warnings attached specifically to this section.
    """

    section: str
    raw_section: Any
    sheet: str
    anchor: Rectangle
    header_level_names: tuple[str, ...]
    header_rows: tuple[int, ...]
    headers: tuple[MatrixHeader, ...]
    identifier_column: int
    data_bounds: Rectangle
    boundary_source: MatrixBoundarySource
    diagnostics: tuple[Diagnostic, ...] = ()

    @property
    def range(self) -> str:
        """Return the section's data rectangle in A1 notation."""

        return self.data_bounds.a1


@dataclass(frozen=True)
class MatrixMatchSet:
    """Discovered matrix sections plus stable missing or ambiguity diagnostics.

    Attributes:
        matches: Matrix sections discovered across the requested worksheets.
        diagnostics: Stable missing, ambiguity, resource, or malformed-layout details.
    """

    matches: tuple[MatrixMatch, ...]
    diagnostics: tuple[Diagnostic, ...] = ()

    def require_section(self, section: str) -> MatrixMatch:
        """Return one logical section or raise not-found or ambiguity diagnostics.

        Args:
            section: Requested logical section name matched after normalization.
        """

        normalized = normalize_header(section)
        matches = tuple(
            match for match in self.matches if normalize_header(match.section) == normalized
        )
        if len(matches) == 1:
            return matches[0]
        if not matches:
            raise ExcelDataReaderError(
                self.diagnostics
                or Diagnostic(
                    DiagnosticCode.MATRIX_SECTION_NOT_FOUND,
                    f"matrix section {section!r} was not found",
                )
            )
        locations = ", ".join(f"{match.sheet}!{match.anchor.a1}" for match in matches)
        raise ExcelDataReaderError(
            Diagnostic(
                DiagnosticCode.AMBIGUOUS_MATRIX_SECTION,
                f"matrix section {section!r} matched {len(matches)} anchors: {locations}",
            )
        )


@dataclass(frozen=True)
class MatrixValue:
    """One source-addressed value at a row identifier and hierarchical column path.

    Attributes:
        section: Logical section containing the value.
        identifier: Authored identifier for the value's physical row.
        identifier_cell: Coordinate-preserving identifier cell data.
        header: Hierarchical header path for the value's physical column.
        cell: Coordinate-preserving matrix value cell data.
    """

    section: str
    identifier: Any
    identifier_cell: CellData
    header: MatrixHeader
    cell: CellData


@dataclass(frozen=True)
class MatrixData:
    """Extracted values for one discovered matrix section.

    Attributes:
        match: Matrix section and hierarchical columns used for extraction.
        values: Source-addressed values in physical row-major order.
    """

    match: MatrixMatch
    values: tuple[MatrixValue, ...]

    def long_records(
        self,
        *,
        include_blank_values: bool = True,
    ) -> tuple[Mapping[str, Any], ...]:
        """Return schema-stable analytical records with one row per matrix cell.

        Args:
            include_blank_values: Whether records whose matrix value is ``None`` are
                retained.
        """

        records: list[Mapping[str, Any]] = []
        for item in self.values:
            if item.cell.value is None and not include_blank_values:
                continue
            record: dict[str, Any] = {
                "section": item.section,
                "identifier": item.identifier,
            }
            record.update(zip(self.match.header_level_names, item.header.labels, strict=True))
            record.update(
                {
                    "value": item.cell.value,
                    "source_sheet": item.cell.sheet,
                    "source_cell": item.cell.address,
                    "source_row": item.cell.coordinate.row,
                    "source_column": item.cell.coordinate.column,
                    "identifier_cell": item.identifier_cell.address,
                }
            )
            records.append(MappingProxyType(record))
        return tuple(records)

    def wide_records(self, *, separator: str = "__") -> tuple[Mapping[str, Any], ...]:
        """Return one record per identifier with flattened hierarchical headers.

        Args:
            separator: Text inserted between header levels in generated column names.
        """

        names = tuple(header.flatten(separator) for header in self.match.headers)
        normalized = tuple(normalize_header(name) for name in names)
        fixed = {
            normalize_header(name)
            for name in (
                "section",
                "identifier",
                "source_sheet",
                "source_row",
                "identifier_cell",
            )
        }
        duplicates = sorted({name for name in normalized if normalized.count(name) > 1})
        collisions = sorted(set(normalized) & fixed)
        if duplicates or collisions:
            details = duplicates + collisions
            raise ExcelDataReaderError(
                Diagnostic(
                    DiagnosticCode.DUPLICATE_MATRIX_HEADER,
                    "flattened matrix headers are not unique: " + ", ".join(details),
                    sheet=self.match.sheet,
                    address=self.match.data_bounds.a1,
                )
            )

        by_row: dict[int, dict[str, Any]] = {}
        header_names = {
            header.source_column: name
            for header, name in zip(self.match.headers, names, strict=True)
        }
        for item in self.values:
            row = item.cell.coordinate.row
            record = by_row.setdefault(
                row,
                {
                    "section": item.section,
                    "identifier": item.identifier,
                    **dict.fromkeys(names),
                    "source_sheet": item.cell.sheet,
                    "source_row": row,
                    "identifier_cell": item.identifier_cell.address,
                },
            )
            record[header_names[item.header.source_column]] = item.cell.value
        return tuple(MappingProxyType(record) for record in by_row.values())


@dataclass(frozen=True)
class SheetData:
    name: str
    cells: tuple[CellData, ...]
    bounds: Rectangle | None

    def to_matrix(self, *, fill: Any = None) -> tuple[tuple[Any, ...], ...]:
        """Materialize the sparse sheet bounds as a dense immutable matrix.

        Args:
            fill: Value used for coordinates absent from the sparse cell collection.
        """

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
