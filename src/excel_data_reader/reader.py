"""OpenPyXL adapter and public workbook reader."""

from __future__ import annotations

from collections.abc import Iterator, Sequence
from dataclasses import replace
from itertools import product
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, range_boundaries
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.worksheet import Worksheet

from excel_data_reader.diagnostics import (
    Diagnostic,
    DiagnosticCode,
    ExcelDataReaderError,
    Severity,
)
from excel_data_reader.model import (
    CellData,
    ColumnInfo,
    Confidence,
    Coordinate,
    DataRow,
    FormulaValue,
    MatchSet,
    MatchSource,
    NamedRangeInfo,
    NativeTableInfo,
    RangeReference,
    Rectangle,
    SheetData,
    SheetInfo,
    TableData,
    TableMatch,
    ValueMode,
    WorkbookInventory,
)
from excel_data_reader.normalization import normalize_header


class ExcelReader:
    """Discover and extract data from one workbook.

    Use :meth:`open` and a context manager so both formula and cached-value
    workbook handles are closed deterministically.
    """

    def __init__(
        self,
        path: Path,
        workbook: Any,
        cached_workbook: Any | None,
        *,
        value_mode: ValueMode,
        max_scan_cells: int,
        max_candidates: int,
    ) -> None:
        self.path = path
        self._workbook = workbook
        self._cached_workbook = cached_workbook
        self.value_mode = value_mode
        self.max_scan_cells = max_scan_cells
        self.max_candidates = max_candidates
        self._closed = False

    @classmethod
    def open(
        cls,
        path: str | Path,
        *,
        value_mode: ValueMode | str = ValueMode.FORMULA,
        max_scan_cells: int = 2_000_000,
        max_candidates: int = 100,
    ) -> ExcelReader:
        """Open a workbook for deterministic discovery and extraction."""

        mode = ValueMode(value_mode)
        if max_scan_cells < 1:
            raise ValueError("max_scan_cells must be positive")
        if max_candidates < 1:
            raise ValueError("max_candidates must be positive")

        workbook_path = Path(path)
        workbook = load_workbook(
            workbook_path,
            read_only=False,
            data_only=False,
            keep_links=False,
        )
        cached_workbook = None
        if mode in {ValueMode.CACHED, ValueMode.BOTH}:
            try:
                cached_workbook = load_workbook(
                    workbook_path,
                    read_only=False,
                    data_only=True,
                    keep_links=False,
                )
            except Exception:
                workbook.close()
                raise
        return cls(
            workbook_path,
            workbook,
            cached_workbook,
            value_mode=mode,
            max_scan_cells=max_scan_cells,
            max_candidates=max_candidates,
        )

    def __enter__(self) -> ExcelReader:
        self._require_open()
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    def close(self) -> None:
        if self._closed:
            return
        self._workbook.close()
        if self._cached_workbook is not None:
            self._cached_workbook.close()
        self._closed = True

    @property
    def sheet_names(self) -> tuple[str, ...]:
        self._require_open()
        return tuple(self._workbook.sheetnames)

    def inventory(self) -> WorkbookInventory:
        """Return workbook-authored structure without extracting table bodies."""

        self._require_open()
        sheets: list[SheetInfo] = []
        native_tables: list[NativeTableInfo] = []
        for sheet in self._workbook.worksheets:
            tables = tuple(self._table_info(sheet, table) for table in sheet.tables.values())
            native_tables.extend(tables)
            sheets.append(
                SheetInfo(
                    name=sheet.title,
                    state=sheet.sheet_state,
                    apparent_bounds=self._apparent_bounds(sheet),
                    dimension=sheet.calculate_dimension(),
                    table_names=tuple(table.name for table in tables),
                    auto_filter_ref=sheet.auto_filter.ref,
                    merged_ranges=tuple(str(item) for item in sheet.merged_cells.ranges),
                )
            )
        return WorkbookInventory(
            sheets=tuple(sheets),
            native_tables=tuple(native_tables),
            named_ranges=tuple(self._iter_named_range_info()),
        )

    def read_range(
        self,
        sheet: str,
        cell_range: str,
        *,
        header: int | None = 0,
    ) -> TableData:
        """Extract one explicit rectangular A1 range."""

        worksheet = self._sheet(sheet)
        bounds = self._parse_rectangle(cell_range, sheet=sheet)
        match = self._match_from_rectangle(
            worksheet,
            bounds,
            header=header,
            source=MatchSource.EXPLICIT_RANGE,
        )
        return self.extract(match)

    def find_native_tables(
        self,
        name: str | None = None,
        *,
        sheet: str | None = None,
    ) -> MatchSet:
        """Find native Excel Tables by optional name and worksheet."""

        self._require_open()
        if sheet is not None:
            self._sheet(sheet)
        matches = tuple(
            self._table_match(worksheet, table)
            for worksheet in self._workbook.worksheets
            if sheet is None or worksheet.title == sheet
            for table in worksheet.tables.values()
            if name is None or self._table_name(table).casefold() == name.casefold()
        )
        diagnostics: tuple[Diagnostic, ...] = ()
        if not matches and name is not None:
            diagnostics = (
                Diagnostic(
                    DiagnosticCode.NATIVE_TABLE_NOT_FOUND,
                    f"native Excel Table {name!r} was not found",
                    sheet=sheet,
                ),
            )
        return MatchSet(matches, diagnostics)

    def get_table(self, name: str, *, sheet: str | None = None) -> TableData:
        """Extract exactly one native Excel Table."""

        return self.extract(self.find_native_tables(name, sheet=sheet).require_one())

    def find_named_ranges(
        self,
        name: str | None = None,
        *,
        sheet: str | None = None,
        header: int | None = 0,
    ) -> MatchSet:
        """Find rectangular destinations of defined names."""

        self._require_open()
        if sheet is not None:
            self._sheet(sheet)

        matching_infos = [
            item
            for item in self._iter_named_range_info()
            if name is None or item.name.casefold() == name.casefold()
        ]
        matches: list[TableMatch] = []
        diagnostics: list[Diagnostic] = []
        for item in matching_infos:
            if not item.is_resolvable:
                diagnostics.append(
                    Diagnostic(
                        (
                            DiagnosticCode.DYNAMIC_NAMED_RANGE_UNRESOLVED
                            if item.is_dynamic
                            else DiagnosticCode.NON_RECTANGULAR_NAMED_RANGE
                        ),
                        f"defined name {item.name!r} does not resolve to a rectangular range",
                        sheet=item.scope,
                    )
                )
                continue
            for destination in item.destinations:
                if sheet is not None and destination.sheet != sheet:
                    continue
                worksheet = self._sheet(destination.sheet)
                matches.append(
                    self._match_from_rectangle(
                        worksheet,
                        destination.bounds,
                        header=header,
                        source=MatchSource.NAMED_RANGE,
                        name=item.name,
                    )
                )

        if not matches and name is not None and not matching_infos:
            diagnostics.append(
                Diagnostic(
                    DiagnosticCode.NAMED_RANGE_NOT_FOUND,
                    f"defined name {name!r} was not found",
                    sheet=sheet,
                )
            )
        elif not matches and name is not None and not diagnostics:
            diagnostics.append(
                Diagnostic(
                    DiagnosticCode.NAMED_RANGE_NOT_FOUND,
                    f"defined name {name!r} has no destination on the requested worksheet",
                    sheet=sheet,
                )
            )
        return MatchSet(tuple(matches), tuple(diagnostics))

    def get_named_range(
        self,
        name: str,
        *,
        sheet: str | None = None,
        header: int | None = 0,
    ) -> TableData:
        """Extract exactly one rectangular defined-name destination."""

        match = self.find_named_ranges(name, sheet=sheet, header=header).require_one()
        return self.extract(match)

    def find_tables(
        self,
        headers: Sequence[str],
        *,
        sheet: str | None = None,
        allow_non_adjacent_columns: bool = True,
        max_blank_rows: int = 2,
    ) -> MatchSet:
        """Find tables containing every exact normalized header on one row."""

        self._require_open()
        requested, normalized = self._validate_header_query(headers)
        if max_blank_rows < 1:
            raise ExcelDataReaderError(
                Diagnostic(
                    DiagnosticCode.INVALID_HEADER_QUERY,
                    "max_blank_rows must be at least one",
                )
            )
        worksheets = (
            (self._sheet(sheet),) if sheet is not None else tuple(self._workbook.worksheets)
        )
        matches: list[TableMatch] = []
        diagnostics: list[Diagnostic] = []
        structural_signatures: set[tuple[str, int, tuple[int, ...]]] = set()

        for worksheet in worksheets:
            for table in worksheet.tables.values():
                projected = self._project_match(
                    self._table_match(worksheet, table),
                    requested,
                    normalized,
                    allow_non_adjacent_columns=allow_non_adjacent_columns,
                )
                for match in projected:
                    matches.append(match)
                    structural_signatures.add(self._header_signature(match))

        for worksheet in worksheets:
            try:
                self._enforce_scan_limit(worksheet)
            except ExcelDataReaderError as error:
                diagnostics.extend(error.diagnostics)
                continue
            stop_sheet = False
            for row in worksheet.iter_rows(
                min_row=1,
                max_row=worksheet.max_row,
                min_col=1,
                max_col=worksheet.max_column,
            ):
                positions: dict[str, list[tuple[int, Any]]] = {header: [] for header in normalized}
                for cell in row:
                    if cell.value is None:
                        continue
                    canonical = normalize_header(cell.value)
                    if canonical in positions:
                        positions[canonical].append((cell.column, cell.value))
                if not all(positions[header] for header in normalized):
                    continue

                duplicate = any(len(positions[header]) > 1 for header in normalized)
                if duplicate:
                    diagnostics.append(
                        Diagnostic(
                            DiagnosticCode.DUPLICATE_HEADER,
                            "a requested header appears more than once on this row",
                            severity=Severity.WARNING,
                            sheet=worksheet.title,
                            address=row[0].coordinate,
                        )
                    )

                for selected in product(*(positions[header] for header in normalized)):
                    selected_columns = tuple(item[0] for item in selected)
                    if len(set(selected_columns)) != len(selected_columns):
                        continue
                    if not allow_non_adjacent_columns and not self._columns_are_adjacent(
                        selected_columns
                    ):
                        continue
                    header_row = row[0].row
                    columns = tuple(
                        ColumnInfo(
                            name=self._column_name(raw, index),
                            source_column=column,
                            raw_header=raw,
                            requested_header=requested[index - 1],
                            header_coordinate=Coordinate(header_row, column),
                        )
                        for index, (column, raw) in enumerate(selected, start=1)
                    )
                    bottom = self._infer_body_bottom(
                        worksheet,
                        header_row,
                        selected_columns,
                        max_blank_rows=max_blank_rows,
                    )
                    match = TableMatch(
                        sheet=worksheet.title,
                        bounds=Rectangle(
                            header_row,
                            min(selected_columns),
                            max(header_row, bottom),
                            max(selected_columns),
                        ),
                        columns=columns,
                        data_start_row=header_row + 1,
                        data_end_row=bottom,
                        source=MatchSource.HEADER,
                        confidence=Confidence.HIGH,
                        header_row=header_row,
                        diagnostics=(diagnostics[-1],) if duplicate else (),
                    )
                    if self._header_signature(match) in structural_signatures:
                        continue
                    matches.append(match)
                    if len(matches) >= self.max_candidates:
                        diagnostics.append(
                            Diagnostic(
                                DiagnosticCode.SCAN_LIMIT_EXCEEDED,
                                f"candidate limit of {self.max_candidates} was reached",
                                sheet=worksheet.title,
                            )
                        )
                        stop_sheet = True
                        break
                if stop_sheet:
                    break

        matches = self._deduplicate_matches(matches)
        if not matches:
            diagnostics.append(
                Diagnostic(
                    DiagnosticCode.TABLE_NOT_FOUND,
                    "no table contained all requested headers",
                    sheet=sheet,
                )
            )
        return MatchSet(tuple(matches), tuple(diagnostics))

    def extract(self, match: TableMatch) -> TableData:
        """Extract the selected logical columns from a discovered match."""

        worksheet = self._sheet(match.sheet)
        hidden_rows = self._hidden_rows(worksheet)
        hidden_columns = self._hidden_columns(worksheet)
        rows: list[DataRow] = []
        if not match.is_empty:
            for row_index in range(match.data_start_row, match.data_end_row + 1):
                cells = tuple(
                    self._cell_data(
                        worksheet,
                        row_index,
                        column.source_column,
                        hidden_rows=hidden_rows,
                        hidden_columns=hidden_columns,
                    )
                    for column in match.columns
                )
                rows.append(DataRow(row_index, cells))
        return TableData(match, tuple(rows))

    def read_sheet(
        self,
        sheet: str,
        *,
        include_styled_blanks: bool = False,
    ) -> SheetData:
        """Return a sparse, coordinate-preserving worksheet snapshot."""

        worksheet = self._sheet(sheet)
        self._enforce_scan_limit(worksheet)
        hidden_rows = self._hidden_rows(worksheet)
        hidden_columns = self._hidden_columns(worksheet)
        cells: list[CellData] = []
        for row in worksheet.iter_rows(
            min_row=1,
            max_row=worksheet.max_row,
            min_col=1,
            max_col=worksheet.max_column,
        ):
            for cell in row:
                if cell.value is None and not (include_styled_blanks and cell.has_style):
                    continue
                cells.append(
                    self._cell_data(
                        worksheet,
                        cell.row,
                        cell.column,
                        hidden_rows=hidden_rows,
                        hidden_columns=hidden_columns,
                    )
                )
        if not cells:
            bounds = None
        else:
            rows = [cell.coordinate.row for cell in cells]
            columns = [cell.coordinate.column for cell in cells]
            bounds = Rectangle(min(rows), min(columns), max(rows), max(columns))
        return SheetData(worksheet.title, tuple(cells), bounds)

    def _require_open(self) -> None:
        if self._closed:
            raise ExcelDataReaderError(
                Diagnostic(DiagnosticCode.READER_CLOSED, "the workbook reader is closed")
            )

    def _sheet(self, name: str) -> Worksheet:
        self._require_open()
        if name not in self._workbook.sheetnames:
            raise ExcelDataReaderError(
                Diagnostic(
                    DiagnosticCode.SHEET_NOT_FOUND,
                    f"worksheet {name!r} was not found",
                    sheet=name,
                )
            )
        return self._workbook[name]

    def _cached_sheet(self, name: str) -> Worksheet | None:
        if self._cached_workbook is None:
            return None
        return self._cached_workbook[name]

    def _cell_data(
        self,
        worksheet: Worksheet,
        row: int,
        column: int,
        *,
        hidden_rows: set[int],
        hidden_columns: set[int],
    ) -> CellData:
        formula_cell = worksheet.cell(row, column)
        cached_sheet = self._cached_sheet(worksheet.title)
        cached_cell = cached_sheet.cell(row, column) if cached_sheet is not None else None
        is_formula = formula_cell.data_type == "f"
        formula = formula_cell.value if is_formula else None
        cached_value = (
            cached_cell.value
            if cached_cell is not None
            else (None if is_formula else formula_cell.value)
        )
        if self.value_mode is ValueMode.FORMULA:
            value = formula_cell.value
        elif self.value_mode is ValueMode.CACHED:
            value = cached_value
        elif is_formula:
            value = FormulaValue(formula, cached_value)
        else:
            value = formula_cell.value
        date_cell = (
            cached_cell
            if cached_cell is not None and cached_cell.value is not None
            else formula_cell
        )
        return CellData(
            sheet=worksheet.title,
            coordinate=Coordinate(row, column),
            value=value,
            formula=formula,
            cached_value=cached_value,
            data_type=str(formula_cell.data_type),
            number_format=str(formula_cell.number_format),
            is_date=bool(date_cell.is_date),
            hidden_row=row in hidden_rows,
            hidden_column=column in hidden_columns,
        )

    def _match_from_rectangle(
        self,
        worksheet: Worksheet,
        bounds: Rectangle,
        *,
        header: int | None,
        source: MatchSource,
        name: str | None = None,
    ) -> TableMatch:
        if header is not None and (header < 0 or bounds.top + header > bounds.bottom):
            raise ExcelDataReaderError(
                Diagnostic(
                    DiagnosticCode.INVALID_RANGE,
                    "header offset lies outside the selected range",
                    sheet=worksheet.title,
                    address=bounds.a1,
                )
            )
        header_row = None if header is None else bounds.top + header
        columns: list[ColumnInfo] = []
        for index, column in enumerate(range(bounds.left, bounds.right + 1), start=1):
            raw = None if header_row is None else worksheet.cell(header_row, column).value
            columns.append(
                ColumnInfo(
                    name=self._column_name(raw, index),
                    source_column=column,
                    raw_header=raw,
                    header_coordinate=(
                        None if header_row is None else Coordinate(header_row, column)
                    ),
                )
            )
        return TableMatch(
            sheet=worksheet.title,
            bounds=bounds,
            columns=tuple(columns),
            data_start_row=bounds.top if header_row is None else header_row + 1,
            data_end_row=bounds.bottom,
            source=source,
            confidence=Confidence.STRUCTURAL,
            header_row=header_row,
            name=name,
        )

    def _table_match(self, worksheet: Worksheet, table: Table) -> TableMatch:
        bounds = self._parse_rectangle(str(table.ref), sheet=worksheet.title)
        header_count = self._optional_int(table.headerRowCount)
        totals_count = self._optional_int(table.totalsRowCount)
        if totals_count == 0 and table.totalsRowShown:
            totals_count = 1
        header_row = bounds.top if header_count else None
        table_names = tuple(str(item) for item in table.column_names)
        columns: list[ColumnInfo] = []
        for index, column in enumerate(range(bounds.left, bounds.right + 1), start=1):
            raw = None if header_row is None else worksheet.cell(header_row, column).value
            if (raw is None or not normalize_header(raw)) and index <= len(table_names):
                raw = table_names[index - 1]
            columns.append(
                ColumnInfo(
                    name=self._column_name(raw, index),
                    source_column=column,
                    raw_header=raw,
                    header_coordinate=(
                        None if header_row is None else Coordinate(header_row, column)
                    ),
                )
            )
        data_start = bounds.top + header_count
        data_end = bounds.bottom - totals_count
        return TableMatch(
            sheet=worksheet.title,
            bounds=bounds,
            columns=tuple(columns),
            data_start_row=data_start,
            data_end_row=data_end,
            source=MatchSource.NATIVE_TABLE,
            confidence=Confidence.STRUCTURAL,
            header_row=header_row,
            name=self._table_name(table),
        )

    def _project_match(
        self,
        match: TableMatch,
        requested: tuple[str, ...],
        normalized: tuple[str, ...],
        *,
        allow_non_adjacent_columns: bool,
    ) -> tuple[TableMatch, ...]:
        positions: dict[str, list[ColumnInfo]] = {header: [] for header in normalized}
        for column in match.columns:
            canonical = normalize_header(column.name)
            if canonical in positions:
                positions[canonical].append(column)
        if not all(positions[header] for header in normalized):
            return ()
        projected: list[TableMatch] = []
        for selected in product(*(positions[header] for header in normalized)):
            physical = tuple(column.source_column for column in selected)
            if len(set(physical)) != len(physical):
                continue
            if not allow_non_adjacent_columns and not self._columns_are_adjacent(physical):
                continue
            columns = tuple(
                replace(column, requested_header=requested[index])
                for index, column in enumerate(selected)
            )
            diagnostics = match.diagnostics
            if any(len(positions[header]) > 1 for header in normalized):
                diagnostics += (
                    Diagnostic(
                        DiagnosticCode.DUPLICATE_HEADER,
                        "a requested header occurs more than once in the table",
                        severity=Severity.WARNING,
                        sheet=match.sheet,
                        address=match.range,
                    ),
                )
            projected.append(replace(match, columns=columns, diagnostics=diagnostics))
        return tuple(projected)

    def _infer_body_bottom(
        self,
        worksheet: Worksheet,
        header_row: int,
        columns: tuple[int, ...],
        *,
        max_blank_rows: int,
    ) -> int:
        last_nonblank = header_row
        blank_run = 0
        for row in range(header_row + 1, worksheet.max_row + 1):
            if any(worksheet.cell(row, column).value is not None for column in columns):
                last_nonblank = row
                blank_run = 0
                continue
            blank_run += 1
            if blank_run >= max_blank_rows:
                break
        return last_nonblank

    def _validate_header_query(
        self, headers: Sequence[str]
    ) -> tuple[tuple[str, ...], tuple[str, ...]]:
        requested = tuple(headers)
        normalized = tuple(normalize_header(header) for header in requested)
        if not requested or any(not header for header in normalized):
            raise ExcelDataReaderError(
                Diagnostic(
                    DiagnosticCode.INVALID_HEADER_QUERY,
                    "at least one non-empty header is required",
                )
            )
        duplicates = sorted({item for item in normalized if normalized.count(item) > 1})
        if duplicates:
            raise ExcelDataReaderError(
                Diagnostic(
                    DiagnosticCode.INVALID_HEADER_QUERY,
                    "requested headers are not unique after normalization: "
                    + ", ".join(duplicates),
                )
            )
        return requested, normalized

    def _iter_named_range_info(self) -> Iterator[NamedRangeInfo]:
        for name, definition in self._workbook.defined_names.items():
            yield self._named_range_info(name, definition, scope=None)
        for worksheet in self._workbook.worksheets:
            for name, definition in worksheet.defined_names.items():
                yield self._named_range_info(name, definition, scope=worksheet.title)

    def _named_range_info(
        self,
        name: str,
        definition: Any,
        *,
        scope: str | None,
    ) -> NamedRangeInfo:
        destinations: list[RangeReference] = []
        unresolved = False
        try:
            raw_destinations = tuple(definition.destinations)
        except (AttributeError, TypeError, ValueError):
            raw_destinations = ()
            unresolved = True
        for sheet, cell_range in raw_destinations:
            try:
                bounds = self._parse_rectangle(str(cell_range), sheet=str(sheet))
            except ExcelDataReaderError:
                unresolved = True
                continue
            destinations.append(RangeReference(str(sheet), bounds))
        if not destinations:
            unresolved = True
        return NamedRangeInfo(
            name=str(name),
            scope=scope,
            value=str(definition.value),
            destinations=tuple(destinations),
            is_dynamic=str(getattr(definition, "type", "")).upper() == "FUNC",
            is_resolvable=bool(destinations) and not unresolved,
        )

    def _table_info(self, worksheet: Worksheet, table: Table) -> NativeTableInfo:
        bounds = self._parse_rectangle(str(table.ref), sheet=worksheet.title)
        totals_count = self._optional_int(table.totalsRowCount)
        if totals_count == 0 and table.totalsRowShown:
            totals_count = 1
        return NativeTableInfo(
            name=self._table_name(table),
            sheet=worksheet.title,
            bounds=bounds,
            column_names=tuple(str(item) for item in table.column_names),
            header_row_count=self._optional_int(table.headerRowCount),
            totals_row_count=totals_count,
        )

    @staticmethod
    def _table_name(table: Table) -> str:
        return str(table.displayName or table.name)

    @staticmethod
    def _optional_int(value: Any) -> int:
        return 0 if value is None else int(value)

    @staticmethod
    def _column_name(raw: Any | None, position: int) -> str:
        if raw is None or not normalize_header(raw):
            return f"column_{position}"
        return str(raw)

    @staticmethod
    def _columns_are_adjacent(columns: tuple[int, ...]) -> bool:
        ordered = sorted(columns)
        return ordered[-1] - ordered[0] + 1 == len(ordered)

    @staticmethod
    def _header_signature(match: TableMatch) -> tuple[str, int, tuple[int, ...]]:
        return (
            match.sheet,
            match.header_row or match.bounds.top,
            tuple(column.source_column for column in match.columns),
        )

    @staticmethod
    def _deduplicate_matches(matches: list[TableMatch]) -> list[TableMatch]:
        unique: dict[tuple[str, int, tuple[int, ...]], TableMatch] = {}
        for match in matches:
            signature = ExcelReader._header_signature(match)
            existing = unique.get(signature)
            if existing is None or match.source is MatchSource.NATIVE_TABLE:
                unique[signature] = match
        return list(unique.values())

    def _parse_rectangle(self, value: str, *, sheet: str) -> Rectangle:
        try:
            min_col, min_row, max_col, max_row = range_boundaries(value)
            if None in {min_col, min_row, max_col, max_row}:
                raise ValueError("range must have finite row and column bounds")
            return Rectangle(int(min_row), int(min_col), int(max_row), int(max_col))
        except (TypeError, ValueError) as error:
            raise ExcelDataReaderError(
                Diagnostic(
                    DiagnosticCode.INVALID_RANGE,
                    f"{value!r} is not one finite rectangular A1 range",
                    sheet=sheet,
                )
            ) from error

    def _enforce_scan_limit(self, worksheet: Worksheet) -> None:
        apparent_cells = int(worksheet.max_row) * int(worksheet.max_column)
        if apparent_cells > self.max_scan_cells:
            raise ExcelDataReaderError(
                Diagnostic(
                    DiagnosticCode.SCAN_LIMIT_EXCEEDED,
                    f"apparent sheet area is {apparent_cells:,} cells; "
                    f"limit is {self.max_scan_cells:,}",
                    sheet=worksheet.title,
                    address=worksheet.calculate_dimension(),
                )
            )

    @staticmethod
    def _apparent_bounds(worksheet: Worksheet) -> Rectangle | None:
        if (
            worksheet.max_row == 1
            and worksheet.max_column == 1
            and worksheet["A1"].value is None
            and not worksheet["A1"].has_style
            and not worksheet.merged_cells.ranges
        ):
            return None
        return Rectangle(1, 1, int(worksheet.max_row), int(worksheet.max_column))

    @staticmethod
    def _hidden_rows(worksheet: Worksheet) -> set[int]:
        return {
            int(index) for index, dimension in worksheet.row_dimensions.items() if dimension.hidden
        }

    @staticmethod
    def _hidden_columns(worksheet: Worksheet) -> set[int]:
        hidden: set[int] = set()
        for key, dimension in worksheet.column_dimensions.items():
            if not dimension.hidden:
                continue
            minimum = int(dimension.min or column_index_from_string(key))
            maximum = int(dimension.max or minimum)
            hidden.update(range(minimum, maximum + 1))
        return hidden
