"""Reusable recipes covering the most common public API entry points."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any, BinaryIO

from excel_data_reader import (
    AnalysisControl,
    AnalysisRequest,
    AnalysisResponse,
    BodyPolicy,
    DiscoveryReport,
    ExcelReader,
    MatchSet,
    MatrixData,
    MatrixQuery,
    SheetData,
    TableData,
    TableQuery,
    ValueMode,
    WorkbookInventory,
    analyze_workbook,
    analyze_workbook_bytes,
    to_json,
)

WORKBOOKS = Path(__file__).parent / "workbooks"


def inspect_structure(path: str | Path) -> WorkbookInventory:
    """List sheets, native tables, and defined names without extracting data."""

    with ExcelReader.open(path) as reader:
        return reader.inventory()


def read_native_table(
    path: str | Path,
    table_name: str,
    *,
    sheet: str | None = None,
    value_mode: ValueMode | str = ValueMode.FORMULA,
) -> TableData:
    """Extract one workbook-authored Excel Table by name."""

    with ExcelReader.open(path, value_mode=value_mode) as reader:
        return reader.get_table(table_name, sheet=sheet)


def read_defined_name(
    path: str | Path,
    name: str,
    *,
    sheet: str | None = None,
    header: int | None = 0,
) -> TableData:
    """Extract one rectangular workbook or worksheet defined name."""

    with ExcelReader.open(path) as reader:
        return reader.get_named_range(name, sheet=sheet, header=header)


def read_known_range(
    path: str | Path,
    sheet: str,
    cell_range: str,
    *,
    header: int | None = 0,
) -> TableData:
    """Extract a known A1 rectangle, optionally generating column names."""

    with ExcelReader.open(path) as reader:
        return reader.read_range(sheet, cell_range, header=header)


def read_sparse_sheet(
    path: str | Path,
    sheet: str,
    *,
    include_styled_blanks: bool = False,
) -> SheetData:
    """Read populated worksheet cells without guessing a table boundary."""

    with ExcelReader.open(path) as reader:
        return reader.read_sheet(sheet, include_styled_blanks=include_styled_blanks)


def find_matching_tables(
    path: str | Path,
    headers: Sequence[str],
    *,
    sheet: str | None = None,
) -> MatchSet:
    """Return every exact header match so the caller can handle ambiguity."""

    with ExcelReader.open(path) as reader:
        return reader.find_tables(headers, sheet=sheet)


def extract_matching_table(path: str | Path, query: TableQuery) -> TableData:
    """Require one structured-query match and extract its projected columns."""

    with ExcelReader.open(path) as reader:
        match = reader.query_tables(query).require_one()
        return reader.extract(match)


def explain_matching_tables(path: str | Path, query: TableQuery) -> DiscoveryReport:
    """Return scan evidence, candidates, and diagnostics for a table query."""

    with ExcelReader.open(path) as reader:
        return reader.explain(query)


def extract_matrix_section(
    path: str | Path,
    query: MatrixQuery,
    section: str,
) -> MatrixData:
    """Discover and extract one logical section of a grouped-header matrix."""

    with ExcelReader.open(path) as reader:
        match = reader.find_matrices(query).require_section(section)
        return reader.extract_matrix(match)


def analyze_trusted_file(
    path: str | Path,
    query: TableQuery,
    *,
    max_output_rows: int = 500,
) -> AnalysisResponse:
    """Analyze a trusted filesystem path through the versioned service API."""

    request = AnalysisRequest.find_tables(
        query,
        include_rows=True,
        max_output_rows=max_output_rows,
    )
    return analyze_workbook(path, request)


def analyze_uploaded_stream(
    stream: BinaryIO,
    filename: str,
    query: TableQuery,
    *,
    timeout_seconds: float = 10,
    max_output_rows: int = 500,
) -> AnalysisResponse:
    """Validate, stage, and analyze an uploaded binary stream with bounded output."""

    request = AnalysisRequest.find_tables(
        query,
        include_rows=True,
        max_output_rows=max_output_rows,
    )
    return analyze_workbook_bytes(
        stream,
        filename,
        request,
        control=AnalysisControl(timeout_seconds=timeout_seconds),
    )


def serialize_result(value: Any) -> str:
    """Serialize a public result with the library's typed JSON encoding."""

    return to_json(value, indent=2)


def main() -> None:
    """Run a representative subset of the recipes against checked-in workbooks."""

    native_path = WORKBOOKS / "native_table.xlsx"
    named_path = WORKBOOKS / "named_and_headerless.xlsx"
    scattered_path = WORKBOOKS / "scattered_headers.xlsx"
    matrix_path = WORKBOOKS / "sectioned_matrix.xlsx"

    inventory = inspect_structure(native_path)
    native = read_native_table(native_path, "OrdersTable", value_mode="both")
    named = read_defined_name(named_path, "InventoryData")
    headerless = read_known_range(named_path, "Raw Import", "C5:F8", header=None)
    sparse = read_sparse_sheet(named_path, "Raw Import")

    query = TableQuery(
        required_headers=("account number", "amount"),
        optional_headers=("invoice date", "owner"),
        aliases={"account number": ("customer id", "client no")},
        sheet="Scattered Orders",
        within="A4:G20",
        body=BodyPolicy.until_blank_rows(2),
    )
    candidates = find_matching_tables(
        scattered_path,
        ("customer id", "amount"),
        sheet="Scattered Orders",
    )
    discovered = extract_matching_table(scattered_path, query)
    report = explain_matching_tables(scattered_path, query)

    matrix_query = MatrixQuery(
        sections=("Country Identifier", "Sector Identifier"),
        header_level_names=("group", "attribute"),
        sheet="Sectioned Matrix",
    )
    matrix = extract_matrix_section(matrix_path, matrix_query, "Country Identifier")

    trusted_response = analyze_trusted_file(scattered_path, query, max_output_rows=2)
    with scattered_path.open("rb") as stream:
        upload_response = analyze_uploaded_stream(
            stream,
            "customer-upload.xlsx",
            query,
            max_output_rows=2,
        )

    summaries: Mapping[str, Any] = {
        "sheets": [sheet.name for sheet in inventory.sheets],
        "native_rows": len(native.rows),
        "named_rows": len(named.rows),
        "headerless_columns": [column.name for column in headerless.columns],
        "sparse_cells": len(sparse.cells),
        "header_candidates": len(candidates.matches),
        "discovered_rows": len(discovered.rows),
        "explain_candidates": len(report.candidates),
        "matrix_long_rows": len(matrix.long_records()),
        "matrix_wide_rows": len(matrix.wide_records()),
        "trusted_status": trusted_response.status,
        "upload_status": upload_response.status,
    }
    print(serialize_result(summaries))


if __name__ == "__main__":
    main()
