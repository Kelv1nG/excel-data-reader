from __future__ import annotations

from datetime import datetime
from pathlib import Path

from excel_data_reader import (
    ANALYSIS_SCHEMA_VERSION,
    AnalysisRequest,
    AnalysisStatus,
    DiagnosticCode,
    ExcelReader,
    Severity,
    TableQuery,
    WorkbookFormat,
    analyze_workbook,
    analyze_workbook_bytes,
)

ROOT = Path(__file__).parents[1]
WORKBOOK = ROOT / "examples" / "workbooks" / "legacy_scattered.xls"


def _query() -> TableQuery:
    return TableQuery(
        ("amount", "customer id", "invoice date"),
        sheet="Legacy Orders",
    )


def test_xls_reader_finds_scattered_headers_and_preserves_coordinates() -> None:
    with ExcelReader.open(WORKBOOK) as reader:
        match = reader.query_tables(_query()).require_one()
        table = reader.extract(match)

        assert reader.workbook_format is WorkbookFormat.LEGACY_XLS
        assert reader.diagnostics[0].code is DiagnosticCode.LEGACY_XLS_LIMITED
        assert reader.diagnostics[0].severity is Severity.WARNING

    assert match.range == "A4:G8"
    assert [column.source_column for column in match.columns] == [7, 1, 4]
    assert table.rows[0].values == (1250.0, "L-001", datetime(2026, 5, 1))
    assert table.rows[2].values == (None, None, None)
    assert table.rows[0].cells[2].address == "D5"
    assert table.rows[0].cells[2].is_date is True
    assert table.rows[0].cells[2].hidden_column is True


def test_xls_reader_supports_inventory_explicit_ranges_and_sparse_reads() -> None:
    with ExcelReader.open(WORKBOOK) as reader:
        inventory = reader.inventory()
        explicit = reader.read_range("Legacy Orders", "A4:G5", header=0)
        sparse = reader.read_sheet("Legacy Orders")

    assert inventory.sheets[0].name == "Legacy Orders"
    assert inventory.sheets[0].apparent_bounds is not None
    assert inventory.sheets[0].apparent_bounds.a1 == "A1:G8"
    assert not inventory.native_tables
    assert explicit.rows[0].cells[0].value == "L-001"
    assert any(cell.address == "G8" and cell.value == 2190.0 for cell in sparse.cells)


def test_xls_path_analysis_uses_the_versioned_service_contract() -> None:
    response = analyze_workbook(
        WORKBOOK,
        AnalysisRequest.find_tables(_query(), include_rows=True),
    )

    assert response.schema_version == ANALYSIS_SCHEMA_VERSION == "1.1"
    assert response.status is AnalysisStatus.SUCCESS
    assert response.inspection is not None
    assert response.inspection.format is WorkbookFormat.LEGACY_XLS
    assert response.tables[0].total_row_count == 4
    assert response.diagnostics[0].code is DiagnosticCode.LEGACY_XLS_LIMITED


def test_xls_uploaded_bytes_are_supported_and_cleaned_up(tmp_path: Path) -> None:
    response = analyze_workbook_bytes(
        WORKBOOK.read_bytes(),
        "../../legacy-upload.xls",
        AnalysisRequest.find_tables(_query(), include_rows=True, max_output_rows=2),
        temp_dir=tmp_path,
    )

    assert response.status is AnalysisStatus.SUCCESS
    assert response.source_name == "legacy-upload.xls"
    assert response.tables[0].truncated is True
    assert len(response.tables[0].rows) == 2
    assert list(tmp_path.iterdir()) == []


def test_invalid_compound_document_is_rejected_by_the_xls_parser(tmp_path: Path) -> None:
    path = tmp_path / "invalid.xls"
    path.write_bytes(bytes.fromhex("D0CF11E0A1B11AE1") + b"not a real compound document")

    response = analyze_workbook(path, AnalysisRequest.inventory())

    assert response.status is AnalysisStatus.REJECTED
    assert response.diagnostics[0].code is DiagnosticCode.INVALID_LEGACY_WORKBOOK
