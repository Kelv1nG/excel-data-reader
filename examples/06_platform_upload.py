"""Analyze uploaded workbook bytes through the versioned platform boundary."""

from pathlib import Path

from excel_data_reader import (
    AnalysisControl,
    AnalysisRequest,
    AnalysisStatus,
    TableQuery,
    analyze_workbook_bytes,
)

workbook_path = Path(__file__).parent / "workbooks" / "scattered_headers.xlsx"
request = AnalysisRequest.find_tables(
    TableQuery(
        required_headers=("customer id", "amount"),
        optional_headers=("invoice date",),
        sheet="Scattered Orders",
        allow_non_adjacent_columns=True,
    ),
    include_rows=True,
    max_output_rows=100,
    request_id="example-upload",
)

response = analyze_workbook_bytes(
    workbook_path.read_bytes(),
    "customer-upload.xlsx",
    request,
    control=AnalysisControl(timeout_seconds=5),
)

print("status:", response.status)
if response.inspection is not None:
    print("sha256:", response.inspection.sha256)
if response.status is AnalysisStatus.SUCCESS:
    table = response.tables[0]
    print("columns:", [column.name for column in table.match.columns])
    print("rows:", table.total_row_count, "truncated:", table.truncated)
