"""Use the versioned platform boundary for several upload conditions."""

from pathlib import Path

from excel_data_reader import (
    AnalysisControl,
    AnalysisRequest,
    AnalysisResponse,
    AnalysisStatus,
    TableQuery,
    analyze_workbook_bytes,
)

WORKBOOK = Path(__file__).parent / "workbooks" / "scattered_headers.xlsx"


def _orders_request(*, max_output_rows: int = 100) -> AnalysisRequest:
    """Build the reusable table-discovery request used by upload examples."""

    return AnalysisRequest.find_tables(
        TableQuery(
            required_headers=("customer id", "amount"),
            optional_headers=("invoice date",),
            sheet="Scattered Orders",
        ),
        include_rows=True,
        max_output_rows=max_output_rows,
        request_id="example-upload",
    )


def _print_response(label: str, response: AnalysisResponse) -> None:
    """Print status, diagnostics, and any extracted table summary."""

    print(label, response.status)
    print("diagnostics:", [diagnostic.code for diagnostic in response.diagnostics])
    if response.status is AnalysisStatus.SUCCESS and response.tables:
        table = response.tables[0]
        print("columns:", [column.name for column in table.match.columns])
        print("rows:", table.total_row_count, "truncated:", table.truncated)


def usual_case() -> None:
    """Analyze in-memory upload bytes and bound the returned rows."""

    response = analyze_workbook_bytes(
        WORKBOOK.read_bytes(),
        "customer-upload.xlsx",
        _orders_request(max_output_rows=2),
        control=AnalysisControl(timeout_seconds=5),
    )
    _print_response("Uploaded bytes:", response)


def analyze_an_open_binary_stream() -> None:
    """Stream an upload into staging instead of first loading all bytes."""

    with WORKBOOK.open("rb") as stream:
        response = analyze_workbook_bytes(
            stream,
            "customer-upload.xlsx",
            _orders_request(),
        )
    _print_response("Uploaded stream:", response)


def request_inventory_only() -> None:
    """Inspect workbook-authored structure without running table discovery."""

    response = analyze_workbook_bytes(
        WORKBOOK.read_bytes(),
        "customer-upload.xlsx",
        AnalysisRequest.inventory(request_id="inventory-example"),
    )
    print("Inventory only:", response.status)
    if response.inventory is not None:
        print("sheets:", [sheet.name for sheet in response.inventory.sheets])


def handle_no_match_as_a_status() -> None:
    """Receive a stable no-match response instead of a reader exception."""

    request = AnalysisRequest.find_tables(
        TableQuery(("missing header",), sheet="Scattered Orders"),
        include_rows=True,
    )
    response = analyze_workbook_bytes(
        WORKBOOK.read_bytes(),
        "customer-upload.xlsx",
        request,
    )
    _print_response("No matching table:", response)


def main() -> None:
    """Run every platform-upload usage example."""

    usual_case()
    analyze_an_open_binary_stream()
    request_inventory_only()
    handle_no_match_as_a_status()


if __name__ == "__main__":
    main()
