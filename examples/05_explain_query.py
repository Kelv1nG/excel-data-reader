"""Explain successful, partial, and constraint-rejected table discovery."""

from pathlib import Path

from excel_data_reader import DiscoveryReport, ExcelReader, TableQuery

WORKBOOK = Path(__file__).parent / "workbooks" / "scattered_headers.xlsx"


def _print_report(label: str, report: DiscoveryReport) -> None:
    """Print a compact summary of one discovery report."""

    print(label)
    print("selected matches:", len(report.selected_matches))
    print("diagnostics:", [diagnostic.code for diagnostic in report.diagnostics])
    for scan in report.scans:
        bounds = "none" if scan.bounds is None else scan.bounds.a1
        print("scan:", scan.sheet, bounds, scan.cells_considered, scan.completed)
    for candidate in report.candidates:
        evidence = {
            item.requested_header: [coordinate.a1 for coordinate in item.coordinates]
            for item in candidate.evidence
        }
        print(
            "candidate:",
            candidate.sheet,
            candidate.header_row,
            "selected=" + str(candidate.selected),
            "evidence=" + str(evidence),
            "reasons=" + str([reason.value for reason in candidate.reasons]),
        )


def usual_case() -> None:
    """Explain a successful query with optional fields and a search window."""

    query = TableQuery(
        required_headers=("customer id", "amount"),
        optional_headers=("invoice date", "owner"),
        sheet="Scattered Orders",
        within="A1:G20",
        near="A4",
    )
    with ExcelReader.open(WORKBOOK) as reader:
        report = reader.explain(query)

    _print_report("Successful query:", report)


def explain_a_missing_required_header() -> None:
    """Show partial evidence when one required header does not exist."""

    query = TableQuery(
        required_headers=("customer id", "amount", "tax code"),
        sheet="Scattered Orders",
        within="A1:G20",
    )
    with ExcelReader.open(WORKBOOK) as reader:
        report = reader.explain(query)

    _print_report("Missing required header:", report)


def explain_an_adjacency_rejection() -> None:
    """Show why otherwise matching headers fail an adjacency constraint."""

    query = TableQuery(
        required_headers=("customer id", "amount"),
        sheet="Scattered Orders",
        allow_non_adjacent_columns=False,
    )
    with ExcelReader.open(WORKBOOK) as reader:
        report = reader.explain(query)

    _print_report("Rejected non-adjacent headers:", report)


def main() -> None:
    """Run every explainable-discovery usage example."""

    usual_case()
    explain_a_missing_required_header()
    explain_an_adjacency_rejection()


if __name__ == "__main__":
    main()
