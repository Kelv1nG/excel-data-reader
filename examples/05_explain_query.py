"""Explain how discovery evaluated and selected a table."""

from pathlib import Path

from excel_data_reader import ExcelReader, TableQuery

WORKBOOK = Path(__file__).parent / "workbooks" / "scattered_headers.xlsx"


def main() -> None:
    query = TableQuery(
        required_headers=("customer id", "amount"),
        optional_headers=("invoice date", "owner"),
        sheet="Scattered Orders",
        within="A1:G20",
        near="A4",
    )

    with ExcelReader.open(WORKBOOK) as reader:
        report = reader.explain(query)

    for scan in report.scans:
        bounds = "none" if scan.bounds is None else scan.bounds.a1
        print("Scan:", scan.sheet, bounds, scan.cells_considered, scan.completed)
    for candidate in report.candidates:
        evidence = {
            item.requested_header: [coordinate.a1 for coordinate in item.coordinates]
            for item in candidate.evidence
        }
        print(
            "Candidate:",
            candidate.sheet,
            candidate.header_row,
            "selected=" + str(candidate.selected),
            "evidence=" + str(evidence),
            "reasons=" + str([reason.value for reason in candidate.reasons]),
        )


if __name__ == "__main__":
    main()
