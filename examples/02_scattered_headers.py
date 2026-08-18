"""Find requested headers even when their physical columns are separated."""

from pathlib import Path

from excel_data_reader import ExcelReader

WORKBOOK = Path(__file__).parent / "workbooks" / "scattered_headers.xlsx"


def main() -> None:
    requested = ["amount", "customer id", "invoice date"]
    with ExcelReader.open(WORKBOOK) as reader:
        match = reader.find_tables(
            requested,
            sheet="Scattered Orders",
            allow_non_adjacent_columns=True,
            max_blank_rows=2,
        ).require_one()
        table = reader.extract(match)

    print("Matched worksheet range:", match.range)
    print("Physical columns:", [column.source_column for column in match.columns])
    print("Logical order:", [column.requested_header for column in match.columns])
    for row in table.rows:
        print(row.source_row, row.values)


if __name__ == "__main__":
    main()
