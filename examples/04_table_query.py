"""Use aliases, optional fields, a search window, and a body policy."""

from pathlib import Path

from excel_data_reader import BodyPolicy, ExcelReader, TableQuery

WORKBOOK = Path(__file__).parent / "workbooks" / "scattered_headers.xlsx"


def main() -> None:
    query = TableQuery(
        required_headers=("account number", "amount"),
        optional_headers=("invoice date", "owner", "purchase order"),
        aliases={"account number": ("customer id", "client no")},
        sheet="Scattered Orders",
        within="A4:G20",
        near="A4",
        body=BodyPolicy.until_blank_rows(2),
    )

    with ExcelReader.open(WORKBOOK) as reader:
        match = reader.query_tables(query).require_one()
        table = reader.extract(match)

    print("Matched worksheet range:", match.range)
    print("Logical fields:", [column.requested_header for column in match.columns])
    print("Source headers:", [column.raw_header for column in match.columns])
    for row in table.rows:
        print(row.source_row, row.values)


if __name__ == "__main__":
    main()
