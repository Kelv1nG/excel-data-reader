"""Use structured table queries for several common discovery conditions."""

from pathlib import Path

from excel_data_reader import BodyPolicy, ExcelReader, TableQuery

WORKBOOK = Path(__file__).parent / "workbooks" / "scattered_headers.xlsx"


def usual_case() -> None:
    """Combine aliases, optional fields, location hints, and a blank-row body."""

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

    print("Aliases and optional fields:")
    print("logical fields:", [column.name for column in table.columns])
    print("source headers:", [column.raw_header for column in match.columns])
    print("rows:", len(table.rows))


def continue_to_the_last_populated_selected_cell() -> None:
    """Ignore unrelated columns when choosing the final table body row."""

    query = TableQuery(
        required_headers=("customer id", "amount"),
        sheet="Scattered Orders",
        body=BodyPolicy.last_populated(),
    )

    with ExcelReader.open(WORKBOOK) as reader:
        match = reader.query_tables(query).require_one()

    print("Last populated selected row:", match.range)


def use_a_known_bottom_row() -> None:
    """Keep rows through a caller-supplied one-based worksheet row."""

    query = TableQuery(
        required_headers=("customer id", "amount"),
        sheet="Scattered Orders",
        within="A4:G8",
        body=BodyPolicy.through_row(8),
    )

    with ExcelReader.open(WORKBOOK) as reader:
        match = reader.query_tables(query).require_one()
        table = reader.extract(match)

    print("Explicit bottom row:", match.range, "rows:", len(table.rows))


def require_adjacent_columns() -> None:
    """Return no match when the query forbids gaps between selected columns."""

    query = TableQuery(
        required_headers=("customer id", "amount"),
        sheet="Scattered Orders",
        allow_non_adjacent_columns=False,
    )

    with ExcelReader.open(WORKBOOK) as reader:
        matches = reader.query_tables(query)

    print("Adjacent-only query matches:", len(matches.matches))


def main() -> None:
    """Run every structured-query usage example."""

    usual_case()
    continue_to_the_last_populated_selected_cell()
    use_a_known_bottom_row()
    require_adjacent_columns()


if __name__ == "__main__":
    main()
