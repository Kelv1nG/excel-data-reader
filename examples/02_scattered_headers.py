"""Discover headers across separated columns under different constraints."""

from pathlib import Path

from excel_data_reader import ExcelReader

WORKBOOK = Path(__file__).parent / "workbooks" / "scattered_headers.xlsx"


def usual_case() -> None:
    """Find scattered headers and extract them in requested logical order."""

    with ExcelReader.open(WORKBOOK) as reader:
        match = reader.find_tables(
            ("amount", "customer id", "invoice date"),
            sheet="Scattered Orders",
            allow_non_adjacent_columns=True,
            max_blank_rows=2,
        ).require_one()
        table = reader.extract(match)

    print("Usual scattered-header discovery:")
    print("range:", match.range)
    print("physical columns:", [column.source_column for column in match.columns])
    print("first logical row:", table.rows[0].values)


def request_a_different_column_order() -> None:
    """Change output order without changing the workbook's physical columns."""

    with ExcelReader.open(WORKBOOK) as reader:
        match = reader.find_tables(
            ("invoice date", "amount", "customer id"),
            sheet="Scattered Orders",
        ).require_one()
        table = reader.extract(match)

    print("Different logical order:")
    print([column.name for column in table.columns])
    print(table.rows[0].values)


def stop_at_the_first_blank_row() -> None:
    """Use a one-row blank tolerance when later data is a separate island."""

    with ExcelReader.open(WORKBOOK) as reader:
        match = reader.find_tables(
            ("customer id", "amount"),
            sheet="Scattered Orders",
            max_blank_rows=1,
        ).require_one()
        table = reader.extract(match)

    print("Stop at first blank row:", match.range, "rows:", len(table.rows))


def reject_non_adjacent_columns() -> None:
    """Show that an adjacency requirement can deliberately reject this layout."""

    with ExcelReader.open(WORKBOOK) as reader:
        matches = reader.find_tables(
            ("customer id", "amount"),
            sheet="Scattered Orders",
            allow_non_adjacent_columns=False,
        )

    print("Require adjacent columns:")
    print("matches:", len(matches.matches))
    print("diagnostics:", [diagnostic.code for diagnostic in matches.diagnostics])


def main() -> None:
    """Run every scattered-header usage example."""

    usual_case()
    request_a_different_column_order()
    stop_at_the_first_blank_row()
    reject_non_adjacent_columns()


if __name__ == "__main__":
    main()
