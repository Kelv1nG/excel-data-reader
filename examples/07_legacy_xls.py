"""Use the same reader API across several legacy XLS workflows."""

from pathlib import Path

from excel_data_reader import ExcelReader

WORKBOOK = Path(__file__).parent / "workbooks" / "legacy_scattered.xls"


def inspect_legacy_workbook() -> None:
    """Inspect format-specific warnings and workbook structure."""

    with ExcelReader.open(WORKBOOK) as reader:
        inventory = reader.inventory()
        print("format:", reader.workbook_format)
        print("sheets:", [sheet.name for sheet in inventory.sheets])
        print("native tables:", len(inventory.native_tables))
        for diagnostic in reader.diagnostics:
            print("warning:", diagnostic.code, diagnostic.message)


def usual_case() -> None:
    """Discover scattered headers and extract records from a BIFF8 workbook."""

    with ExcelReader.open(WORKBOOK) as reader:
        match = reader.find_tables(
            ("customer id", "invoice date", "amount"),
            sheet="Legacy Orders",
        ).require_one()
        table = reader.extract(match)

    print("Legacy header discovery:")
    print("range:", match.range)
    print("source columns:", [column.source_column for column in match.columns])
    print("first record:", dict(table.records()[0]))


def read_a_known_legacy_range() -> None:
    """Bypass discovery when the legacy workbook coordinates are known."""

    with ExcelReader.open(WORKBOOK) as reader:
        table = reader.read_range("Legacy Orders", "A4:G8", header=0)

    print("Explicit legacy range:")
    print("columns:", [column.name for column in table.columns])
    print("rows:", len(table.rows))


def read_a_legacy_sheet_sparsely() -> None:
    """Read populated legacy cells without inferring a table boundary."""

    with ExcelReader.open(WORKBOOK) as reader:
        sheet = reader.read_sheet("Legacy Orders")

    print("Sparse legacy sheet:")
    print("bounds:", None if sheet.bounds is None else sheet.bounds.a1)
    print("populated cells:", len(sheet.cells))


def main() -> None:
    """Run every legacy-XLS usage example."""

    inspect_legacy_workbook()
    usual_case()
    read_a_known_legacy_range()
    read_a_legacy_sheet_sparsely()


if __name__ == "__main__":
    main()
