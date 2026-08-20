"""Read defined, explicit, headerless, and unstructured worksheet data."""

from pathlib import Path

from excel_data_reader import ExcelReader

WORKBOOK = Path(__file__).parent / "workbooks" / "named_and_headerless.xlsx"


def list_defined_names() -> None:
    """Inspect defined names and their destinations before extracting one."""

    with ExcelReader.open(WORKBOOK) as reader:
        inventory = reader.inventory()

    print("Defined names:")
    for item in inventory.named_ranges:
        print(item.name, [destination.a1 for destination in item.destinations])


def read_a_defined_name() -> None:
    """Extract a rectangular defined name using its first row as headers."""

    with ExcelReader.open(WORKBOOK) as reader:
        table = reader.get_named_range("InventoryData")

    print("Defined-name records:")
    for record in table.records():
        print(dict(record))


def read_the_same_cells_as_an_explicit_range() -> None:
    """Read known coordinates when a workbook does not provide a defined name."""

    with ExcelReader.open(WORKBOOK) as reader:
        table = reader.read_range("Inventory", "B4:E8", header=0)

    print("Explicit-range columns:", [column.name for column in table.columns])
    print("Explicit-range first row:", table.rows[0].values)


def read_a_headerless_range() -> None:
    """Generate column names while preserving every source coordinate."""

    with ExcelReader.open(WORKBOOK) as reader:
        table = reader.read_range("Raw Import", "C5:F8", header=None)

    print("Headerless columns:", [column.name for column in table.columns])
    print("First source cell:", table.rows[0].cells[0].address)
    print("First headerless row:", table.rows[0].values)


def read_an_unstructured_sheet() -> None:
    """Read populated cells sparsely when there is no trustworthy table boundary."""

    with ExcelReader.open(WORKBOOK) as reader:
        sheet = reader.read_sheet("Raw Import")

    print("Sparse sheet:")
    print("bounds:", None if sheet.bounds is None else sheet.bounds.a1)
    print("populated cells:", len(sheet.cells))
    print("first five:", [(cell.address, cell.value) for cell in sheet.cells[:5]])


def main() -> None:
    """Run every named, explicit, and headerless-range usage example."""

    list_defined_names()
    read_a_defined_name()
    read_the_same_cells_as_an_explicit_range()
    read_a_headerless_range()
    read_an_unstructured_sheet()


if __name__ == "__main__":
    main()
