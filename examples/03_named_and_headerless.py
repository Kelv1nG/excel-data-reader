"""Read a rectangular defined name and a separate headerless range."""

from pathlib import Path

from excel_data_reader import ExcelReader

WORKBOOK = Path(__file__).parent / "workbooks" / "named_and_headerless.xlsx"


def main() -> None:
    with ExcelReader.open(WORKBOOK) as reader:
        inventory = reader.get_named_range("InventoryData")
        raw = reader.read_range("Raw Import", "C5:F8", header=None)
        sparse = reader.read_sheet("Raw Import")

    print("Named-range records:")
    for record in inventory.records():
        print(dict(record))

    print("Headerless columns:", [column.name for column in raw.columns])
    print("First headerless row:", raw.rows[0].values)
    print("Sparse populated cells:", len(sparse.cells))


if __name__ == "__main__":
    main()
