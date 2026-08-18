"""Read an authored native Excel Table by its workbook-defined name."""

from pathlib import Path

from excel_data_reader import ExcelReader

WORKBOOK = Path(__file__).parent / "workbooks" / "native_table.xlsx"


def main() -> None:
    with ExcelReader.open(WORKBOOK, value_mode="formula") as reader:
        inventory = reader.inventory()
        table = reader.get_table("OrdersTable")

    print("Native tables:", [item.name for item in inventory.native_tables])
    print("Data range:", table.match.range)
    print("Columns:", [column.name for column in table.columns])
    for record in table.records():
        print(dict(record))


if __name__ == "__main__":
    main()
