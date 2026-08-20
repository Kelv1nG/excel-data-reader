"""Read an Excel 97-2003 XLS workbook through the legacy adapter."""

from pathlib import Path

from excel_data_reader import ExcelReader

WORKBOOK = Path(__file__).parent / "workbooks" / "legacy_scattered.xls"


def main() -> None:
    with ExcelReader.open(WORKBOOK) as reader:
        match = reader.find_tables(
            ["customer id", "invoice date", "amount"],
            sheet="Legacy Orders",
        ).require_one()
        table = reader.extract(match)

        print("format:", reader.workbook_format)
        for diagnostic in reader.diagnostics:
            print("warning:", diagnostic)

    print("range:", match.range)
    print("source columns:", [column.source_column for column in match.columns])
    for record in table.records():
        print(record)


if __name__ == "__main__":
    main()
