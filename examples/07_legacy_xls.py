"""Read an Excel 97-2003 XLS workbook through the legacy adapter."""

from pathlib import Path

from excel_data_reader import ExcelReader

workbook_path = Path(__file__).parent / "workbooks" / "legacy_scattered.xls"

with ExcelReader.open(workbook_path) as reader:
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
