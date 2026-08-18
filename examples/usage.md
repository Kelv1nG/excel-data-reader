# Usage guide

Use the most explicit workbook structure available. The reader follows this order of preference:

1. a native Excel Table;
2. a rectangular defined name;
3. a known cell range;
4. exact normalized header discovery;
5. a sparse whole-sheet read when no table boundary is known.

All extracted cells retain their original worksheet coordinates. Excel 97-2003 `.xls` and OOXML
`.xlsx`, `.xlsm`, `.xltx`, and `.xltm` files use the same public API.

## Native Excel Table

Use an authored Excel Table when one exists because its boundaries are explicit:

```python
from excel_data_reader import ExcelReader

with ExcelReader.open("orders.xlsx") as reader:
    table = reader.get_table("OrdersTable")
    records = table.records()
```

`get_table()` requires exactly one native table with that name. Native-table metadata is available
for OOXML workbooks; legacy `.xls` files should use header, named-range, or explicit-range discovery.

## Find a table by headers

Use `TableQuery` for normal discovery. Required headers must appear on one row, but their physical
order and spacing do not have to match the requested order:

```python
from excel_data_reader import ExcelReader, TableQuery

query = TableQuery(
    required_headers=("customer id", "amount"),
    optional_headers=("invoice date", "owner"),
    aliases={"customer id": ("client no", "account number")},
    sheet="Orders",
    allow_non_adjacent_columns=True,
)

with ExcelReader.open("orders.xls") as reader:
    match = reader.query_tables(query).require_one()
    table = reader.extract(match)
```

Header matching is deterministic: Unicode and whitespace are normalized, underscores and hyphens
act like spaces, and matching is case-insensitive. It does not use fuzzy similarity. The extracted
logical column order follows the query even when worksheet columns are scattered.

Use `reader.explain(query)` when you need scan boundaries, candidate evidence, rejection reasons,
or help diagnosing zero or multiple matches. `require_one()` deliberately reports ambiguity rather
than selecting an arbitrary table.

## No headers

When the data has no header row, provide a known rectangular range and set `header=None`:

```python
with ExcelReader.open("raw-import.xls") as reader:
    table = reader.read_range("Raw", "C5:F100", header=None)
```

The generated logical names are `column_1`, `column_2`, and so on. If the workbook provides a
rectangular defined name, use `get_named_range(name, header=None)` instead.

## Read everything without guessing a table

Use a sparse sheet read when neither headers nor a reliable boundary exist:

```python
with ExcelReader.open("raw-import.xlsx") as reader:
    sheet = reader.read_sheet("Raw")
    cells = sheet.cells
    matrix = sheet.to_matrix()
```

`read_sheet()` returns populated cells with coordinates. `to_matrix()` is available when a dense
rectangular representation is more convenient. Worksheet dimensions are never treated as proof
that the entire sheet is one table.

## Uploaded or untrusted workbooks

Use the platform service rather than opening an uploaded file directly:

```python
from excel_data_reader import AnalysisRequest, TableQuery, analyze_workbook_bytes

request = AnalysisRequest.find_tables(
    TableQuery(("customer id", "amount")),
    include_rows=True,
    max_output_rows=500,
)
response = analyze_workbook_bytes(uploaded_stream, original_filename, request)
```

The service applies format, size, archive, and staging policies and returns stable statuses such as
`success`, `no_match`, `ambiguous`, `rejected`, `cancelled`, and `timeout`. It supports both `.xls`
and OOXML uploads by default.

For legacy `.xls`, formula cells expose their stored calculation results. Unsupported BIFF
features produce a `LEGACY_XLS_LIMITED` warning; table values and header-based extraction remain
available.

## Command line

```powershell
excel-data-reader inspect examples/workbooks/legacy_scattered.xls

excel-data-reader find examples/workbooks/legacy_scattered.xls `
  --headers "customer id,invoice date,amount" `
  --sheet "Legacy Orders" `
  --json
```

## Run the examples

From the repository root:

```powershell
uv sync --all-groups
uv run python examples/01_native_table.py
uv run python examples/02_scattered_headers.py
uv run python examples/03_named_and_headerless.py
uv run python examples/04_table_query.py
uv run python examples/05_explain_query.py
uv run python examples/06_platform_upload.py
uv run python examples/07_legacy_xls.py
```

See `README.md` in this directory for the workbook paired with each script.
