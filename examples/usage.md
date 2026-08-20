# Usage recipes

Use the most explicit workbook structure available. These recipes are intentionally independent:
copy the function that matches the workbook you have and adapt its arguments.

| Workbook shape | Preferred entry point |
| --- | --- |
| Authored Excel Table | `get_table()` |
| Rectangular defined name | `get_named_range()` |
| Known cell rectangle | `read_range()` |
| Headers are known, location is not | `find_tables()` or `query_tables()` |
| Repeated grouped column headers | `find_matrices()` and `extract_matrix()` |
| No trustworthy boundary | `read_sheet()` |
| Uploaded or otherwise untrusted file | `analyze_workbook_bytes()` |

All extracted cells retain their worksheet coordinates. Excel 97-2003 `.xls` and OOXML `.xlsx`,
`.xlsm`, `.xltx`, and `.xltm` files use the same API except that native Excel Tables are an OOXML
feature.

The runnable [`api_recipes.py`](api_recipes.py) module wraps the common paths below in reusable
functions and runs them against the checked-in example workbooks.

## Inspect a workbook before choosing a strategy

Inventory reports authored structure without extracting table bodies:

```python
from pathlib import Path

from excel_data_reader import ExcelReader, WorkbookInventory


def inspect_workbook(path: str | Path) -> WorkbookInventory:
    with ExcelReader.open(path) as reader:
        return reader.inventory()


inventory = inspect_workbook("orders.xlsx")
print([sheet.name for sheet in inventory.sheets])
print([table.name for table in inventory.native_tables])
print([named.name for named in inventory.named_ranges])
```

If you only need native-table candidates, list their matches directly and inspect the source
locations before selecting one:

```python
with ExcelReader.open("orders.xlsx") as reader:
    matches = reader.find_native_tables(sheet="Orders")

for match in matches.matches:
    print(match.name, match.sheet, match.range)
```

## Read workbook-authored structures

Use a native Excel Table when one exists because its boundary is explicit:

```python
from excel_data_reader import ExcelReader, TableData


def read_orders_table(path: str, table_name: str = "OrdersTable") -> TableData:
    with ExcelReader.open(path) as reader:
        return reader.get_table(table_name)


orders = read_orders_table("orders.xlsx")
for record in orders.records():
    print(dict(record))
```

Use a rectangular defined name in the same way. Pass `sheet=` when a worksheet-scoped name or a
multi-destination name needs disambiguation:

```python
def read_defined_inventory(path: str) -> TableData:
    with ExcelReader.open(path) as reader:
        return reader.get_named_range("InventoryData", sheet="Inventory")
```

Use a known rectangle when the workbook has no authored table object. `header=0` means the first
row in the range supplies column names; `header=None` keeps every row and creates stable synthetic
names:

```python
def read_known_report(path: str) -> TableData:
    with ExcelReader.open(path) as reader:
        return reader.read_range("Report", "B4:F100", header=0)


def read_headerless_import(path: str) -> TableData:
    with ExcelReader.open(path) as reader:
        return reader.read_range("Raw Import", "C5:F100", header=None)


raw = read_headerless_import("raw-import.xlsx")
print([column.name for column in raw.columns])  # column_1, column_2, ...
print(raw.rows[0].cells[0].address)  # original worksheet coordinate
```

`get_named_range(name, header=None)` provides the same headerless behavior for a rectangular
defined name.

## Discover a table by headers

The short form is useful when all fields are required:

```python
def find_orders(path: str) -> TableData:
    with ExcelReader.open(path) as reader:
        match = reader.find_tables(
            ("amount", "customer id", "invoice date"),
            sheet="Orders",
            allow_non_adjacent_columns=True,
            max_blank_rows=2,
        ).require_one()
        return reader.extract(match)
```

Requested order becomes logical output order even when the worksheet columns are scattered.
Header matching normalizes case, Unicode, whitespace, underscores, and hyphens, but never uses
fuzzy similarity.

Use `TableQuery` for aliases, optional fields, search bounds, location hints, and reusable body
rules:

```python
from excel_data_reader import BodyPolicy, ExcelReader, TableQuery

ORDERS_QUERY = TableQuery(
    required_headers=("account number", "amount"),
    optional_headers=("invoice date", "owner", "purchase order"),
    aliases={"account number": ("customer id", "client no")},
    sheet="Orders",
    within="A1:M5000",
    near="A20",
    body=BodyPolicy.until_blank_rows(2),
)


def query_orders(path: str) -> TableData:
    with ExcelReader.open(path) as reader:
        match = reader.query_tables(ORDERS_QUERY).require_one()
        return reader.extract(match)
```

Alternative body boundaries are explicit:

```python
through_last_value = TableQuery(
    ("customer id", "amount"),
    body=BodyPolicy.last_populated(),
)

through_known_row = TableQuery(
    ("customer id", "amount"),
    body=BodyPolicy.through_row(500),
)
```

Native Excel Tables always keep their authored boundaries regardless of the query body policy.

## Handle zero, one, or several matches

Use `.require_one()` when anything other than one result is an error. When a caller should decide
between candidates, retain the `MatchSet` instead:

```python
from excel_data_reader import ExcelDataReaderError


def show_candidates(path: str, query: TableQuery) -> None:
    with ExcelReader.open(path) as reader:
        matches = reader.query_tables(query)

    if not matches.matches:
        print("no match", [item.code for item in matches.diagnostics])
        return
    if len(matches.matches) > 1:
        for match in matches.matches:
            print("candidate", match.sheet, match.range)
        return
    print("selected", matches.matches[0].sheet, matches.matches[0].range)


def require_one_candidate(path: str, query: TableQuery):
    with ExcelReader.open(path) as reader:
        return reader.query_tables(query).require_one()


show_candidates("orders.xlsx", ORDERS_QUERY)

try:
    match = require_one_candidate("orders.xlsx", ORDERS_QUERY)
except ExcelDataReaderError as error:
    for diagnostic in error.diagnostics:
        print(diagnostic.code, diagnostic.sheet, diagnostic.address)
```

The library reports ambiguity instead of choosing an arbitrary candidate. Adding `sheet`,
`within`, or `near` to the query is the deterministic way to narrow results; an equal-distance
`near` tie remains ambiguous.

## Explain why discovery did or did not match

`explain()` runs the same query while retaining scan and candidate evidence:

```python
def explain_orders(path: str, query: TableQuery):
    with ExcelReader.open(path) as reader:
        return reader.explain(query)


report = explain_orders("orders.xlsx", ORDERS_QUERY)
for scan in report.scans:
    print(scan.sheet, scan.bounds, scan.cells_considered, scan.completed)
for candidate in report.candidates:
    print(candidate.selected, candidate.evidence, candidate.reasons)
```

This is useful for user-facing diagnostics and for tightening a query without hiding partial or
ambiguous matches.

## Work with values and source coordinates

Choose the result shape that fits the next step:

```python
table.values  # immutable tuples of typed values
table.records()  # mappings keyed by unique logical column names
table.rows[0].source_row  # original one-based worksheet row
table.rows[0].cells[0].value
table.rows[0].cells[0].address
```

For formulas, choose one of the three value modes:

```python
from excel_data_reader import ExcelReader, FormulaValue, ValueMode


def read_formulas_and_cached_values(path: str) -> TableData:
    with ExcelReader.open(path, value_mode=ValueMode.BOTH) as reader:
        return reader.get_table("OrdersTable")


table = read_formulas_and_cached_values("orders.xlsx")
value = table.rows[0].cells[-1].value
if isinstance(value, FormulaValue):
    print(value.formula, value.cached)
```

`formula` returns formula text, `cached` returns Excel's last stored result, and `both` returns a
`FormulaValue` pair. The reader does not calculate formulas.

## Read a sheet without guessing a table

Use a sparse sheet snapshot when neither headers nor a reliable boundary exist:

```python
from excel_data_reader import SheetData


def read_unstructured_sheet(path: str, sheet_name: str) -> SheetData:
    with ExcelReader.open(path) as reader:
        return reader.read_sheet(sheet_name)


sheet = read_unstructured_sheet("raw-import.xlsx", "Raw")
for cell in sheet.cells:
    print(cell.address, cell.value)

dense = sheet.to_matrix(fill=None)
```

`read_sheet(..., include_styled_blanks=True)` can include blank cells that carry styles. Worksheet
dimensions and styles describe the sheet; they are not treated as evidence that it contains one
table.

## Extract a sectioned matrix

Use `MatrixQuery` when row sections share grouped column headers and repeated leaf labels only
become unique together with their parent labels:

```python
from excel_data_reader import MatrixData, MatrixQuery

MATRIX_QUERY = MatrixQuery(
    sections=("Country Identifier", "Sector Identifier"),
    header_level_names=("group", "attribute"),
    sheet="Sectioned Matrix",
)


def read_matrix_section(path: str, section: str) -> MatrixData:
    with ExcelReader.open(path) as reader:
        matches = reader.find_matrices(MATRIX_QUERY)
        match = matches.require_section(section)
        return reader.extract_matrix(match)


country = read_matrix_section("sectioned_matrix.xlsx", "Country Identifier")
country_long = country.long_records()
country_wide = country.wide_records(separator="__")
```

Prefer `long_records()` for analytical stores because its schema stays stable as groups change.
Use `wide_records()` when downstream code needs one row per identifier. Both shapes retain source
metadata.

## Use the platform boundary

For a trusted path, the service API adds validation, stable statuses, bounded output, and a
versioned JSON shape:

```python
from excel_data_reader import AnalysisRequest, AnalysisStatus, analyze_workbook


def analyze_existing_file(path: str, query: TableQuery):
    request = AnalysisRequest.find_tables(
        query,
        include_inventory=True,
        include_rows=True,
        max_output_rows=500,
        request_id="job-42",
    )
    return analyze_workbook(path, request)


response = analyze_existing_file("orders.xlsx", ORDERS_QUERY)
if response.status is AnalysisStatus.SUCCESS:
    print(response.tables[0].total_row_count)
else:
    print([(item.code, item.message) for item in response.diagnostics])

json_text = response.to_json(indent=2)
```

Use `AnalysisRequest.inventory()` when no discovery is needed.

For an upload, pass bytes or an open binary stream. The upload API safely stages and removes the
file while applying the default untrusted-workbook policy:

```python
from excel_data_reader import AnalysisControl, analyze_workbook_bytes


def analyze_upload(stream, filename: str, query: TableQuery):
    request = AnalysisRequest.find_tables(
        query,
        include_rows=True,
        max_output_rows=500,
    )
    return analyze_workbook_bytes(
        stream,
        filename,
        request,
        control=AnalysisControl(timeout_seconds=10),
    )


with open("customer-upload.xlsx", "rb") as stream:
    response = analyze_upload(stream, "customer-upload.xlsx", ORDERS_QUERY)
```

Cancellation and timeouts are cooperative. Use an isolated, resource-limited worker when a host
requires a hard execution limit.

## Legacy `.xls`

The same discovery, explicit-range, named-range, sparse-sheet, and service recipes work for real
Excel 97-2003 files:

```python
with ExcelReader.open("legacy-orders.xls") as reader:
    print(reader.workbook_format)
    for warning in reader.diagnostics:
        print(warning.code, warning.message)
    table = reader.extract(
        reader.find_tables(
            ("customer id", "invoice date", "amount"),
            sheet="Legacy Orders",
        ).require_one()
    )
```

Legacy formula source text and native Excel Table metadata are unavailable, so use stored values
and the other deterministic entry points.

## Command line

```powershell
excel-data-reader inspect examples/workbooks/native_table.xlsx
excel-data-reader inspect examples/workbooks/legacy_scattered.xls --json

excel-data-reader find examples/workbooks/scattered_headers.xlsx `
  --headers "customer id,amount" `
  --optional "invoice date,owner" `
  --alias "customer id=client no|account number" `
  --sheet "Scattered Orders" `
  --within A1:G20 `
  --near A4 `
  --json
```

## Run the examples

From the repository root:

```powershell
uv sync --all-groups
uv run python examples/api_recipes.py
uv run python examples/01_native_table.py
uv run python examples/02_scattered_headers.py
uv run python examples/03_named_and_headerless.py
uv run python examples/04_table_query.py
uv run python examples/05_explain_query.py
uv run python examples/06_platform_upload.py
uv run python examples/07_legacy_xls.py
uv run python examples/08_sectioned_matrix.py
```

See [`README.md`](README.md) for the workbook paired with each focused script.
