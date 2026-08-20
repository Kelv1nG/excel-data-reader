# Usage guide

Use the most explicit workbook structure available. The reader follows this order of preference:

1. a native Excel Table;
2. a rectangular defined name;
3. a known cell range;
4. exact normalized header discovery;
5. sectioned-matrix discovery for shared hierarchical headers;
6. a sparse whole-sheet read when no table boundary is known.

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

## Extract a sectioned matrix

Use `MatrixQuery` when repeated leaf headers are grouped by parent labels and multiple row sections
share that column structure:

```python
from excel_data_reader import ExcelReader, MatrixQuery

query = MatrixQuery(
    sections=("Country Identifier", "Sector Identifier"),
    header_level_names=("group", "attribute"),
    sheet="Sectioned Matrix",
)

with ExcelReader.open("sectioned_matrix.xlsx") as reader:
    matches = reader.find_matrices(query)
    country = reader.extract_matrix(matches.require_section("Country Identifier"))
    sector = reader.extract_matrix(matches.require_section("Sector Identifier"))

country_long = country.long_records()
sector_wide = sector.wide_records()
```

The example workbook contains a vertically merged country anchor and an ordinary, unmerged sector
anchor. A merged anchor supplies its authored row span. An unmerged anchor continues until the next
requested section or the configured `body` policy. Blank section-label cells inherit the active
section only in the extracted result; blank row identifiers remain blank and produce a warning
when their row contains values.

The `MatrixQuery` parameters are:

- `sections`: one or more required logical section labels;
- `header_level_names`: names used in long records, such as `group` and `attribute`;
- `aliases`: alternate exact-normalized labels for declared sections;
- `sheet`: optional exact worksheet name;
- `within`: optional finite A1 search rectangle for disambiguation and scan limits;
- `body`: `BodyPolicy` used when a section label has no vertical merge;
- `header_rows`: optional one-based row numbers when automatic header inference is ambiguous;
- `identifier_column`: optional one-based index or column letters when density inference ties.

Prefer `long_records()` for DuckDB because it keeps a stable schema as groups change. Use
`wide_records(separator="__")` when downstream code expects one row per identifier. Both forms
retain source metadata; the long form retains one source cell for every value.

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
uv run python examples/08_sectioned_matrix.py
```

See `README.md` in this directory for the workbook paired with each script.
