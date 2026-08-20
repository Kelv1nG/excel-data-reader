# Examples

These examples pair small Python programs with real `.xlsx` and `.xls` workbooks under
`workbooks/`. The workbooks are checked in so each script runs immediately. `build_workbooks.py`
recreates the four original table-oriented fixtures with OpenPyXL and xlwt; the sectioned-matrix
fixture is retained as a checked-in layout example.

For a function-oriented recipe guide covering most public entry points, see [`usage.md`](usage.md).
The runnable [`api_recipes.py`](api_recipes.py) module contains parameterized versions of those
recipes, so they can be copied independently instead of treating each numbered script as the only
way to use that feature.

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

Rebuild all workbooks:

```powershell
uv run python examples/build_workbooks.py
```

## 1. Native Excel Table

Workbook: `workbooks/native_table.xlsx`

`01_native_table.py` shows the most deterministic path:

```python
with ExcelReader.open(path, value_mode="formula") as reader:
    table = reader.get_table("OrdersTable")
```

The workbook contains an actual Excel Table named `OrdersTable`, four data rows, formula-derived
amounts, and a totals row. The reader uses the authored table boundary and excludes the totals row
from `table.rows`.

## 2. Scattered headers

Workbook: `workbooks/scattered_headers.xlsx`

`02_scattered_headers.py` searches for three headers in columns A, D, and G:

```python
match = reader.find_tables(
    ["amount", "customer id", "invoice date"],
    sheet="Scattered Orders",
    allow_non_adjacent_columns=True,
    max_blank_rows=2,
).require_one()
```

The logical output follows the requested order even though the physical columns differ. One blank
row appears inside the data and remains present rather than ending the table.

## 3. Named and headerless ranges

Workbook: `workbooks/named_and_headerless.xlsx`

`03_named_and_headerless.py` demonstrates both:

```python
inventory = reader.get_named_range("InventoryData")
raw = reader.read_range("Raw Import", "C5:F8", header=None)
```

The named range supplies a deterministic rectangular boundary. The raw import has no headers, so
the reader supplies `column_1` through `column_4` while preserving every source coordinate.

## 4. Structured table query

Workbook: `workbooks/scattered_headers.xlsx`

`04_table_query.py` searches the same scattered worksheet using a reusable query:

```python
query = TableQuery(
    required_headers=("account number", "amount"),
    optional_headers=("invoice date", "owner", "purchase order"),
    aliases={"account number": ("customer id", "client no")},
    sheet="Scattered Orders",
    within="A4:G20",
    near="A4",
    body=BodyPolicy.until_blank_rows(2),
)
```

The alias maps the workbook's `Customer ID` header to the logical `account number` field. Present
optional columns are returned after required columns; the absent `purchase order` field does not
reject the match. `within` limits scanning, while `near` resolves repeated matching blocks by
distance and leaves equal-distance ties ambiguous.

## 5. Explainable discovery

Workbook: `workbooks/scattered_headers.xlsx`

`05_explain_query.py` calls `reader.explain(query)` and prints the scan window, number of cells
considered, observed header coordinates, selected candidate, and rejection reasons. The returned
`DiscoveryReport` contains the same `MatchSet` that `query_tables()` would return.

## 6. Platform upload service

Workbook: `workbooks/scattered_headers.xlsx`

`06_platform_upload.py` passes uploaded workbook bytes through `analyze_workbook_bytes()`, applies
a cooperative timeout, requests bounded row extraction, and prints the stable response status,
inspection hash, logical columns, and row count. The same service also accepts an open binary
stream without loading the entire upload into memory.

## 7. Legacy Excel 97-2003 workbook

Workbook: `workbooks/legacy_scattered.xls`

`07_legacy_xls.py` opens a genuine BIFF8 `.xls` file, displays the explicit legacy-compatibility
warning, finds non-adjacent header columns, and extracts source-addressed values through the same
`ExcelReader` API used for OOXML workbooks.

## 8. Sectioned matrix

Workbook: `workbooks/sectioned_matrix.xlsx`

`08_sectioned_matrix.py` discovers country and sector row sections beneath shared grouped headers:

```python
query = MatrixQuery(
    sections=("Country Identifier", "Sector Identifier"),
    header_level_names=("group", "attribute"),
    sheet="Sectioned Matrix",
)

matches = reader.find_matrices(query)
country = reader.extract_matrix(matches.require_section("Country Identifier"))
```

The country label is vertically merged and therefore supplies an authored body boundary. The
sector label is an ordinary cell whose following blank label cells inherit the section until two
blank matrix rows. Parent header merges resolve repeated `attr1`, `attr2`, and `attr3` labels into
unique paths such as `("group2", "attr1")`. The script prints both normalized long records and
collision-checked wide records.

## Reusable API recipes

`api_recipes.py` complements the focused numbered examples with reusable functions for:

- workbook inventory;
- native tables, defined names, explicit ranges, and sparse sheets;
- all-match inspection and exactly-one structured discovery;
- discovery reports;
- long and wide matrix records;
- trusted-path and uploaded-stream service calls; and
- stable typed JSON serialization.

Its `main()` function runs a representative subset and prints one JSON summary, while every
recipe can also be imported on its own.
