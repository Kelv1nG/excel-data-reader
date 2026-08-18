# Examples

These examples pair small Python programs with real `.xlsx` workbooks under `workbooks/`.
The workbooks are checked in so each script runs immediately, and `build_workbooks.py` recreates
all of them with OpenPyXL.

From the repository root:

```powershell
uv sync --all-groups
uv run python examples/01_native_table.py
uv run python examples/02_scattered_headers.py
uv run python examples/03_named_and_headerless.py
uv run python examples/04_table_query.py
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
