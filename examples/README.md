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

