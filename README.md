# Excel Data Reader

A Python 3.12 library for deterministic table discovery and coordinate-preserving extraction from
`.xlsx`, `.xlsm`, `.xltx`, and `.xltm` workbooks.

The reader distinguishes native Excel Tables, named ranges, explicit ranges, and header-discovered
regions. It reports ambiguity instead of choosing an arbitrary result, supports non-adjacent
columns, and provides a sparse whole-sheet fallback.

## Quick start

```python
from excel_data_reader import ExcelReader

with ExcelReader.open("orders.xlsx") as workbook:
    matches = workbook.find_tables(
        ["customer id", "invoice date", "amount"],
        sheet="Orders",
        allow_non_adjacent_columns=True,
        max_blank_rows=2,
    )
    match = matches.require_one()
    table = workbook.extract(match)

    for record in table.records():
        print(record)
```

The requested header order becomes the logical output order even if the physical worksheet
columns are scattered.

## Other deterministic entry points

```python
with ExcelReader.open("orders.xlsx", value_mode="both") as workbook:
    inventory = workbook.inventory()

    native = workbook.get_table("SalesTable")
    named = workbook.get_named_range("InvoiceData", header=0)
    headerless = workbook.read_range("Raw", "C5:H100", header=None)

    sparse_sheet = workbook.read_sheet("Raw")
    matrix = sparse_sheet.to_matrix()
```

`header=None` creates `column_1`, `column_2`, and so on. Every extracted value still carries its
original worksheet coordinate.

## Development

```powershell
uv sync --all-groups
uv run pytest
uv run ruff check src tests examples
uv run ruff format --check src tests examples
uv run ty check
```

See [`SPEC.md`](SPEC.md) for the behavioral contract.

## Examples

[`examples/`](examples/) contains runnable programs paired with real workbooks demonstrating:

- native Excel Table extraction;
- normalized header discovery across non-adjacent columns;
- blank-row tolerance;
- rectangular defined names;
- headerless ranges and sparse sheet reads.

Start with [`examples/README.md`](examples/README.md), or run:

```powershell
uv run python examples/01_native_table.py
```
