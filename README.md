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

## Structured queries

Use `TableQuery` when workbook producers use different labels, some fields are optional, or the
same header set appears more than once:

```python
from excel_data_reader import BodyPolicy, ExcelReader, TableQuery

query = TableQuery(
    required_headers=("customer id", "amount"),
    optional_headers=("invoice date", "owner"),
    aliases={
        "customer id": ("client no", "account number"),
        "amount": ("gross value",),
    },
    sheet="Orders",
    within="A1:M5000",
    near="A20",
    body=BodyPolicy.until_blank_rows(2),
)

with ExcelReader.open("orders.xlsx") as workbook:
    match = workbook.query_tables(query).require_one()
    table = workbook.extract(match)
```

All required headers must appear on one row. Optional headers are included when present and are
omitted otherwise. Aliases still use exact normalized matching—there is no fuzzy guessing.
`within` restricts the scan to one A1 rectangle, and `near` retains only the closest match; an
equal-distance tie remains explicit ambiguity.

Header-inferred tables support three body policies:

- `BodyPolicy.until_blank_rows(2)` stops after a consecutive blank-row run across the selected
  columns;
- `BodyPolicy.last_populated()` uses the last populated row in the selected columns;
- `BodyPolicy.through_row(500)` uses a caller-supplied bottom row.

Native Excel Tables always retain their authored boundaries.

## Explainable discovery

Use `explain()` when a platform needs to show why a table matched—or why it did not:

```python
with ExcelReader.open("orders.xlsx") as workbook:
    report = workbook.explain(query)

for scan in report.scans:
    print(scan.sheet, scan.bounds, scan.cells_considered, scan.completed)

for candidate in report.candidates:
    print(candidate.selected, candidate.evidence, candidate.reasons)
```

`DiscoveryReport` includes the normal `MatchSet`, worksheet scan boundaries, header evidence with
source coordinates, partial candidates, proximity distances, and stable rejection reasons. Report
candidate collection is bounded by the reader's `max_candidates` safety limit.

## Command line

The installed package includes an inspection and discovery CLI:

```powershell
excel-data-reader inspect orders.xlsx
excel-data-reader inspect orders.xlsx --json

excel-data-reader find orders.xlsx `
  --headers "customer id,amount" `
  --optional "invoice date,owner" `
  --alias "customer id=client no|account number" `
  --sheet Orders `
  --within A1:M5000 `
  --near A20 `
  --json
```

`find` exits with `0` for one match, `2` for no match or invalid input, and `3` for ambiguity. JSON
output serializes the same public inventory and discovery-report models used by Python callers.

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

When no reliable headers exist, prefer a known range (`read_range(..., header=None)`), a defined
name, or a sparse whole-sheet read. The library deliberately does not guess a table solely from
formatting, empty-space patterns, or worksheet dimensions.

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

[`examples/`](examples/) contains runnable programs paired with generated example workbooks
demonstrating:

- native Excel Table extraction;
- normalized header discovery across non-adjacent columns;
- blank-row tolerance;
- rectangular defined names;
- headerless ranges and sparse sheet reads;
- aliases, optional headers, location hints, and body policies;
- explainable discovery reports and command-line inspection.

Start with [`examples/README.md`](examples/README.md), or run:

```powershell
uv run python examples/01_native_table.py
```

The manifest-driven seed acceptance corpus lives under `tests/acceptance/`. A private production
corpus can be included locally through `EXCEL_DATA_READER_ACCEPTANCE_MANIFEST` without committing
its workbooks.
