# Excel Data Reader Specification

## Purpose

`excel-data-reader` discovers and extracts tabular values from `.xlsx`-family workbooks without
pretending that every populated worksheet is one table. It preserves worksheet coordinates,
reports ambiguous discovery, and keeps `openpyxl` behind its public model.

## Supported discovery sources

The MVP supports:

1. explicit rectangular A1 ranges;
2. native Excel Tables;
3. rectangular workbook- or worksheet-scoped defined names;
4. exact normalized header signatures on one worksheet row.

Native Excel Tables take precedence over an equivalent header-only candidate. Dynamic,
constant-valued, non-rectangular, and multi-destination defined names remain visible in inventory
but are not silently collapsed into one table.

## Header normalization

Header matching applies Unicode NFKC normalization, converts non-breaking spaces to ordinary
spaces, replaces underscores and hyphens with spaces, collapses whitespace, strips the result,
and performs Unicode case folding. It does not use fuzzy similarity.

Requested headers may appear in any physical order. Non-adjacent columns are allowed unless the
caller disables them. The extracted logical column order follows the requested header order.

## Structured table queries

`TableQuery` separates required and optional logical fields. Every required field must match on
one row. Optional fields are appended to the logical projection when present and do not prevent a
match when absent.

Aliases map a declared logical field to additional accepted header spellings. Primary names and
aliases all use the same exact normalization. An alias cannot belong to multiple fields, and an
alias key must refer to a declared required or optional field. These validation rules prevent one
physical cell from satisfying multiple logical fields.

`within` restricts header scanning to a finite A1 rectangle. A native Excel Table is eligible only
when its full authored bounds lie inside that rectangle. Restricting the scan also restricts its
resource-limit calculation. `near` selects matches with the minimum Manhattan distance from one
A1 cell to each match rectangle. Equal-distance matches remain ambiguous.

## Header-discovered body bounds

The header row is followed downward through the selected columns. A row is blank for boundary
purposes only when all selected columns have no stored value. Discovery stops after the configured
number of consecutive blank rows and excludes trailing blank rows. Internal blank rows remain in
the extracted table.

Header-inferred queries expose three explicit boundary policies:

1. stop after a configured run of blank rows across the selected columns;
2. continue through the last populated row in the selected columns;
3. use a caller-supplied bottom row.

Optional columns that are present participate in boundary detection. Native Excel Tables ignore
these policies and retain their authored data boundaries. A caller-supplied bottom row must not
fall outside `within` when that search constraint is present.

Worksheet dimensions are scan bounds and resource-limit inputs, not evidence that the bounded
rectangle is a table.

## Headerless data

Explicit and named ranges may be read with `header=None`. Columns then receive stable synthetic
names (`column_1`, `column_2`, and so on) while retaining their physical column numbers.

## Results and ambiguity

Discovery returns zero or more immutable `TableMatch` objects. `MatchSet.require_one()` raises a
structured diagnostic error when no match or more than one match exists. Extraction returns
ordered columns and source-addressed cells. Mapping records are available only when normalized
column names are unique.

## Formula values

Readers support `formula`, `cached`, and `both` value modes. `both` represents formula cells as a
`FormulaValue(formula, cached)` pair. Cached values are whatever the workbook last stored; this
package does not calculate Excel formulas.

## Generic sheet reads

`read_sheet()` returns a sparse sequence of value-bearing cells, optionally including styled blank
cells. Its bounding rectangle is descriptive only. Callers can request a matrix from that sparse
result without losing original coordinates.

## Resource limits

Whole-sheet scans check the apparent rectangular cell count before iteration. Work exceeding the
configured `max_scan_cells` fails with `SCAN_LIMIT_EXCEEDED`. Callers can use explicit ranges or
raise the limit deliberately.

## Deferred behavior

The library does not implement fuzzy headers, automatic dense-region detection, multi-row
headers, ordinal compression of unrelated columns, joins, or label/value form extraction.
