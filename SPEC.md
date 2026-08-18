# Excel Data Reader Specification

## Purpose

`excel-data-reader` discovers and extracts tabular values from Excel 97-2003 `.xls` and
`.xlsx`-family workbooks without pretending that every populated worksheet is one table. It
preserves worksheet coordinates, reports ambiguous discovery, and keeps format adapters behind
its public model.

## Supported discovery sources

The MVP supports:

1. explicit rectangular A1 ranges;
2. native Excel Tables;
3. rectangular workbook- or worksheet-scoped defined names;
4. exact normalized header signatures on one worksheet row.

Native Excel Tables take precedence over an equivalent header-only candidate. Dynamic,
constant-valued, non-rectangular, and multi-destination defined names remain visible in inventory
but are not silently collapsed into one table.

Native Excel Tables are OOXML-only in this implementation. The legacy adapter can expose
resolvable rectangular `.xls` defined names but does not synthesize native-table metadata.

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

## Discovery reports

`ExcelReader.explain(query)` executes the same discovery path as `query_tables(query)` and returns
a `DiscoveryReport`. The report includes:

- the exact worksheet rectangles considered and their cell counts;
- whether each scan completed or was stopped by a resource limit;
- native-table and interesting header-row candidates;
- required and optional header evidence with physical coordinates and raw labels;
- candidate match counts, proximity distances, selection state, and stable rejection reasons;
- the final `MatchSet` and its diagnostics.

An interesting header row contains at least one primary header or alias from the query. Candidate
reporting is bounded by `max_candidates`; discovery itself retains its existing match limit and
scan-cell limit. An incomplete report is marked on its corresponding `SheetScan`.

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

## Platform analysis contract

`analyze_workbook()` and `analyze_workbook_bytes()` expose analysis schema `1.1`. Each
`AnalysisResponse` includes the typed-value schema version, request identifier, sanitized source
name, operation, status, diagnostics, and any inventory, discovery, extracted-table, or archive
inspection result produced before completion. Host paths are never part of the response.

The inventory operation returns workbook-authored structure. The table-finding operation uses the
same `TableQuery` and explainable discovery semantics as `ExcelReader.explain()`. Row extraction is
opt-in and bounded by `max_output_rows`; ambiguous results are never extracted implicitly.

Stable statuses are `success`, `no_match`, `ambiguous`, `rejected`, `error`, `cancelled`, and
`timeout`. Expected policy, workbook, and reader failures are represented as response diagnostics
instead of crossing the service boundary as library-specific exceptions. Invalid request objects
and invalid function argument types remain programmer errors.

## Untrusted workbook policy

Before a format adapter parses a service input, `WorkbookPolicy` validates the allowed extension,
file size, and format signature. OOXML inputs additionally validate the ZIP entry count,
individual and aggregate expanded size, compression ratio, member-name safety, supported
compression methods, encryption flags, and required package members. OOXML macros and external
links are rejected by default. `defusedxml` is a runtime dependency for XML entity-expansion
protection.

`analyze_workbook_bytes()` accepts bytes-like values or a binary stream. It copies at most the
configured file-size limit into a generated temporary directory, uses only the sanitized basename
as response metadata, and removes the staged file and directory on every handled outcome.

These checks are defense in depth, not antivirus or a complete sandbox. Production platforms that
need hard memory, CPU, or wall-clock isolation must execute analysis in a resource-limited worker
process.

## Cooperative execution control

`AnalysisControl` can provide a timeout and/or cancellation callback. A fresh budget begins for
each service call. Checkpoints run during upload copying, archive inspection and hashing, workbook
opening boundaries, worksheet and native-table iteration, header-row scanning, body inference,
and row extraction. Cancellation takes precedence when both signals are observed.

Timeouts are cooperative and therefore cannot interrupt one blocking format-adapter parse call. A
cancelled or expired operation returns the corresponding stable status and diagnostic, while
context-managed workbook handles and staged uploads are cleaned up.

## Legacy XLS behavior

Excel 97-2003 `.xls` files are read through `xlrd` and copied into the internal worksheet surface
used by discovery and extraction. The adapter retains sheet names and visibility, populated cell
coordinates, scalar values, dates, booleans, error values, number formats, merged rectangles,
hidden rows and columns, and resolvable rectangular defined names.

The adapter exposes formula cells as stored calculation results because formula source text is not
available. It does not expose native Excel Tables, macros, embedded objects, filters, comments,
hyperlinks, pivot tables, conditional formatting, or data validation. Every successfully opened
legacy workbook therefore carries a `LEGACY_XLS_LIMITED` warning. Password-protected or malformed
compound documents are rejected with stable diagnostics.

`WorkbookInspection.format` distinguishes `ooxml` from `xls`. Archive-size and compression fields,
plus macro and external-link detection, are `null` for `.xls` because those checks are specific to
OOXML ZIP packages. `.xlsb` is not supported.

## Acceptance corpus

Workbook acceptance cases are described by a versioned JSON manifest. Every case asserts the
discovery source, physical range, logical projection, row count, and representative source cells.
The checked-in corpus is a non-sensitive seed. Additional local manifests can point to private
production workbooks without making those files part of the package or repository.

## Deferred behavior

The library does not implement fuzzy headers, automatic dense-region detection, multi-row
headers, ordinal compression of unrelated columns, joins, or label/value form extraction.
