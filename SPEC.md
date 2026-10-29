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

## Anchor-relative scalar values

`ExcelReader.find_values_from_anchor()` discovers text-valued anchor cells and returns the cells at
a caller-supplied signed row and column offset. Anchor text uses the same exact deterministic
normalization as headers; fuzzy matching is not performed. Results retain both the anchor
`CellData` and the offset value `CellData`, including their physical coordinates and typed values.
Numeric, boolean, date, error, and formula cells are not text anchors even when their rendered or
source text would normalize to the query.

`row_offset` follows worksheet row numbering: positive values move down and negative values move
up. `column_offset` follows worksheet column numbering: positive values move right and negative
values move left. Both offsets are integers, booleans are invalid, and they may not both be zero.
The default `(row_offset=0, column_offset=1)` selects the cell immediately to the right.

An optional exact worksheet name limits discovery to one sheet. An optional finite `within`
rectangle limits both anchor discovery and eligible target cells; without it, discovery scans each
selected worksheet's apparent bounds. All scans obey `max_scan_cells`. Targets may be blank, but
must remain inside the supplied rectangle when present and inside the source format's worksheet
grid. Out-of-bounds targets are reported with stable diagnostics rather than clipped.
Reading an absent blank target does not materialize a worksheet cell or change apparent bounds.

Matches are ordered by workbook sheet order and then row-major coordinate order. Repeated anchors
remain explicit: `find_values_from_anchor()` returns every eligible match, while
`get_value_from_anchor()` requires exactly one and reports ambiguity otherwise. Merged-cell values
are never propagated: discovery naturally identifies an authored value at its physical anchor,
and an offset target inside a non-anchor merged coordinate remains a blank physical cell.
Discovery is bounded by the reader's `max_candidates` limit. A truncated or otherwise incomplete
scan carries `SCAN_LIMIT_EXCEEDED`; an exactly-one accessor does not treat an incomplete scan as a
unique result.

## Blank-separated row blocks

`RowBlockQuery` discovers ordered rectangular row blocks that share one explicitly located header
row but have no workbook-authored section labels. This contract is separate from `TableQuery` and
`MatrixQuery`: row-block discovery does not infer semantic section identities and does not search
for headers. The caller supplies an exact worksheet, a finite one-row `header_range`, a finite
`scan_range` below it, and a `BlankRowSeparator` policy. An optional positive `max_blocks` limits
discovery to the first requested number of physical blocks. This prefix limit is independent of
the finite scan range, which remains the spatial and resource-safety boundary.

An optional `column_names` mapping assigns caller-authored logical names by physical Excel column
letters. Keys are case-insensitive column letters and must identify columns inside `scan_range`;
values must be non-empty strings. An explicit name takes precedence over an authored or synthetic
output name, while `ColumnInfo.raw_header`, `header_coordinate`, and `source_column` continue to
preserve the workbook source. `ColumnInfo.requested_header` records the explicit logical name.

The header range's columns must lie inside the scan range's columns. Columns in the scan range but
outside the header range receive stable synthetic names, allowing a layout such as headers in
`B6:D6` and an unnamed row-label column in `A8:A5000`. Blank cells inside the declared header range
also receive synthetic names. Authored, synthetic, and explicitly named columns retain their
physical coordinates. Without an explicit name, the reader never invents a semantic identity for
an unnamed column.

A scan row is populated when at least one cell across the scan range's columns has a stored value.
Styles and worksheet dimensions are not row-block evidence. A run containing at least
`BlankRowSeparator.min_rows` consecutive unpopulated rows separates adjacent blocks. Separator
rows, leading blank rows, and trailing blank rows are excluded. Shorter internal blank runs remain
inside a block and are returned during extraction.

`ExcelReader.find_row_blocks()` returns `RowBlockMatch` objects in physical top-to-bottom order.
Each match has a stable one-based ordinal, shared column metadata, exact data bounds, and the blank
separator following it when another block exists. `RowBlockMatchSet.require_count()` validates an
expected structural count before callers assign positional application names such as tab labels.
The reader never invents semantic names for unlabeled blocks. When `max_blocks` is set, discovery
returns at most that many matches and may stop after the requested final block is closed by the
declared separator. Blocks below that prefix are intentionally neither returned nor counted.
Consequently, `require_count(max_blocks)` means that the requested prefix was available; it does
not assert that the worksheet contains no additional blocks. If fewer blocks exist within the scan
range, they are returned normally and `require_count()` supplies the stable count-mismatch error.

`ExcelReader.extract_row_block()` returns coordinate-preserving `RowBlockData`. Its `values` and
`records()` interfaces mirror ordinary table data, including the requirement that normalized
column names be unique before mapping records are produced. Discovery obeys `max_scan_cells` and
`max_candidates`. A populated final scan row reached during discovery produces a stable warning
because the explicit scan range may have truncated the last block.

## Sectioned matrices

`MatrixQuery` discovers repeated row sections beneath a shared hierarchical column header. This
contract is separate from `TableQuery`: ordinary table discovery continues to require all logical
headers on one row, while a matrix column is identified by an ordered path such as
`("group1", "attr1")` assembled from multiple header levels.

A query declares one or more exact normalized section labels and names for the column-header
levels. Section aliases follow the same deterministic normalization and ownership rules as table
header aliases. Optional `within`, `header_rows`, and `identifier_column` hints can constrain
discovery, but physical columns and rows are inferred when those hints are absent.

Horizontal merged cells are structural evidence for a parent header's column span. When parent
headers are not merged, a non-empty label applies to following value columns until the next label
on that header row. The closest populated row above the first requested section supplies leaf
headers. Every value column must resolve one label and source coordinate for every declared header
level. Repeated leaf labels are valid when their complete header paths differ.

A section label may be vertically merged or stored in one ordinary cell. A vertical merge supplies
an authored row boundary. Otherwise, the label is a semantic anchor: following rows remain in the
section until the next requested section anchor, the configured blank-row boundary, an explicit
bottom row, or the last populated selected row. Section labels are propagated only in the result;
the workbook is never modified. Individual row identifiers are never propagated into blank cells.

When no identifier-column override is supplied, discovery considers columns between the section
anchor and the first value column. It selects the uniquely densest populated candidate across the
section rows. A tie remains explicit ambiguity. Rows containing matrix values without an
identifier produce a stable warning rather than inheriting the preceding identifier.

`ExcelReader.find_matrices()` returns one `MatrixMatch` per discovered section anchor plus stable
missing, ambiguity, and malformed-layout diagnostics. `MatrixMatchSet.require_section()` requires
exactly one match for a logical section. `ExcelReader.extract_matrix()` returns source-addressed
matrix values. When a selected matrix value is a non-anchor coordinate inside a merged range, it
is returned as a blank physical cell with `is_date=False`; the top-left anchor value is not
propagated. `MatrixData.long_records()` produces the normalized, schema-stable representation
recommended for analytical databases; `MatrixData.wide_records()` provides one record per
identifier with collision-checked flattened header paths.

## Multilevel table headers

`MultiLevelTableQuery` discovers ordinary tables in a caller-declared positive integer
`header_depth` of consecutive rows. It uses a dedicated reader API:
`find_multilevel_tables()` and `extract_multilevel_table()`. Ordinary `TableQuery` keeps
its one-row contract; native tables retain their existing authored-boundary API.

Required selectors are authored labels at any level or complete root-to-node tuple
paths. Matching uses existing exact normalization. Every selector must resolve within
one candidate; duplicate authored occurrences require a fuller path or tighter `within`
scope. Synthetic names are output metadata and cannot act as search anchors. Selecting
a group includes its descendant physical columns. Output follows selector order, then
left-to-right descendant order, deduplicating physical columns on their first encounter.
An explicit `within` range defines the full scan area, including empty header rows
beyond apparent worksheet dimensions, and remains subject to the scan-cell limit.

Each output path includes every covering authored label inside the declared header band.
Horizontal merges supply explicit column spans; vertical merges contribute a label only
once. Unmerged labels cover one column and are never filled sideways. Empty header cells
add nothing. Identical text in distinct authored cells remains distinct path components.
Authored labels retain their spelling and are joined with `__` for flat record keys.
For example both plain C2 = Store with empty C3 and merged C2:C3 = Store under Sales
Report produce `Sales Report__Store`.

When the last covering authored node spans several columns, an unlabeled leaf receives
`colN`, using its physical one-based offset within that group. Numbering restarts per
group and table and counts named positions: `New`, blank, blank produces `New`, `col2`,
`col3`. Wholly unlabeled columns outside a named span are excluded. Reject normalized
output-name collisions, including synthetic/authored collisions and flattened-path
collisions, with `DUPLICATE_MULTILEVEL_HEADER` rather than renaming or overwriting.

Candidates require top-row evidence inside their final footprint and complete merge
containment. Clipped, crossing, or overlapping groups are malformed. The footprint
contains selected nodes, ancestors, and all descendants of those ancestors. Required
selectors may deliberately combine independent root groups and scattered columns;
blank gaps alone do not divide them. Ambiguous same-band selector occurrences are
reported without manufacturing combinations. Overlapping fitting row windows remain
separate ambiguous candidates, rather than guessing the correct start row.

Discover header candidates before inferring bodies and before applying `near`. A later
nonoverlapping header band caps the earlier body when its depth, complete normalized
paths, synthetic markers, horizontal source spans, and physical left/right footprint
match. Include unselected siblings and unlabeled gaps in that comparison. Absolute
header row numbers and equivalent vertical merge-versus-blank layouts do not affect it.
The earlier body ends before the later header even without blank separator rows.
A single repeated word, a different signature, or a disjoint column footprint is not
this boundary. Return independent matches in workbook order, then header row and column.

Data starts immediately below the header band. Within the repeated-header cap and query
scope, use the selected columns and existing `BodyPolicy`; the default is two blank rows.
`last_populated()` also obeys the cap. An explicit `through_row()` that crosses it or
lies outside the scan range produces `MULTILEVEL_BODY_BOUNDARY_CONFLICT`.
`near` uses Manhattan distance to complete match rectangles; ties remain ambiguous.

Results preserve typed `CellData`, source columns, and each header label's actual source
coordinate and cell/merge bounds. Synthetic parts have no invented source address.
`MultiLevelTableData` exposes `columns`, `rows`, `values`, and `records()`.

Discovery obeys scan-cell and global candidate limits and checks cancellation throughout.
Exactly the candidate limit is complete if scanning finishes; finding an additional
candidate is incomplete. Discard provisional matches from an incomplete sheet because
unseen repetitions may change their bodies. Keep completed earlier sheets but set
`is_complete=False` and attach `SCAN_LIMIT_EXCEEDED`. `require_one()` refuses incomplete
results and unresolved `AMBIGUOUS_MULTILEVEL_HEADER` diagnostics before counting matches.
`candidate_count` records resolved bands before body-policy rejection or `near` filtering,
so rejected explicit bodies still consume the global candidate budget. On an incomplete
scan this count covers only the bands encountered. Header diagnostics are also bounded
per sheet by the remaining candidate budget (at least one); overflow discards that
sheet's provisional matches and reports `SCAN_LIMIT_EXCEEDED`.

Additional diagnostics are `INVALID_MULTILEVEL_QUERY`, `MALFORMED_MULTILEVEL_HEADER`,
`AMBIGUOUS_MULTILEVEL_HEADER`, and `DUPLICATE_MULTILEVEL_HEADER`, with source locations.
The initial contract supports OOXML and BIFF through their existing adapters. Header
depth inference, nested query mappings, aliases, optional selectors, adjacent wholly
unlabeled columns, shared-header row-block splitting, CLI commands, and service operations
are outside this delivery.

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

The library does not implement fuzzy headers, automatic dense-region detection, arbitrary
multi-row table headers outside the explicit sectioned-matrix contract, ordinal compression of
unrelated columns, joins, or multi-field form extraction beyond anchor-relative scalar lookup.
