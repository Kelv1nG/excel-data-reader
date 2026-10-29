# Multilevel table header discovery

Status: implementation authorized and delivered in the working tree. The behavior below
incorporates the discussion through 2026-10-29; `SPEC.md` now defines the public contract.
See the [implementation report](../reports/2026-10-29-multilevel-table-headers.md) for
verification results.

## Intent and agreed behavior

Discover ordinary tables with a caller-declared number of consecutive header rows.
Requested labels may identify individual columns or groups. A group selects its descendant
physical columns. Return flat records while retaining typed cell values, physical column
numbers, and the source of each header label.

The agreed examples, with a three-row header beginning at row 1, are:

| Layout | Authored structure | Output columns |
| --- | --- | --- |
| Plain Store | A1:C1 = Sales Report; A2:A3 = col1; B2 = Online; B3 = New; C2 = Store; C3 empty | Sales Report__col1; Sales Report__Online__New; Sales Report__Store |
| Merged Store | As above, with C2:C3 merged | Same names as Plain Store |
| Wide Store | A1:D1 = Sales Report; C2:D2 = Store; C3:D3 empty | Sales Report__Store__col1 and Sales Report__Store__col2 for C and D |
| Mixed leaves | Store spans C:E; C3 = New; D3:E3 empty | Store__New, Store__col2, Store__col3, prefixed by any ancestors |

Every authored label in the declared header band contributes to the full path, including
a top merge that looks like a title. Each merged cell contributes once per physical column.
Empty cells add no label. Two equal labels in distinct authored cells remain two path parts;
deduplication is by merge identity, not by label text.

Synthetic names use `colN`, where N is the physical offset from the immediate containing
merged group's left edge, starting at one. Named columns count toward the offset. Numbering
restarts in each group and each table. An authored single-column label needs no synthetic
suffix, even when lower header cells are empty.

Each repeated complete header block produces a separate match. A complete repetition in
the same physical columns ends the previous table immediately before the next header,
including when no blank row intervenes. Both matches can have identical column names.

## Approaches and proposed choice

1. Extend `TableQuery` with depth and path-aware headers. This keeps one entry point but
   changes the assumptions behind `TableMatch`, one-row evidence, and the service's existing
   serialization contract.
2. Add `MultiLevelTableQuery` and focused result types. This adds public names but isolates
   the new semantics and can reuse coordinate models, normalization, typed extraction,
   body policies, and reader limits. **Recommended planning assumption.**
3. Generalize `MatrixQuery`. Its section anchors, identifiers, and analytical result shape
   do not fit ordinary independent tables; this is not the proposed implementation.

## Proposed public API

```python
query = MultiLevelTableQuery(
    required_headers=("col1", "main_header"),
    header_depth=3,
    sheet="Report",
    within="A1:H200",
    body=BodyPolicy.until_blank_rows(2),
)

with ExcelReader.open("report.xlsx") as reader:
    matches = reader.find_multilevel_tables(query)
    match = matches.require_one()
    data = reader.extract_multilevel_table(match)
    records = data.records()
```

Fields: `required_headers: Sequence[str | tuple[str, ...]]`, `header_depth: int`,
`sheet: str | None = None`, `within: Rectangle | str | None = None`,
`near: Coordinate | str | None = None`, and `body: BodyPolicy = BodyPolicy()`.
Depth must be a positive integer, excluding booleans. Freeze sequences and validate
nonempty selectors/path components using the existing normalization rules.

A string matches an authored label at any level. A tuple is a complete root-to-node path,
for example `("Sales Report", "Online")`; it is not a joined output key or a suffix path.
All requested selectors must match within one coherent header band. An ambiguous selector
produces a located diagnostic; do not silently pair different occurrences. Use a full
path or `within` to resolve it. Synthetic `colN` names are output metadata, not authored
search anchors, so they cannot accidentally compete with an actual header named `col1`.

Selecting a group expands all descendant columns. Projection follows selector order;
descendants follow physical left-to-right order. Repeated selections of a physical column
include it only on its first encounter. The stored path always includes all ancestors,
even when the query uses only a lower-level label.

The initial API uses strings and path tuples. A nested mapping syntax can later compile
to these selectors; it is not needed to implement the agreed behavior.

## Result and naming contract

Use immutable dataclasses, exported from the package root:

| Type | Fields / behavior |
| --- | --- |
| MultiLevelHeaderPart | `label: str`, `raw_header: Any`, `source_coordinate: Coordinate | None`, `source_bounds: Rectangle | None`; `is_synthetic` is true only when no authored source exists |
| MultiLevelColumn | `source_column: int`, `parts: tuple[MultiLevelHeaderPart, ...]`; `path` returns labels and `name` joins them with `__` |
| MultiLevelTableMatch | `sheet`, `header_bounds`, `columns`, `data_start_row`, `data_end_row`, `diagnostics`; `bounds`, `range`, `is_empty`, and `source = MatchSource.HEADER` mirror ordinary table concepts |
| MultiLevelTableMatchSet | `matches`, `diagnostics`, `is_complete`, `candidate_count` (before body rejection or near filtering); `require_one()` rejects incomplete, missing, or ambiguous discovery, including an unresolved `AMBIGUOUS_MULTILEVEL_HEADER` even when one other valid match exists |
| MultiLevelTableData | `match`, `rows: tuple[DataRow, ...]`; exposes `columns`, `values`, and `records()` |

`header_bounds` covers the entire recognized header footprint, including ancestors and
unselected siblings needed to interpret it. `bounds` extends that footprint to the last
data row, or just the header if the body is empty. Empty bodies have
`data_end_row == data_start_row - 1`; do not create inverted rectangles.

Preserve authored spelling in output labels (`str(raw_header)`); normalization is for
matching and collision checks. Synthetic parts have `raw_header=None` and no invented
source coordinate or source bounds. Authored parts point at the actual merge anchor or
single cell. Data cells retain the existing `CellData` representation, including formulas,
dates, errors, hidden flags, and merged placeholders.

Before producing a match, reject output names that collide under existing normalization.
Examples include a real `col2` colliding with synthetic `col2`, or different paths that
flatten to the same `__` key. Report both locations with `DUPLICATE_MULTILEVEL_HEADER`;
do not add silent suffixes or overwrite records. Positional synthesis applies to missing
leaves, not to duplicate authored labels.

## Header geometry and candidate discovery

Build a hierarchy from actual cell/merge spans within each candidate band. A plain label
covers only its physical column. A horizontal merge covers its authored columns; a
vertical merge is one node spanning rows. Lower nodes must nest within their parents'
horizontal spans. Crossing/overlapping groups are malformed, with a diagnostic at the
conflicting span. Do not inherit the matrix API's horizontal fill-forward heuristic.

For a physical column, traverse its covering nodes top to bottom. If the last authored
node spans one column, the path is complete. If it spans multiple columns and has no
more specific descendant covering this column, append a synthetic leaf whose position
is relative to that node's left edge. This also handles mixed named/unnamed leaves and
merges spanning both rows and columns. A wholly unlabeled column outside a named span is
excluded; neither dimensions, styles, nor body values create a header.

Scan eligible consecutive windows of exactly `header_depth` rows. A candidate needs an
authored label on its top row inside its final footprint, all selectors resolved, and
complete merge containment. A label elsewhere on the same worksheet row does not qualify.
Never clip an intersecting merge at the header-band or query boundary to invent a path.
The footprint is the smallest interval containing the selected nodes, their ancestors,
and their descendant spans; inspect all header columns inside it for a full signature.

The required selectors are the caller's declaration that independent root groups belong
to one requested table. For example, A1:B1 = Sales with Code/Amount below and D1:E1 = Costs
with Tax/Total below can be queried by `("Code", "Tax")`: one match covers A:E, projects
columns A and D, and retains the Sales/Costs ancestors. The unlabeled gap C adds no output
column. Do not infer separate tables merely from that gap. Use separate queries/scopes
when the caller intends independent tables. This is a proposed scope rule consistent
with querying scattered columns today.

If several overlapping windows fit, retain the ambiguity rather than guessing which
rows are data. If duplicate selectors in one band leave neighboring table groups
indistinguishable, return `AMBIGUOUS_MULTILEVEL_HEADER` and require tighter scope; do not
manufacture combinations from separate groups. Side-by-side tables with disjoint,
unambiguous scopes remain independent.

## Body boundaries and ordering

First discover complete header candidates, then determine bodies, then apply `near`.
Near filtering must not remove a later header needed to stop an earlier table.

A repeated-header boundary requires the same sheet, exact left/right footprint, a later
nonoverlapping header band, and an equal normalized complete signature. The signature
contains depth and every physical column's full label path, synthetic flags, and
horizontal source spans, including unselected siblings. Ignore absolute header row
numbers and vertical merge-versus-blank differences that produce the same paths. A lone
repeated word, a different signature, or a separate side-by-side table is not this boundary.
Include empty path entries for wholly unlabeled physical columns inside the footprint
so gaps do not change the meaning of "same columns".

The body's first row is the row after the header band. The next repetition is a hard cap;
otherwise apply the existing `BodyPolicy` to the selected physical columns within scope.
Use the current default of two consecutive blank rows and retain shorter internal blank
runs. `last_populated()` obeys the cap. Proposed rule for `through_row(n)`: reject a query
result if its explicit bottom would cross a recognized next header or lie outside scope;
do not silently clip an explicitly requested row.

Return matches in workbook sheet order, then header top row, then left column. `near`
retains minimum Manhattan distance to the complete match rectangle; ties remain matches.
Choosing the first is explicit via the ordered collection after checking completeness
and the absence of unresolved ambiguous-header diagnostics.
No automatic coalescing of repeated headers, including repeated page headings.

## Architecture, limits, diagnostics, and compatibility

Add public models in `multilevel.py` and geometry/discovery in `_multilevel.py`. The
internal engine accepts plain values, rectangles, a value callback, and a checkpoint;
it imports neither openpyxl nor xlrd. `reader.py` supplies those through its current
adapter surface and uses `_cell_data()` for extraction. OOXML and BIFF share the engine.
Existing matrix/header APIs do not need a broad refactor.

Use `max_scan_cells` before scanning and the existing global `max_candidates` budget.
Exactly the candidate budget is complete if the scan finishes; discovering one more
candidate makes it incomplete. The budget applies before near filtering.
Cache merge lookups and reusable band data; do not expand an unbounded merged area into
objects before enforcing limits. Check cancellation while reading, resolving bands,
expanding selections, inferring bodies, and extracting. Limits remain cooperative.

On an incomplete sheet scan, discard that sheet's provisional matches: unseen repeated
headers could change their bodies. Retain completed earlier sheets but set
`is_complete=False`, attach `SCAN_LIMIT_EXCEEDED`, and make `require_one()` reject the
result before counting matches. Never present a partial scan as a complete table.

Add `INVALID_MULTILEVEL_QUERY`, `MALFORMED_MULTILEVEL_HEADER`,
`AMBIGUOUS_MULTILEVEL_HEADER`, `DUPLICATE_MULTILEVEL_HEADER`, and
`MULTILEVEL_BODY_BOUNDARY_CONFLICT`. Reuse `TABLE_NOT_FOUND`, `AMBIGUOUS_TABLE`, range,
sheet, reader-state, and limit codes where applicable. Include sheet and cell/range
locations. Preserve reader-level adapter warnings, including the legacy XLS warning.

Generic `to_jsonable()` already handles new dataclasses and existing typed cells. The
initial delivery is the reader API; CLI and analysis-service request/response additions
are separate work. Native Excel Tables retain their existing authored-boundary API;
this new query performs header discovery only.

## Scope and verification

Include at least one-, two-, three-, and four-row headers; the three demonstration cases;
mixed leaves; subgroup selection; normalization and collisions; overlapping selections;
repeated bands without blank rows; separate column footprints; incomplete scans; both
workbook formats; typed values, coordinates, and cooperative cancellation.

Defer inference of header depth, fuzzy matching, optional selectors/aliases, unmerged
horizontal parent propagation, independent blank-separated blocks sharing one header,
and adjacent wholly unlabeled columns such as the original `col0`. The user explicitly
set blank columns aside. The ignored demo workbook remains a visual aid; tests generate
their own synthetic workbooks, with a separate public example eligible for the checked-in
acceptance manifest.

Update `SPEC.md` before public implementation, keep dependencies unchanged, target CPython
3.12 only, and use uv. Every new package function/method needs a concise behavioral
docstring describing every explicit parameter except self/cls. Run all repository gates
at implementation handoff. Private workbook validation continues to use
`EXCEL_DATA_READER_ACCEPTANCE_MANIFEST`.

The implementation plan is [here](../plans/2026-10-29-multilevel-table-headers.md).

Builder review: the two material questions concerned grouping independent roots and the
location of the top-row evidence. Both are now resolved above as explicit proposed scope
rules. No further user input is required to make the draft implementable.
