# Multilevel Table Headers Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Discover tables with explicitly sized multilevel headers and extract flat,
source-addressed records using the agreed merge and positional naming rules.

**Architecture:** Add a dedicated `MultiLevelTableQuery` and immutable result models.
Keep merge geometry and discovery in an adapter-neutral engine, with `ExcelReader`
handling workbook access and typed extraction. Resolve all header candidates before
assigning bodies so repeated headers provide reliable boundaries.

**Tech Stack:** CPython 3.12, uv, existing openpyxl/xlrd adapters, pytest, ruff, ty.

**Spec:** [Multilevel table header design](../specs/2026-10-29-multilevel-table-headers-design.md).

Status: Tasks 1–5 are implemented in the working tree. On 2026-10-30 the full
locked-environment suite and all required tooling gates passed. See the
[implementation report](../reports/2026-10-29-multilevel-table-headers.md) for evidence
and earlier execution deviations. Unchecked phase-specific command steps below refer to
the original red/green chronology; the final full suite includes those tests.

## Global Constraints

- Target CPython 3.12 only.
- Use uv; introduce no dependencies or lockfile changes for this feature.
- Update `SPEC.md` before changing a public semantic contract.
- Keep openpyxl at the OOXML adapter boundary and xlrd at the legacy BIFF adapter boundary.
- Preserve source coordinates and typed values in all extracted results.
- Keep header normalization exact and deterministic; fuzzy matching is not an MVP behavior.
- Report ambiguity rather than silently selecting one candidate.
- Every package function/method needs a behavioral docstring and descriptions of explicit
  parameters other than `self` and `cls`.
- Keep `excel_workbooks/` ignored. Its demo workbook is not an acceptance fixture or commit input.
- Private workbook checks use `EXCEL_DATA_READER_ACCEPTANCE_MANIFEST`.
- Limits and deadlines remain cooperative; hard limits require an isolated worker.

## Review Focus

1. A generated `col1` must not make a query for an authored `col1` ambiguous (Tasks 2, 4).
2. An explicit depth can fit overlapping row windows: report ambiguity instead of
   silently shifting into data or dropping a title (Tasks 2, 3).
3. Near filtering or a candidate limit must not hide a later header and extend the
   first table through it (Tasks 3, 4).
4. A named child before a blank leaf affects physical numbering; a real `col2` may
   collide with a generated name and must not be overwritten (Tasks 1, 2).
5. A merge clipped by `within`, a crossing group, or a huge merged range must not
   fabricate paths or bypass scan limits (Tasks 2, 4).

## Files and ownership

| File | Responsibility |
| --- | --- |
| `SPEC.md` | Publish the new reader contract before implementation |
| `src/excel_data_reader/multilevel.py` (new) | Public immutable query, header-part, column, match, match-set, and data models |
| `src/excel_data_reader/_multilevel.py` (new) | Merge hierarchy, selector resolution, candidate discovery, signatures, body caps |
| `src/excel_data_reader/reader.py` | Adapter access, new public methods, limits, cross-sheet ordering and near filtering, typed extraction |
| `src/excel_data_reader/diagnostics.py` | New stable diagnostic codes |
| `src/excel_data_reader/__init__.py` | Public exports |
| `tests/test_multilevel_model.py` (new) | Query/result validation and naming collisions |
| `tests/test_multilevel_headers.py` (new) | Geometry and projection using plain value/rectangle fixtures |
| `tests/test_multilevel_discovery.py` (new) | Header windows, independent matches, ordering, body caps and incomplete scans |
| `tests/test_multilevel_reader.py` (new) | Actual OOXML/BIFF workbooks, typed data, cancellation, limits and serialization |
| `examples/multilevel_workbook.py`, new example script and public example workbook | Reproducible synthetic example for documented usage |
| `tests/test_acceptance.py`, `tests/acceptance/manifest.json` | Acceptance support for the new operation |
| `README.md`, `examples/README.md`, `examples/usage.md`, `tests/test_examples.py` | Usage and runnable-example coverage |

The existing `model.py`, service response shapes, matrix implementation, and CLI do not
need structural changes. Generic serialization already accepts dataclasses; add a
contract test before deciding whether any serializer edit is necessary.

## Task 1: Public contract and result models

**Files:** `SPEC.md`, new `multilevel.py`, `diagnostics.py`, `__init__.py`,
new `tests/test_multilevel_model.py`.

**Interfaces:** Produce the six dataclasses and query fields specified in the design.
Public methods/properties: `MultiLevelColumn.path`, `.name`,
`MultiLevelHeaderPart.is_synthetic`, `MultiLevelTableMatch.bounds`, `.range`, `.is_empty`,
`.source`, `MultiLevelTableMatchSet.require_one() -> MultiLevelTableMatch`,
`MultiLevelTableData.columns`, `.values`, and
`.records() -> tuple[Mapping[str, Any], ...]`.

- [x] Add the design's chosen semantics to `SPEC.md`, including raw spelling with `__`,
  synthetic `colN`, per-group physical offsets, incomplete discovery, and repeated-header caps.
- [x] Write model contract tests. Assert `header_depth=0` fails, `True` is not an integer
  depth, selectors cannot normalize to empty text, and input sequences are frozen.
  Construct header parts directly to verify the following expectations:

  ```python
  assert column.path == ("Sales Report", "Store", "col2")
  assert column.name == "Sales Report__Store__col2"
  assert column.parts[-1].source_coordinate is None
  assert column.parts[-1].is_synthetic
  assert empty_match.data_end_row == empty_match.data_start_row - 1
  assert empty_match.bounds == empty_match.header_bounds
  ```

  Test `require_one()` for missing, multiple, and incomplete results; incomplete wins
  over match count. An unresolved ambiguous selector elsewhere also prevents declaring
  one valid match unique. Test records with exact and normalized name collisions, including
  `("A__B", "C")` versus `("A", "B__C")`, and authored/synthetic `col2` collisions.
- [ ] Run `uv run pytest tests/test_multilevel_model.py -q`; expect failure because the
  new exports do not exist. Confirm this is the intended failure.
- [x] Implement the models and the five new diagnostic codes named by the design.
  Make `records()` defend against collisions even for caller-constructed matches.
- [ ] Run `uv run pytest tests/test_multilevel_model.py tests/test_model.py tests/test_docstrings.py -q`;
  expect all passing. Commit only the files owned by this task when execution permits commits.

## Task 2: Resolve one declared header band

**Files:** new `_multilevel.py`, new `tests/test_multilevel_headers.py`.

**Interfaces:** Consume Task 1's models and existing `Coordinate`, `Rectangle`,
`normalize_header`, and `Diagnostic`. Produce internal immutable `_ResolvedHeaderBand`
with `header_bounds`, `all_columns`, and `selected_columns`, plus:

```python
def resolve_band(
    *, sheet: str, header_window: Rectangle,
    merged_ranges: tuple[Rectangle, ...],
    value_at: Callable[[int, int], Any],
    query: MultiLevelTableQuery,
    checkpoint: Callable[[], None],
) -> tuple[_ResolvedHeaderBand | None, tuple[Diagnostic, ...]]: ...
```

`header_window` has exactly the declared depth and the scan's horizontal limits.
The resolved band's footprint may be narrower after ancestor/descendant span closure.
Internal merge indexing and node types are implementer choices.

- [x] Build plain Python sparse value/merge fixtures for Plain Store, Merged Store,
  Wide Store, mixed leaves, a four-row nested example, and two adjacent independent
  groups. Use actual anchor coordinates for all merges. Assert:

  ```python
  assert plain_names == merged_names == (
      "Sales Report__col1", "Sales Report__Online__New", "Sales Report__Store",
  )
  assert wide_names[-2:] == (
      "Sales Report__Store__col1", "Sales Report__Store__col2",
  )
  assert mixed_names[-3:] == (
      "Sales Report__Store__New", "Sales Report__Store__col2",
      "Sales Report__Store__col3",
  )
  ```

  Verify a second group restarts at `col1`; synthetic positions use that group's left
  edge, even when selecting a later subtree. For Plain/Merged Store, the `col1` source
  is A2 and its source bounds A2:A3; top parts all point to A1/A1:C1. The plain Store
  source bounds are C2, while the merged Store bounds are C2:C3.
- [x] Add selector tests: names at any level, full root paths, group expansion, query
  ordering, overlapping selections deduplicated by physical column, and same labels at
  distinct authored nodes retained. Generated `col1` does not compete with authored
  `col1`. Duplicate bare authored selectors produce `AMBIGUOUS_MULTILEVEL_HEADER`.
- [x] Add malformed/collision tests for clipped merges, crossing parent spans,
  authored `col2` plus a missing second leaf, separator collisions, blank top rows,
  wholly unlabeled columns, and unmerged parent labels that must not fill sideways.
- [x] Pin down candidate scope with two tests: selecting Code under A1:B1 = Sales
  and Tax under D1:E1 = Costs yields footprint A1:E2 and selected source columns
  `(1, 4)`, named `Sales__Code` and `Costs__Tax`. Conversely, with A1 empty, A2 = Code,
  and Z1 = Notes, resolving rows 1–2 must reject A1:A2 as a candidate because its own
  top row has no label; Z1 cannot supply that evidence.
- [ ] Run `uv run pytest tests/test_multilevel_headers.py -q`; expect missing engine
  symbols or unsupported resolution failures before implementation.
- [x] Implement `resolve_band()` using indexed merge anchors, horizontal containment,
  source-aware paths, and positioned synthetic leaves. Validate names before returning
  a resolved band. Keep the engine independent of format libraries.
- [ ] Run `uv run pytest tests/test_multilevel_headers.py tests/test_multilevel_model.py -q`;
  expect all passing, then commit this task's files.

## Task 3: Discover candidates and assign independent bodies

**Files:** `_multilevel.py`, new `tests/test_multilevel_discovery.py`.

**Interfaces:** Consume `resolve_band()` and `_ResolvedHeaderBand`; produce:

```python
def discover_sheet(
    *, sheet: str, scan_bounds: Rectangle,
    merged_ranges: tuple[Rectangle, ...],
    value_at: Callable[[int, int], Any],
    query: MultiLevelTableQuery, max_candidates: int,
    checkpoint: Callable[[], None],
) -> MultiLevelTableMatchSet: ...
```

The caller enforces `max_scan_cells` first and passes the remaining global candidate
budget. `discover_sheet()` does not apply `near`; the reader applies that after all body
boundaries are known. Private signature representation is an implementer choice, with
its comparison fields fixed in the design.
The remaining budget may be zero after an earlier sheet; scan to establish whether
another candidate exists, without publishing an excess match.

- [x] Add a fixture with headers at rows 1–3 and 7–9, data at 4–6 and 10–12, with no
  separator rows. Verify:

  ```python
  assert [(m.data_start_row, m.data_end_row) for m in result.matches] == [(4, 6), (10, 12)]
  assert [c.name for c in result.matches[0].columns] == [c.name for c in result.matches[1].columns]
  assert result.matches[0].columns[0].parts[0].source_coordinate.row == 1
  assert result.matches[1].columns[0].parts[0].source_coordinate.row == 7
  ```

- [x] Add tests for a lone matching word in data, a different complete signature,
  unselected sibling changes, a side-by-side header with different physical columns,
  and overlapping fitting windows. None may silently produce the same repeated-header
  cap. Overlapping windows remain ambiguous choices.
- [x] Test two-blank-row stopping, retained internal blanks, `last_populated()` capped
  by the next repetition, empty bodies, and `through_row()` conflicts. Verify trailing
  blank rows are trimmed under inferred policies and explicit bottoms are honored only
  when valid.
- [x] Test incomplete scanning discards the current sheet's provisional matches and
  returns `is_complete=False` with `SCAN_LIMIT_EXCEEDED`. A limit must not hide a second
  header while publishing an overlong first table. Test cancellation during candidate
  resolution and body inference using a callback that raises.
  Exactly N candidates with a limit of N and a completed scan remains complete; N+1
  triggers incomplete status. Blank physical gaps are part of signature comparison.
- [ ] Run `uv run pytest tests/test_multilevel_discovery.py -q`; confirm failures before
  implementation. Implement two-phase discovery, complete signatures, deterministic
  ordering, boundary caps, and limit propagation through `discover_sheet()`.
- [ ] Run `uv run pytest tests/test_multilevel_model.py tests/test_multilevel_headers.py tests/test_multilevel_discovery.py -q`;
  expect all passing, then commit this task's files.

## Task 4: Integrate the reader and both file formats

**Files:** `reader.py`, new `tests/test_multilevel_reader.py`.

**Interfaces:** Consume `discover_sheet()` and public result models. Add:

```python
ExcelReader.find_multilevel_tables(query: MultiLevelTableQuery) -> MultiLevelTableMatchSet
ExcelReader.extract_multilevel_table(match: MultiLevelTableMatch) -> MultiLevelTableData
```

Use existing range/coordinate parsing, workbook order, `_merged_rectangles()`,
`_enforce_scan_limit()`, `_checkpoint()`, and `_cell_data()` at the reader boundary.
Avoid reusing matrix fill-forward or adding format imports to the core engine.

- [x] Generate OOXML fixtures with openpyxl and corresponding BIFF fixtures with the
  already available test dependency xlwt. Test public query/extraction with Plain,
  Merged, Wide, mixed, and four-level layouts. Verify `table.records()` keys and values,
  exact source addresses, and header anchor/bounds metadata.
- [x] Assert original behavior for dates, booleans, errors, formula/cached modes,
  hidden rows/columns, and body merged placeholders. Verify XLS retains
  `LEGACY_XLS_LIMITED`; do not promise unavailable BIFF formula text.
- [x] Add workbook-order/row/column ordering tests, nearest ties, and a near selection
  of table 1 whose body still stops before table 2. Check unknown sheet/closed reader,
  out-of-scope merges, oversized scans, and global candidate limits across sheets.
- [x] Test serialized new results through existing `to_jsonable()`: typed values and
  coordinates survive, synthetic source fields are null, and existing service response
  fixtures remain unchanged. No service operation is added in this task.
- [ ] Run `uv run pytest tests/test_multilevel_reader.py -q`; confirm the new methods
  are missing before implementation. Add reader methods, pass adapter values/rectangles
  to the engine, and extract selected physical columns with `_cell_data()`.
- [ ] Run `uv run pytest tests/test_multilevel_reader.py tests/test_reader.py tests/test_matrix_query.py tests/test_row_blocks.py tests/test_table_query.py tests/test_serialization.py tests/test_docstrings.py -q`;
  expect all passing, then commit this task's files.

## Task 5: Examples, acceptance, and final checks

**Files:** `examples/multilevel_workbook.py`, new `examples/09_multilevel_tables.py`, new
`examples/workbooks/multilevel_tables.xlsx`, `tests/test_acceptance.py`,
`tests/acceptance/manifest.json`, `tests/test_examples.py`, `README.md`,
`examples/README.md`, `examples/usage.md`.

**Interfaces:** Public reader API from Task 4. Add acceptance operation `multilevel`
with query fields matching the API and optional zero-based `match_index` for deliberate
selection of repeated results. Expected fields: `match_count`, `header_range`, `range`,
`source_columns`, `names`, `row_count`, and selected typed cell values by address.

- [x] Build a separate, entirely synthetic public workbook with plain, merged, wide,
  mixed-leaf, four-level, and repeated-without-blanks sheets. Generate it reproducibly
  through `examples/multilevel_workbook.py`. Do not move or commit the ignored demo file.
- [x] Extend the acceptance harness for the new result type without routing it through
  existing `TableMatch` assertions. Add expectations for the key names, merge anchors,
  selected columns, repeated-header body ends, and independent match counts. Ensure a
  JSON selector array is converted to a tuple path, not a string.
- [x] Add the executable example with one query by two top-level labels, one subgroup
  selection, and explicit extraction of two repeated matches. Document source spelling,
  complete paths, synthetic numbering, ambiguity, and repeated-header stopping.
- [ ] Run `uv run pytest tests/test_acceptance.py tests/test_examples.py -q`; expect the
  new cases to pass alongside all existing manifest cases. Correct actual failures;
  do not weaken established workbook expectations.
- [x] Run every required gate and inspect exit codes:

  ```text
  uv run pytest
  uv run ruff check src tests examples
  uv run ruff format --check src tests examples
  uv run ty check
  git diff --check
  ```

  Success means all exit zero, no failing tests, no lint/format/type errors. If tools
  cannot start, report the environment limitation rather than marking the gate passed.
- [x] Review changes against the design, confirm no private/demo artifacts are tracked,
  and commit only this task's files. Handoff includes the API example and verification.

## Execution notes

Tasks 1–4 share interfaces and should run in order. Inline execution is the recommended
starting point; an independent review after integration can check the whole feature.
Use worktree guidance at implementation time if isolation is needed. This planning pass
does not create a branch, install packages, change product code, or commit the user's work.

During the demo session, the default uv cache was read-only and dependencies were
unavailable offline. A writable cache can be selected with `UV_CACHE_DIR`; this does not
solve missing packages. Recheck the execution environment before tests and use the locked
versions rather than altering dependencies to work around it.
