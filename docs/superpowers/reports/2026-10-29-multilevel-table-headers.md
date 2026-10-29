# Multilevel table headers: implementation and verification

The working tree implements the approved design through `MultiLevelTableQuery`,
`ExcelReader.find_multilevel_tables()`, and `extract_multilevel_table()`. The contract is
in `SPEC.md`; a runnable example is `examples/09_multilevel_tables.py`.

## Delivered behavior

- Explicit header depth with bare authored labels or full root-to-node tuple selectors.
- Group expansion, complete ancestor paths joined with `__`, and source coordinates and
  merge bounds for every authored header component.
- Physical `colN` offsets for missing leaves, restarting per merged group and table.
- Exact matching, explicit ambiguity and name-collision diagnostics, and validation of
  clipped, overlapping, and crossing merges without expanding reads beyond scope.
- Independent repeated-header matches, body caps before later equivalent headers, and
  complete discovery before applying `near`.
- Typed extraction through the existing OOXML/BIFF adapters; cancellation and incomplete
  scan handling preserve the existing execution-control boundary.
- Seven new manifest acceptance cases, a public synthetic workbook, a standalone
  workbook builder, README documentation, and an executable example.

## Execution decisions

- Work stayed in the existing checkout when git metadata was read-only. Grouped commits
  were added on the new feature branch on 2026-10-30. The ignored `excel_workbooks/`
  demo remains untouched.
- `candidate_count` counts resolved header bands before body rejection and near filtering
  so global limits cannot be bypassed by invalid explicit body bounds.
- Explicit `within` ranges include declared empty header rows beyond apparent worksheet
  dimensions. The full range is checked against the cell-scan budget.
- Header diagnostics have a bounded per-sheet budget. Exceeding it marks the sheet
  incomplete and discards its provisional matches, just like candidate overflow.
- The public workbook builder is `examples/multilevel_workbook.py`, separate from the
  existing builder that imports the unavailable legacy writer dependency.
- The independent review found no blocking correctness findings. Its coverage suggestions
  became tests for cancellation inside header resolution/body inference, huge clipped
  merges, vertical-merge/blank equivalence in repeated headers, and nearest-distance ties.
  An additional failing regression exposed an omitted clipped-ancestor diagnostic;
  geometry validation now precedes the top-row-evidence check, and that test passes.

## Full verification on 2026-10-30

The locked project environment now installs successfully. All required gates passed:

| Check | Observed result |
| --- | --- |
| `uv run pytest` | 215 passed |
| `uv run ruff check src tests examples` | All checks passed |
| `uv run ruff format --check src tests examples` | 50 files already formatted |
| `uv run ty check` | All checks passed |
| `git diff --check` | Passed |

The full suite includes service and CLI tests and all generated BIFF cases.

## Initial offline verification on 2026-10-29

All four required `uv run` commands were attempted with a writable cache and offline
mode. They could not start because locked dependencies were unavailable. No dependency
requirements or lockfile versions were changed to work around this.

| Check | Observed result |
| --- | --- |
| `uv run pytest` | Blocked during dependency synchronization |
| `uv run ruff check src tests examples` | Blocked during dependency synchronization |
| `uv run ruff format --check src tests examples` | Blocked during dependency synchronization |
| `uv run ty check` | Blocked during dependency synchronization |
| Limited component pytest run | **142 passed, 5 skipped** |
| Cached Ruff 0.16.2 lint and format checks | Passed; locked version is 0.16.3 |
| Cached ty 0.0.70 | Five unresolved-import diagnostics for missing `xlwt`/`defusedxml.common`; locked version is 0.0.72 |
| `git diff --check` | Passed |

The component run used CPython 3.12, pytest 9.1.1, openpyxl 3.1.5, and system xlrd 2.0.1
(the lockfile requires xlrd 2.0.2). It exercised the five new multilevel test modules plus
`test_model`, `test_reader`, `test_table_query`, `test_anchor_values`, `test_serialization`,
`test_acceptance`, and `test_docstrings`.

The temporary runner loaded the actual package submodules through a namespace package,
bypassing normal package initialization because the service import needs unavailable
`defusedxml`. It did not substitute dependency implementations. Consequently this result
does **not** verify normal package imports, the full service/CLI suite, or the full
repository suite. Five generated BIFF cases were skipped because `xlwt` was missing;
the checked-in legacy workbook and existing BIFF acceptance case did run.

The dependency limit was resolved on 2026-10-30; the full verification above supersedes
the limited component checks.

No feature changes to service operations, CLI commands, existing single-row queries,
or matrix semantics were needed.
