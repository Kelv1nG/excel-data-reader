# Repository Instructions

## Sources of truth

- `SPEC.md` defines discovery and extraction semantics.
- `README.md` defines the public usage examples.
- `pyproject.toml` defines dependencies and tooling.
- `tests/acceptance/manifest.json` defines checked-in workbook acceptance expectations.
- Update the specification before changing a public semantic contract.

## Environment

- Target CPython 3.12 only.
- Use `uv` for environments, dependency management, locking, and command execution.
- Keep `openpyxl` at the OOXML adapter boundary and `xlrd` at the legacy BIFF adapter boundary.
- Keep upload validation, execution control, and response serialization in the platform service
  boundary rather than mixing them into discovery semantics.

## Architecture

- Separate workbook inventory, candidate discovery, and data extraction.
- Preserve source coordinates and typed values in all extracted results.
- Do not infer a table solely from styles, blank cells, or worksheet dimensions.
- Prefer explicit ranges and workbook-authored structure over heuristics.
- Report ambiguity rather than silently selecting one candidate.
- Keep header normalization exact and deterministic; fuzzy matching is not an MVP behavior.

## Change discipline

- Preserve unrelated user changes.
- Add contract-level tests for every public behavior change.
- Give every function and method under `src/excel_data_reader` a concise behavioral docstring,
  with a description for each explicit parameter other than `self` and `cls`.
- Use stable diagnostic codes with worksheet and cell locations when available.
- Run `uv run pytest`, `uv run ruff check src tests examples`,
  `uv run ruff format --check src tests examples`,
  and `uv run ty check` before handoff.
- Treat service timeouts as cooperative. Recommend an isolated, resource-limited worker whenever
  a caller requires hard execution limits.

When validating private production workbooks, use `EXCEL_DATA_READER_ACCEPTANCE_MANIFEST` and do
not add confidential files to the repository.
