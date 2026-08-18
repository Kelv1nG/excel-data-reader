# Workbook acceptance corpus

`manifest.json` is the checked-in seed corpus. Each case names a workbook operation and asserts
the discovered source, range, projected columns, row count, and representative source cells.

Add anonymized production workbooks only when they are safe to commit. Put each workbook beside a
new manifest or reference it with a path relative to that manifest. To run a private local corpus
without committing it, set `EXCEL_DATA_READER_ACCEPTANCE_MANIFEST` to an additional manifest path:

```powershell
$env:EXCEL_DATA_READER_ACCEPTANCE_MANIFEST = "C:\safe\corpus\manifest.json"
uv run pytest tests/test_acceptance.py
```

Keep expected assertions narrow but meaningful. Prefer boundary, column-coordinate, and selected
cell assertions over snapshots of entire workbooks. Never place confidential source files in this
repository.
