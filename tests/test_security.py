from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from openpyxl.xml import defusedxml_available

from excel_data_reader import (
    AnalysisRequest,
    AnalysisStatus,
    DiagnosticCode,
    WorkbookFormat,
    WorkbookPolicy,
    WorkbookRejectedError,
    analyze_workbook,
    inspect_workbook,
)

ROOT = Path(__file__).parents[1]
WORKBOOKS = ROOT / "examples" / "workbooks"


def _minimal_archive(path: Path, extra: dict[str, bytes] | None = None) -> Path:
    members = {
        "[Content_Types].xml": b"<Types/>",
        "xl/workbook.xml": b"<workbook/>",
    }
    members.update(extra or {})
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        for name, value in members.items():
            archive.writestr(name, value)
    return path


def _diagnostic(path: Path, policy: WorkbookPolicy | None = None) -> DiagnosticCode:
    with pytest.raises(WorkbookRejectedError) as captured:
        inspect_workbook(path, policy)
    return captured.value.diagnostics[0].code


def test_default_policy_accepts_a_normal_xlsx_and_reports_archive_metadata() -> None:
    inspection = inspect_workbook(WORKBOOKS / "native_table.xlsx")

    assert inspection.extension == ".xlsx"
    assert inspection.format is WorkbookFormat.OOXML
    assert inspection.file_size > 0
    assert len(inspection.sha256) == 64
    assert inspection.archive_entries is not None
    assert inspection.archive_entries > 0
    assert inspection.uncompressed_size is not None
    assert inspection.compressed_size is not None
    assert inspection.uncompressed_size >= inspection.compressed_size
    assert inspection.has_macros is False
    assert inspection.has_external_links is False
    assert defusedxml_available() is True


def test_policy_rejects_unsupported_extension_even_when_content_is_xlsx(
    tmp_path: Path,
) -> None:
    path = tmp_path / "renamed.zip"
    path.write_bytes((WORKBOOKS / "native_table.xlsx").read_bytes())

    assert _diagnostic(path) is DiagnosticCode.UNSUPPORTED_WORKBOOK_FORMAT


def test_policy_rejects_missing_invalid_and_compound_document_inputs(
    tmp_path: Path,
) -> None:
    invalid = tmp_path / "invalid.xlsx"
    invalid.write_bytes(b"not a ZIP")
    compound = tmp_path / "encrypted.xlsx"
    compound.write_bytes(bytes.fromhex("D0CF11E0A1B11AE1") + b"payload")

    assert _diagnostic(tmp_path / "missing.xlsx") is DiagnosticCode.FILE_NOT_FOUND
    assert _diagnostic(invalid) is DiagnosticCode.INVALID_WORKBOOK_ARCHIVE
    assert _diagnostic(compound) is DiagnosticCode.ENCRYPTED_WORKBOOK


def test_policy_enforces_file_archive_member_and_entry_limits(tmp_path: Path) -> None:
    path = _minimal_archive(
        tmp_path / "limited.xlsx",
        {"xl/worksheets/sheet1.xml": b"x" * 1_000},
    )

    assert _diagnostic(path, WorkbookPolicy(max_file_size=10)) is DiagnosticCode.WORKBOOK_TOO_LARGE
    assert (
        _diagnostic(path, WorkbookPolicy(max_archive_entries=2))
        is DiagnosticCode.ARCHIVE_LIMIT_EXCEEDED
    )
    assert (
        _diagnostic(path, WorkbookPolicy(max_member_size=500))
        is DiagnosticCode.ARCHIVE_LIMIT_EXCEEDED
    )
    assert (
        _diagnostic(path, WorkbookPolicy(max_uncompressed_size=500))
        is DiagnosticCode.ARCHIVE_LIMIT_EXCEEDED
    )
    assert (
        _diagnostic(path, WorkbookPolicy(max_compression_ratio=2))
        is DiagnosticCode.ARCHIVE_LIMIT_EXCEEDED
    )


def test_policy_rejects_unsafe_or_incomplete_archives(tmp_path: Path) -> None:
    unsafe = _minimal_archive(tmp_path / "unsafe.xlsx", {"../escape.xml": b"x"})
    incomplete = tmp_path / "incomplete.xlsx"
    with ZipFile(incomplete, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", b"<Types/>")

    assert _diagnostic(unsafe) is DiagnosticCode.INVALID_WORKBOOK_ARCHIVE
    assert _diagnostic(incomplete) is DiagnosticCode.INVALID_WORKBOOK_ARCHIVE


def test_macro_and_external_link_policies_are_explicit(tmp_path: Path) -> None:
    macro = _minimal_archive(
        tmp_path / "macro.xlsm",
        {"xl/vbaProject.bin": b"macro"},
    )
    external = _minimal_archive(
        tmp_path / "external.xlsx",
        {"xl/externalLinks/externalLink1.xml": b"<externalLink/>"},
    )

    assert _diagnostic(macro) is DiagnosticCode.MACROS_NOT_ALLOWED
    assert _diagnostic(external) is DiagnosticCode.EXTERNAL_LINKS_NOT_ALLOWED
    assert inspect_workbook(macro, WorkbookPolicy(allow_macros=True)).has_macros is True
    assert (
        inspect_workbook(external, WorkbookPolicy(allow_external_links=True)).has_external_links
        is True
    )


def test_analysis_service_maps_policy_failures_to_rejected(tmp_path: Path) -> None:
    path = tmp_path / "invalid.xlsx"
    path.write_bytes(b"not a workbook")

    response = analyze_workbook(path, AnalysisRequest.inventory())

    assert response.status is AnalysisStatus.REJECTED
    assert response.diagnostics[0].code is DiagnosticCode.INVALID_WORKBOOK_ARCHIVE


def test_analysis_service_maps_corrupt_ooxml_to_rejected(tmp_path: Path) -> None:
    path = _minimal_archive(tmp_path / "corrupt.xlsx")

    response = analyze_workbook(path, AnalysisRequest.inventory())

    assert response.status is AnalysisStatus.REJECTED
    assert response.diagnostics[0].code is DiagnosticCode.INVALID_WORKBOOK_ARCHIVE
    assert response.inspection is not None


def test_policy_accepts_a_legacy_xls_compound_document() -> None:
    inspection = inspect_workbook(WORKBOOKS / "legacy_scattered.xls")

    assert inspection.format is WorkbookFormat.LEGACY_XLS
    assert inspection.extension == ".xls"
    assert inspection.archive_entries is None
    assert inspection.compressed_size is None
    assert inspection.has_macros is None


def test_policy_rejects_non_compound_content_with_an_xls_extension(tmp_path: Path) -> None:
    path = tmp_path / "fake.xls"
    path.write_bytes(b"not an XLS workbook")

    assert _diagnostic(path) is DiagnosticCode.INVALID_LEGACY_WORKBOOK
