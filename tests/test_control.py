from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pytest

from excel_data_reader import (
    AnalysisControl,
    AnalysisRequest,
    AnalysisStatus,
    DiagnosticCode,
    TableQuery,
    WorkbookPolicy,
    analyze_workbook,
    analyze_workbook_bytes,
)

ROOT = Path(__file__).parents[1]
WORKBOOKS = ROOT / "examples" / "workbooks"


def test_analysis_can_be_cancelled_before_work_starts() -> None:
    response = analyze_workbook(
        WORKBOOKS / "native_table.xlsx",
        AnalysisRequest.inventory(),
        control=AnalysisControl(is_cancelled=lambda: True),
    )

    assert response.status is AnalysisStatus.CANCELLED
    assert response.diagnostics[0].code is DiagnosticCode.ANALYSIS_CANCELLED


def test_analysis_timeout_is_a_stable_response() -> None:
    response = analyze_workbook(
        WORKBOOKS / "native_table.xlsx",
        AnalysisRequest.inventory(),
        control=AnalysisControl(timeout_seconds=0),
    )

    assert response.status is AnalysisStatus.TIMEOUT
    assert response.diagnostics[0].code is DiagnosticCode.ANALYSIS_TIMEOUT


def test_cancellation_is_checked_during_table_scanning() -> None:
    checkpoints = 0

    def cancel_after_initialization() -> bool:
        nonlocal checkpoints
        checkpoints += 1
        return checkpoints >= 20

    response = analyze_workbook(
        WORKBOOKS / "scattered_headers.xlsx",
        AnalysisRequest.find_tables(TableQuery(("never present",))),
        control=AnalysisControl(is_cancelled=cancel_after_initialization),
    )

    assert response.status is AnalysisStatus.CANCELLED
    assert checkpoints == 20


def test_uploaded_bytes_are_staged_with_a_safe_name_and_cleaned_up(tmp_path: Path) -> None:
    staging = tmp_path / "staging"
    staging.mkdir()
    content = (WORKBOOKS / "native_table.xlsx").read_bytes()

    response = analyze_workbook_bytes(
        content,
        "../../customer-upload.xlsx",
        AnalysisRequest.inventory(request_id="upload-1"),
        temp_dir=staging,
    )

    assert response.status is AnalysisStatus.SUCCESS
    assert response.source_name == "customer-upload.xlsx"
    assert response.request_id == "upload-1"
    assert response.inspection is not None
    assert list(staging.iterdir()) == []


def test_uploaded_binary_stream_is_supported(tmp_path: Path) -> None:
    content = (WORKBOOKS / "native_table.xlsx").read_bytes()

    response = analyze_workbook_bytes(
        BytesIO(content),
        "stream.xlsx",
        AnalysisRequest.inventory(),
        temp_dir=tmp_path,
    )

    assert response.status is AnalysisStatus.SUCCESS
    assert response.source_name == "stream.xlsx"
    assert list(tmp_path.iterdir()) == []


def test_oversized_upload_is_rejected_and_cleaned_up(tmp_path: Path) -> None:
    response = analyze_workbook_bytes(
        b"x" * 11,
        "large.xlsx",
        AnalysisRequest.inventory(),
        policy=WorkbookPolicy(max_file_size=10),
        temp_dir=tmp_path,
    )

    assert response.status is AnalysisStatus.REJECTED
    assert response.diagnostics[0].code is DiagnosticCode.WORKBOOK_TOO_LARGE
    assert list(tmp_path.iterdir()) == []


def test_analysis_control_rejects_negative_timeout() -> None:
    with pytest.raises(ValueError, match="cannot be negative"):
        AnalysisControl(timeout_seconds=-0.1)


def test_upload_staging_failure_is_a_stable_error_response(tmp_path: Path) -> None:
    content = (WORKBOOKS / "native_table.xlsx").read_bytes()

    response = analyze_workbook_bytes(
        content,
        "customer\x00upload.xlsx",
        AnalysisRequest.inventory(),
        temp_dir=tmp_path / "missing",
    )

    assert response.status is AnalysisStatus.ERROR
    assert response.source_name == "customer_upload.xlsx"
    assert response.diagnostics[0].code is DiagnosticCode.UPLOAD_STAGING_FAILED
