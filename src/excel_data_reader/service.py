"""Versioned, platform-facing workbook analysis contract."""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import BinaryIO
from unicodedata import category
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from defusedxml.common import DefusedXmlException
from openpyxl.utils.exceptions import InvalidFileException

from excel_data_reader.control import (
    AnalysisCancelledError,
    AnalysisControl,
    AnalysisTimeoutError,
    _AnalysisBudget,
)
from excel_data_reader.diagnostics import Diagnostic, ExcelDataReaderError
from excel_data_reader.diagnostics import DiagnosticCode as Code
from excel_data_reader.legacy import LegacyWorkbookError
from excel_data_reader.model import (
    DataRow,
    DiscoveryReport,
    TableMatch,
    TableQuery,
    ValueMode,
    WorkbookInventory,
)
from excel_data_reader.reader import ExcelReader
from excel_data_reader.security import (
    WorkbookInspection,
    WorkbookPolicy,
    WorkbookRejectedError,
    inspect_workbook,
)
from excel_data_reader.serialization import JSON_VALUE_SCHEMA_VERSION, to_json

ANALYSIS_SCHEMA_VERSION = "1.1"


class AnalysisOperation(StrEnum):
    INVENTORY = "inventory"
    FIND_TABLES = "find_tables"


class AnalysisStatus(StrEnum):
    SUCCESS = "success"
    NO_MATCH = "no_match"
    AMBIGUOUS = "ambiguous"
    REJECTED = "rejected"
    ERROR = "error"
    CANCELLED = "cancelled"
    TIMEOUT = "timeout"


@dataclass(frozen=True)
class AnalysisRequest:
    operation: AnalysisOperation
    query: TableQuery | None = None
    include_inventory: bool = True
    include_rows: bool = False
    max_output_rows: int = 1_000
    value_mode: ValueMode = ValueMode.FORMULA
    request_id: str | None = None

    def __post_init__(self) -> None:
        object.__setattr__(self, "operation", AnalysisOperation(self.operation))
        object.__setattr__(self, "value_mode", ValueMode(self.value_mode))
        if self.operation is AnalysisOperation.FIND_TABLES and self.query is None:
            raise ValueError("find_tables analysis requires a TableQuery")
        if self.operation is AnalysisOperation.INVENTORY and self.query is not None:
            raise ValueError("inventory analysis does not accept a TableQuery")
        if self.max_output_rows < 0:
            raise ValueError("max_output_rows cannot be negative")

    @classmethod
    def inventory(
        cls,
        *,
        request_id: str | None = None,
    ) -> AnalysisRequest:
        return cls(AnalysisOperation.INVENTORY, request_id=request_id)

    @classmethod
    def find_tables(
        cls,
        query: TableQuery,
        *,
        include_inventory: bool = True,
        include_rows: bool = False,
        max_output_rows: int = 1_000,
        value_mode: ValueMode | str = ValueMode.FORMULA,
        request_id: str | None = None,
    ) -> AnalysisRequest:
        return cls(
            AnalysisOperation.FIND_TABLES,
            query=query,
            include_inventory=include_inventory,
            include_rows=include_rows,
            max_output_rows=max_output_rows,
            value_mode=ValueMode(value_mode),
            request_id=request_id,
        )


@dataclass(frozen=True)
class ExtractedTable:
    match: TableMatch
    rows: tuple[DataRow, ...]
    total_row_count: int
    truncated: bool


@dataclass(frozen=True)
class AnalysisResponse:
    schema_version: str
    value_schema_version: int
    request_id: str | None
    source_name: str
    operation: AnalysisOperation
    status: AnalysisStatus
    inventory: WorkbookInventory | None = None
    discovery: DiscoveryReport | None = None
    tables: tuple[ExtractedTable, ...] = ()
    diagnostics: tuple[Diagnostic, ...] = ()
    inspection: WorkbookInspection | None = None

    def to_json(self, *, indent: int | None = None) -> str:
        return to_json(self, indent=indent)


def analyze_workbook(
    path: str | Path,
    request: AnalysisRequest,
    *,
    max_scan_cells: int = 2_000_000,
    max_candidates: int = 100,
    policy: WorkbookPolicy | None = None,
    control: AnalysisControl | None = None,
) -> AnalysisResponse:
    """Analyze a workbook through a stable, JSON-serializable service boundary."""

    workbook_path = Path(path)
    return _analyze_path(
        workbook_path,
        request,
        source_name=workbook_path.name,
        max_scan_cells=max_scan_cells,
        max_candidates=max_candidates,
        policy=policy or WorkbookPolicy(),
        budget=(control or AnalysisControl()).start(),
    )


def analyze_workbook_bytes(
    data: bytes | bytearray | memoryview | BinaryIO,
    filename: str,
    request: AnalysisRequest,
    *,
    max_scan_cells: int = 2_000_000,
    max_candidates: int = 100,
    policy: WorkbookPolicy | None = None,
    control: AnalysisControl | None = None,
    temp_dir: str | Path | None = None,
) -> AnalysisResponse:
    """Stage an uploaded workbook safely, analyze it, and remove the staged file."""

    source_name = _source_name(filename)
    active_policy = policy or WorkbookPolicy()
    budget = (control or AnalysisControl()).start()
    try:
        budget.checkpoint()
        with TemporaryDirectory(prefix="excel-data-reader-", dir=temp_dir) as directory:
            staged_path = Path(directory) / f"upload{Path(source_name).suffix.casefold()}"
            _write_upload(
                data,
                staged_path,
                max_file_size=active_policy.max_file_size,
                checkpoint=budget.checkpoint,
            )
            return _analyze_path(
                staged_path,
                request,
                source_name=source_name,
                max_scan_cells=max_scan_cells,
                max_candidates=max_candidates,
                policy=active_policy,
                budget=budget,
            )
    except AnalysisCancelledError as error:
        return _response(
            source_name,
            request,
            AnalysisStatus.CANCELLED,
            diagnostics=error.diagnostics,
        )
    except AnalysisTimeoutError as error:
        return _response(
            source_name,
            request,
            AnalysisStatus.TIMEOUT,
            diagnostics=error.diagnostics,
        )
    except WorkbookRejectedError as error:
        return _response(
            source_name,
            request,
            AnalysisStatus.REJECTED,
            diagnostics=error.diagnostics,
        )
    except OSError:
        return _response(
            source_name,
            request,
            AnalysisStatus.ERROR,
            diagnostics=(
                Diagnostic(
                    Code.UPLOAD_STAGING_FAILED,
                    "uploaded workbook could not be staged for analysis",
                ),
            ),
        )


def _analyze_path(
    workbook_path: Path,
    request: AnalysisRequest,
    *,
    source_name: str,
    max_scan_cells: int,
    max_candidates: int,
    policy: WorkbookPolicy,
    budget: _AnalysisBudget,
) -> AnalysisResponse:
    inspection: WorkbookInspection | None = None
    try:
        budget.checkpoint()
        inspection = inspect_workbook(
            workbook_path,
            policy,
            checkpoint=budget.checkpoint,
        )
        with ExcelReader.open(
            workbook_path,
            value_mode=request.value_mode,
            max_scan_cells=max_scan_cells,
            max_candidates=max_candidates,
            checkpoint=budget.checkpoint,
        ) as reader:
            inventory = reader.inventory() if request.include_inventory else None
            if request.operation is AnalysisOperation.INVENTORY:
                if inventory is None:
                    inventory = reader.inventory()
                return _response(
                    source_name,
                    request,
                    AnalysisStatus.SUCCESS,
                    inventory=inventory,
                    diagnostics=reader.diagnostics,
                    inspection=inspection,
                )

            if request.query is None:
                raise AssertionError("validated find_tables request has no query")
            discovery = reader.explain(request.query)
            match_count = len(discovery.selected_matches)
            status = (
                AnalysisStatus.NO_MATCH
                if match_count == 0
                else AnalysisStatus.AMBIGUOUS
                if match_count > 1
                else AnalysisStatus.SUCCESS
            )
            tables: tuple[ExtractedTable, ...] = ()
            if request.include_rows and match_count == 1:
                table = reader.extract(discovery.selected_matches[0])
                rows = table.rows[: request.max_output_rows]
                tables = (
                    ExtractedTable(
                        match=table.match,
                        rows=rows,
                        total_row_count=len(table.rows),
                        truncated=len(rows) < len(table.rows),
                    ),
                )
            return _response(
                source_name,
                request,
                status,
                inventory=inventory,
                discovery=discovery,
                tables=tables,
                diagnostics=reader.diagnostics + discovery.diagnostics,
                inspection=inspection,
            )
    except AnalysisCancelledError as error:
        return _response(
            source_name,
            request,
            AnalysisStatus.CANCELLED,
            diagnostics=error.diagnostics,
            inspection=inspection,
        )
    except AnalysisTimeoutError as error:
        return _response(
            source_name,
            request,
            AnalysisStatus.TIMEOUT,
            diagnostics=error.diagnostics,
            inspection=inspection,
        )
    except WorkbookRejectedError as error:
        return _response(
            source_name,
            request,
            AnalysisStatus.REJECTED,
            diagnostics=error.diagnostics,
            inspection=inspection,
        )
    except LegacyWorkbookError as error:
        return _response(
            source_name,
            request,
            AnalysisStatus.REJECTED,
            diagnostics=error.diagnostics,
            inspection=inspection,
        )
    except (BadZipFile, DefusedXmlException, InvalidFileException, OSError, ParseError) as error:
        return _response(
            source_name,
            request,
            AnalysisStatus.REJECTED,
            diagnostics=(
                Diagnostic(
                    Code.INVALID_WORKBOOK_ARCHIVE,
                    f"workbook content could not be parsed ({type(error).__name__})",
                ),
            ),
            inspection=inspection,
        )
    except ExcelDataReaderError as error:
        return _response(
            source_name,
            request,
            AnalysisStatus.ERROR,
            diagnostics=error.diagnostics,
            inspection=inspection,
        )


def _response(
    source_name: str,
    request: AnalysisRequest,
    status: AnalysisStatus,
    *,
    inventory: WorkbookInventory | None = None,
    discovery: DiscoveryReport | None = None,
    tables: tuple[ExtractedTable, ...] = (),
    diagnostics: tuple[Diagnostic, ...] = (),
    inspection: WorkbookInspection | None = None,
) -> AnalysisResponse:
    return AnalysisResponse(
        schema_version=ANALYSIS_SCHEMA_VERSION,
        value_schema_version=JSON_VALUE_SCHEMA_VERSION,
        request_id=request.request_id,
        source_name=source_name,
        operation=request.operation,
        status=status,
        inventory=inventory,
        discovery=discovery,
        tables=tables,
        diagnostics=diagnostics,
        inspection=inspection,
    )


def _write_upload(
    data: bytes | bytearray | memoryview | BinaryIO,
    path: Path,
    *,
    max_file_size: int,
    checkpoint: Callable[[], None],
) -> None:
    total = 0
    with path.open("wb") as destination:
        if isinstance(data, (bytes, bytearray, memoryview)):
            total = len(data)
            if total > max_file_size:
                _reject_large_upload(total, max_file_size)
            destination.write(bytes(data))
            checkpoint()
            return

        while True:
            checkpoint()
            chunk = data.read(1024 * 1024)
            if not chunk:
                break
            if not isinstance(chunk, bytes):
                raise TypeError("uploaded binary stream must return bytes")
            total += len(chunk)
            if total > max_file_size:
                _reject_large_upload(total, max_file_size)
            destination.write(chunk)


def _reject_large_upload(actual_size: int, maximum_size: int) -> None:
    raise WorkbookRejectedError(
        Diagnostic(
            Code.WORKBOOK_TOO_LARGE,
            f"upload exceeds the {maximum_size:,} byte limit ({actual_size:,} bytes read)",
        )
    )


def _source_name(filename: str) -> str:
    normalized = filename.replace("\\", "/")
    basename = normalized.rsplit("/", maxsplit=1)[-1].strip()
    printable = "".join(
        "_" if category(character).startswith("C") else character for character in basename
    )
    return printable[-255:] or "workbook"
