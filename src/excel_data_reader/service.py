"""Versioned, platform-facing workbook analysis contract."""

from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path

from excel_data_reader.diagnostics import Diagnostic, ExcelDataReaderError
from excel_data_reader.model import (
    DataRow,
    DiscoveryReport,
    TableMatch,
    TableQuery,
    ValueMode,
    WorkbookInventory,
)
from excel_data_reader.reader import ExcelReader
from excel_data_reader.serialization import JSON_VALUE_SCHEMA_VERSION, to_json

ANALYSIS_SCHEMA_VERSION = "1.0"


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

    def to_json(self, *, indent: int | None = None) -> str:
        return to_json(self, indent=indent)


def analyze_workbook(
    path: str | Path,
    request: AnalysisRequest,
    *,
    max_scan_cells: int = 2_000_000,
    max_candidates: int = 100,
) -> AnalysisResponse:
    """Analyze a workbook through a stable, JSON-serializable service boundary."""

    workbook_path = Path(path)
    try:
        with ExcelReader.open(
            workbook_path,
            value_mode=request.value_mode,
            max_scan_cells=max_scan_cells,
            max_candidates=max_candidates,
        ) as reader:
            inventory = reader.inventory() if request.include_inventory else None
            if request.operation is AnalysisOperation.INVENTORY:
                if inventory is None:
                    inventory = reader.inventory()
                return _response(
                    workbook_path,
                    request,
                    AnalysisStatus.SUCCESS,
                    inventory=inventory,
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
                workbook_path,
                request,
                status,
                inventory=inventory,
                discovery=discovery,
                tables=tables,
                diagnostics=discovery.diagnostics,
            )
    except ExcelDataReaderError as error:
        return _response(
            workbook_path,
            request,
            AnalysisStatus.ERROR,
            diagnostics=error.diagnostics,
        )


def _response(
    path: Path,
    request: AnalysisRequest,
    status: AnalysisStatus,
    *,
    inventory: WorkbookInventory | None = None,
    discovery: DiscoveryReport | None = None,
    tables: tuple[ExtractedTable, ...] = (),
    diagnostics: tuple[Diagnostic, ...] = (),
) -> AnalysisResponse:
    return AnalysisResponse(
        schema_version=ANALYSIS_SCHEMA_VERSION,
        value_schema_version=JSON_VALUE_SCHEMA_VERSION,
        request_id=request.request_id,
        source_name=path.name,
        operation=request.operation,
        status=status,
        inventory=inventory,
        discovery=discovery,
        tables=tables,
        diagnostics=diagnostics,
    )
