"""Public API for excel-data-reader."""

from excel_data_reader.diagnostics import (
    Diagnostic,
    DiagnosticCode,
    ExcelDataReaderError,
    Severity,
)
from excel_data_reader.model import (
    CellData,
    ColumnInfo,
    Confidence,
    Coordinate,
    DataRow,
    FormulaValue,
    MatchSet,
    MatchSource,
    NamedRangeInfo,
    NativeTableInfo,
    RangeReference,
    Rectangle,
    SheetData,
    SheetInfo,
    TableData,
    TableMatch,
    ValueMode,
    WorkbookInventory,
)
from excel_data_reader.normalization import normalize_header
from excel_data_reader.reader import ExcelReader

__all__ = [
    "CellData",
    "ColumnInfo",
    "Confidence",
    "Coordinate",
    "DataRow",
    "Diagnostic",
    "DiagnosticCode",
    "ExcelDataReaderError",
    "ExcelReader",
    "FormulaValue",
    "MatchSet",
    "MatchSource",
    "NamedRangeInfo",
    "NativeTableInfo",
    "RangeReference",
    "Rectangle",
    "Severity",
    "SheetData",
    "SheetInfo",
    "TableData",
    "TableMatch",
    "ValueMode",
    "WorkbookInventory",
    "normalize_header",
]
