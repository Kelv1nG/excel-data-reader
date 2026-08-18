"""Cooperative cancellation and deadline controls for workbook analysis."""

from __future__ import annotations

import time
from collections.abc import Callable
from dataclasses import dataclass, field

from excel_data_reader.diagnostics import Diagnostic, DiagnosticCode, ExcelDataReaderError


class AnalysisCancelledError(ExcelDataReaderError):
    """Raised at a checkpoint after a caller requests cancellation."""


class AnalysisTimeoutError(ExcelDataReaderError):
    """Raised at a checkpoint after the analysis timeout expires."""


@dataclass(frozen=True)
class AnalysisControl:
    """Configuration for one cooperatively controlled analysis run."""

    timeout_seconds: float | None = None
    is_cancelled: Callable[[], bool] | None = field(default=None, repr=False, compare=False)
    clock: Callable[[], float] = field(default=time.monotonic, repr=False, compare=False)

    def __post_init__(self) -> None:
        if self.timeout_seconds is not None and self.timeout_seconds < 0:
            raise ValueError("timeout_seconds cannot be negative")

    def start(self) -> _AnalysisBudget:
        """Start a fresh execution budget from this reusable configuration."""

        return _AnalysisBudget(
            timeout_seconds=self.timeout_seconds,
            is_cancelled=self.is_cancelled,
            clock=self.clock,
            started_at=self.clock(),
        )


@dataclass(frozen=True)
class _AnalysisBudget:
    timeout_seconds: float | None
    is_cancelled: Callable[[], bool] | None
    clock: Callable[[], float]
    started_at: float

    def checkpoint(self) -> None:
        if self.is_cancelled is not None and self.is_cancelled():
            raise AnalysisCancelledError(
                Diagnostic(DiagnosticCode.ANALYSIS_CANCELLED, "analysis was cancelled")
            )
        if (
            self.timeout_seconds is not None
            and self.clock() - self.started_at >= self.timeout_seconds
        ):
            raise AnalysisTimeoutError(
                Diagnostic(
                    DiagnosticCode.ANALYSIS_TIMEOUT,
                    f"analysis exceeded its {self.timeout_seconds:g} second timeout",
                )
            )
