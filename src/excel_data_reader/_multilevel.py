"""Adapter-neutral resolution of authored multilevel header geometry."""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from itertools import pairwise
from typing import Any

from excel_data_reader.diagnostics import Diagnostic, DiagnosticCode
from excel_data_reader.model import BodyPolicy, BodyPolicyMode, Coordinate, Rectangle
from excel_data_reader.multilevel import (
    MultiLevelColumn,
    MultiLevelHeaderPart,
    MultiLevelTableMatch,
    MultiLevelTableMatchSet,
    MultiLevelTableQuery,
    _name_diagnostics,
)
from excel_data_reader.normalization import normalize_header


@dataclass(frozen=True)
class _ResolvedHeaderBand:
    """Complete header geometry and its deduplicated requested projection."""

    header_bounds: Rectangle
    all_columns: tuple[MultiLevelColumn, ...]
    selected_columns: tuple[MultiLevelColumn, ...]


def _intersects(first: Rectangle, second: Rectangle) -> bool:
    """Return whether two inclusive cell rectangles intersect.

    Args:
        first: First rectangle to compare.
        second: Second rectangle to compare.
    """

    return (
        first.top <= second.bottom
        and second.top <= first.bottom
        and first.left <= second.right
        and second.left <= first.right
    )


def _header_nodes(
    window: Rectangle,
    merges: tuple[Rectangle, ...],
    value_at: Callable[[int, int], Any],
    checkpoint: Callable[[], None],
) -> tuple[MultiLevelHeaderPart, ...]:
    """Read each authored anchor once without propagating unmerged labels sideways.

    Args:
        window: Finite header window to inspect.
        merges: Merge rectangles intersecting the window.
        value_at: Adapter callback returning the stored cell value.
        checkpoint: Cooperative cancellation callback.
    """

    nodes: list[MultiLevelHeaderPart] = []
    for row in range(window.top, window.bottom + 1):
        checkpoint()
        spans = sorted(
            (merge for merge in merges if merge.top <= row <= merge.bottom),
            key=lambda merge: merge.left,
        )
        index = 0
        column = window.left
        while column <= window.right:
            checkpoint()
            while index < len(spans) and spans[index].right < column:
                index += 1
            merge = spans[index] if index < len(spans) and spans[index].left <= column else None
            if merge is not None and (row, column) != (merge.top, merge.left):
                column = merge.right + 1
                continue
            raw = value_at(row, column)
            if raw is not None and normalize_header(raw):
                nodes.append(
                    MultiLevelHeaderPart(
                        str(raw),
                        raw,
                        Coordinate(row, column),
                        merge or Rectangle(row, column, row, column),
                    )
                )
            column = merge.right + 1 if merge is not None else column + 1
    return tuple(nodes)


def _authored_paths(
    window: Rectangle,
    nodes: tuple[MultiLevelHeaderPart, ...],
    checkpoint: Callable[[], None],
) -> dict[int, tuple[MultiLevelHeaderPart, ...]]:
    """Collect covering anchors in row order for every labeled physical column.

    Args:
        window: Bounded columns receiving header labels.
        nodes: Authored nodes in top-to-bottom reading order.
        checkpoint: Cooperative cancellation callback.
    """

    paths: dict[int, list[MultiLevelHeaderPart]] = {}
    for node in nodes:
        checkpoint()
        bounds = node.source_bounds
        if bounds is None:
            raise AssertionError("authored nodes require source bounds")
        for column in range(max(window.left, bounds.left), min(window.right, bounds.right) + 1):
            checkpoint()
            paths.setdefault(column, []).append(node)
    return {column: tuple(parts) for column, parts in paths.items()}


def _selector_occurrences(
    paths: dict[int, tuple[MultiLevelHeaderPart, ...]],
    selector: str | tuple[str, ...],
    checkpoint: Callable[[], None],
) -> dict[Coordinate, tuple[MultiLevelHeaderPart, ...]]:
    """Resolve an authored selector to distinct node identities and complete prefixes.

    Args:
        paths: Authored paths indexed by physical data column.
        selector: Bare label or complete root-to-node path.
        checkpoint: Cooperative cancellation callback.
    """

    target = (
        normalize_header(selector)
        if isinstance(selector, str)
        else tuple(normalize_header(label) for label in selector)
    )
    found: dict[Coordinate, tuple[MultiLevelHeaderPart, ...]] = {}
    for parts in paths.values():
        for index, part in enumerate(parts):
            checkpoint()
            prefix = parts[: index + 1]
            actual = (
                normalize_header(part.label)
                if isinstance(selector, str)
                else tuple(normalize_header(item.label) for item in prefix)
            )
            if actual == target and part.source_coordinate is not None:
                found[part.source_coordinate] = prefix
    return found


def _geometry_error(
    footprint: Rectangle,
    window: Rectangle,
    merges: tuple[Rectangle, ...],
    paths: dict[int, tuple[MultiLevelHeaderPart, ...]],
    checkpoint: Callable[[], None],
) -> str | None:
    """Explain clipped or crossing spans inside a candidate footprint.

    Args:
        footprint: Complete proposed header footprint.
        window: Caller-bounded candidate window.
        merges: All merge rectangles intersecting the window.
        paths: Authored column paths used to check parent containment.
        checkpoint: Cooperative cancellation callback.
    """

    if not window.contains_rectangle(footprint):
        return "header ancestry extends outside the query window"
    relevant = tuple(merge for merge in merges if _intersects(merge, footprint))
    for index, merge in enumerate(relevant):
        checkpoint()
        if not footprint.contains_rectangle(merge):
            return f"merge {merge.a1} is clipped by the header or query boundary"
        for other in relevant[index + 1 :]:
            checkpoint()
            if _intersects(merge, other):
                return f"merges {merge.a1} and {other.a1} overlap"
    for column, parts in paths.items():
        if not footprint.left <= column <= footprint.right:
            continue
        for parent, child in pairwise(parts):
            checkpoint()
            first, second = parent.source_bounds, child.source_bounds
            if first is None or second is None:
                raise AssertionError("authored parts require bounds")
            if not first.left <= second.left <= second.right <= first.right:
                return f"child span {second.a1} crosses parent span {first.a1}"
    return None


def resolve_band(
    *,
    sheet: str,
    header_window: Rectangle,
    merged_ranges: tuple[Rectangle, ...],
    value_at: Callable[[int, int], Any],
    query: MultiLevelTableQuery,
    checkpoint: Callable[[], None],
) -> tuple[_ResolvedHeaderBand | None, tuple[Diagnostic, ...]]:
    """Resolve one bounded header window into complete paths and selected columns.

    Args:
        sheet: Worksheet name used in diagnostics.
        header_window: Finite window with exactly the declared header depth.
        merged_ranges: Worksheet-authored merged rectangles.
        value_at: Adapter callback returning stored values without filling merges.
        query: Required authored selectors and header depth.
        checkpoint: Cooperative cancellation callback.
    """

    if header_window.height != query.header_depth:
        raise ValueError("header window must have the declared depth")
    merges = tuple(merge for merge in merged_ranges if _intersects(merge, header_window))
    nodes = _header_nodes(header_window, merges, value_at, checkpoint)
    paths = _authored_paths(header_window, nodes, checkpoint)
    occurrences = tuple(
        _selector_occurrences(paths, selector, checkpoint) for selector in query.required_headers
    )
    if any(not found for found in occurrences):
        return None, ()
    ancestry = [
        part.source_bounds for found in occurrences for prefix in found.values() for part in prefix
    ]
    spans = [bounds for bounds in ancestry if bounds is not None]
    footprint = Rectangle(
        header_window.top,
        min(bounds.left for bounds in spans),
        header_window.bottom,
        max(bounds.right for bounds in spans),
    )
    malformed = _geometry_error(footprint, header_window, merges, paths, checkpoint)
    if malformed:
        return None, (
            Diagnostic(
                DiagnosticCode.MALFORMED_MULTILEVEL_HEADER,
                malformed,
                sheet=sheet,
                address=footprint.a1,
            ),
        )
    if not any(
        node.source_coordinate is not None
        and node.source_coordinate.row == footprint.top
        and footprint.left <= node.source_coordinate.column <= footprint.right
        for node in nodes
    ):
        return None, ()
    if any(len(found) > 1 for found in occurrences):
        locations = sorted(
            {coordinate.a1 for found in occurrences if len(found) > 1 for coordinate in found}
        )
        return None, (
            Diagnostic(
                DiagnosticCode.AMBIGUOUS_MULTILEVEL_HEADER,
                f"selector matches multiple authored nodes: {', '.join(locations)}",
                sheet=sheet,
                address=footprint.a1,
            ),
        )

    columns: list[MultiLevelColumn] = []
    for column in range(footprint.left, footprint.right + 1):
        checkpoint()
        parts = paths.get(column, ())
        if not parts:
            continue
        last = parts[-1].source_bounds
        if last is not None and last.width > 1:
            parts += (MultiLevelHeaderPart(f"col{column - last.left + 1}", None, None, None),)
        columns.append(MultiLevelColumn(column, parts))
    projection: dict[int, MultiLevelColumn] = {}
    for found in occurrences:
        coordinate = next(iter(found))
        for column in columns:
            checkpoint()
            if any(part.source_coordinate == coordinate for part in column.parts):
                projection.setdefault(column.source_column, column)
    selected = tuple(projection.values())
    diagnostics = _name_diagnostics(selected, sheet, footprint)
    if diagnostics:
        return None, diagnostics
    return _ResolvedHeaderBand(footprint, tuple(columns), selected), ()


def _signature(band: _ResolvedHeaderBand) -> tuple[Any, ...]:
    """Return the full physical header signature independent of absolute row positions.

    Args:
        band: Complete authored band including unselected sibling columns.
    """

    columns = {column.source_column: column for column in band.all_columns}
    paths = tuple(
        tuple(
            (
                normalize_header(part.label),
                part.is_synthetic,
                part.source_bounds.left if part.source_bounds else column,
                part.source_bounds.right if part.source_bounds else column,
            )
            for part in columns[column].parts
        )
        if column in columns
        else ()
        for column in range(band.header_bounds.left, band.header_bounds.right + 1)
    )
    return band.header_bounds.height, band.header_bounds.left, band.header_bounds.right, paths


def _body_bottom(
    band: _ResolvedHeaderBand,
    hard_bottom: int,
    policy: BodyPolicy,
    value_at: Callable[[int, int], Any],
    checkpoint: Callable[[], None],
) -> int:
    """Infer the last populated row within a fully discovered structural cap.

    Args:
        band: Header and projected physical columns.
        hard_bottom: Inclusive last row permitted by scope and the next repetition.
        policy: Caller-selected blank, populated, or explicit body policy.
        value_at: Callback returning stored data values.
        checkpoint: Cooperative cancellation callback.
    """

    if policy.mode is BodyPolicyMode.EXPLICIT:
        if policy.bottom_row is None:
            raise AssertionError("explicit policy requires a bottom row")
        return policy.bottom_row
    last = band.header_bounds.bottom
    blanks = 0
    for row in range(last + 1, hard_bottom + 1):
        checkpoint()
        if any(value_at(row, column.source_column) is not None for column in band.selected_columns):
            last, blanks = row, 0
        else:
            blanks += 1
            if policy.mode is BodyPolicyMode.BLANK_ROWS and blanks >= policy.blank_rows:
                break
    return last


def discover_sheet(
    *,
    sheet: str,
    scan_bounds: Rectangle,
    merged_ranges: tuple[Rectangle, ...],
    value_at: Callable[[int, int], Any],
    query: MultiLevelTableQuery,
    max_candidates: int,
    checkpoint: Callable[[], None],
) -> MultiLevelTableMatchSet:
    """Discover complete headers before independently bounding their data bodies.

    Args:
        sheet: Worksheet name for matches and diagnostics.
        scan_bounds: Finite scan area already checked against the reader cell budget.
        merged_ranges: Authored worksheet merges, including those crossing scope.
        value_at: Adapter callback returning stored cell values.
        query: Authored selectors, explicit header depth, and body policy.
        max_candidates: Remaining global candidate budget, possibly zero.
        checkpoint: Cooperative cancellation callback.
    """

    if query.header_depth > scan_bounds.height:
        return MultiLevelTableMatchSet(())
    labels = {
        normalize_header(selector if isinstance(selector, str) else selector[-1])
        for selector in query.required_headers
    }
    rows_by_label: dict[str, set[int]] = {label: set() for label in labels}
    values: dict[tuple[int, int], Any] = {}
    for row in range(scan_bounds.top, scan_bounds.bottom + 1):
        for column in range(scan_bounds.left, scan_bounds.right + 1):
            checkpoint()
            value = value_at(row, column)
            if value is not None:
                values[row, column] = value
                normalized = normalize_header(value)
                if normalized in rows_by_label:
                    rows_by_label[normalized].add(row)

    def stored(row: int, column: int) -> Any:
        """Read a sparse cell snapshot without repeating adapter work.

        Args:
            row: One-based physical row.
            column: One-based physical column.
        """

        return values.get((row, column))

    starts: set[int] | None = None
    for rows in rows_by_label.values():
        eligible: set[int] = set()
        for row in rows:
            checkpoint()
            eligible.update(
                range(
                    max(scan_bounds.top, row - query.header_depth + 1),
                    min(row, scan_bounds.bottom - query.header_depth + 1) + 1,
                )
            )
        starts = eligible if starts is None else starts & eligible
    bands: list[_ResolvedHeaderBand] = []
    diagnostics: list[Diagnostic] = []
    for top in sorted(starts or ()):
        checkpoint()
        band, errors = resolve_band(
            sheet=sheet,
            header_window=Rectangle(
                top, scan_bounds.left, top + query.header_depth - 1, scan_bounds.right
            ),
            merged_ranges=merged_ranges,
            value_at=stored,
            query=query,
            checkpoint=checkpoint,
        )
        diagnostics.extend(errors)
        if band is not None:
            bands.append(band)
        if len(bands) > max_candidates or len(diagnostics) > max(1, max_candidates):
            return MultiLevelTableMatchSet(
                (),
                (
                    *diagnostics[: max(1, max_candidates)],
                    Diagnostic(
                        DiagnosticCode.SCAN_LIMIT_EXCEEDED,
                        "multilevel candidate or diagnostic budget exceeded; "
                        "sheet results discarded",
                        sheet=sheet,
                        address=scan_bounds.a1,
                    ),
                ),
                is_complete=False,
                candidate_count=len(bands),
            )
    signatures = [_signature(band) for band in bands]
    matches: list[MultiLevelTableMatch] = []
    for index, band in enumerate(bands):
        checkpoint()
        cap = scan_bounds.bottom
        for following in range(index + 1, len(bands)):
            checkpoint()
            if (
                bands[following].header_bounds.top > band.header_bounds.bottom
                and signatures[following] == signatures[index]
            ):
                cap = bands[following].header_bounds.top - 1
                break
        if query.body.mode is BodyPolicyMode.EXPLICIT:
            bottom = query.body.bottom_row
            if bottom is None or not band.header_bounds.bottom <= bottom <= cap:
                diagnostics.append(
                    Diagnostic(
                        DiagnosticCode.MULTILEVEL_BODY_BOUNDARY_CONFLICT,
                        f"explicit bottom {bottom} is outside body bounds "
                        f"{band.header_bounds.bottom}:{cap}",
                        sheet=sheet,
                        address=band.header_bounds.a1,
                    )
                )
                continue
        bottom = _body_bottom(band, cap, query.body, stored, checkpoint)
        matches.append(
            MultiLevelTableMatch(
                sheet,
                band.header_bounds,
                band.selected_columns,
                band.header_bounds.bottom + 1,
                bottom,
            )
        )
    return MultiLevelTableMatchSet(tuple(matches), tuple(diagnostics), candidate_count=len(bands))
