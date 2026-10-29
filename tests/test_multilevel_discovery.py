import pytest

from excel_data_reader import BodyPolicy, DiagnosticCode, MultiLevelTableQuery, Rectangle
from excel_data_reader import _multilevel as engine
from excel_data_reader._multilevel import discover_sheet
from tests.test_multilevel_headers import layout


def repeated():
    values, merges, _ = layout()
    cells = dict(values)
    cells.update({(row + 6, col): value for (row, col), value in values.items()})
    combined = merges + tuple(Rectangle(m.top + 6, m.left, m.bottom + 6, m.right) for m in merges)
    for row in (4, 5, 6, 10, 11, 12):
        cells[row, 1] = row
        cells[row, 2] = row * 10
    return cells, combined


def discover(
    cells,
    merges=(),
    *,
    depth=3,
    headers=("Sales Report",),
    bottom=12,
    width=3,
    body=None,
    limit=100,
    checkpoint=lambda: None,
):
    return discover_sheet(
        sheet="Report",
        scan_bounds=Rectangle(1, 1, bottom, width),
        merged_ranges=merges,
        value_at=lambda row, col: cells.get((row, col)),
        query=MultiLevelTableQuery(headers, depth, body=body or BodyPolicy()),
        max_candidates=limit,
        checkpoint=checkpoint,
    )


def test_repeated_headers_without_separator_end_independent_bodies():
    result = discover(*repeated())
    assert result.is_complete
    assert [(m.data_start_row, m.data_end_row) for m in result.matches] == [(4, 6), (10, 12)]
    first, second = result.matches
    assert [c.name for c in first.columns] == [c.name for c in second.columns]
    assert first.columns[0].parts[0].source_coordinate.row == 1
    assert second.columns[0].parts[0].source_coordinate.row == 7


def test_limit_requires_complete_scan_not_just_exact_count():
    cells, merges = repeated()
    result = discover(cells, merges, limit=2)
    assert result.is_complete and len(result.matches) == 2
    result = discover(cells, merges, limit=1)
    assert not result.is_complete and not result.matches
    assert result.diagnostics[-1].code is DiagnosticCode.SCAN_LIMIT_EXCEEDED
    assert not discover(cells, merges, limit=0).is_complete


def test_only_full_same_footprint_signature_is_a_boundary():
    cells, merges = repeated()
    cells[9, 2] = "Different"
    result = discover(
        cells, merges, headers=(("Sales Report", "col1"),), body=BodyPolicy.last_populated()
    )
    assert result.matches[0].data_end_row == 12
    cells[9, 2] = "New"
    assert (
        discover(
            cells, merges, headers=(("Sales Report", "col1"),), body=BodyPolicy.last_populated()
        )
        .matches[0]
        .data_end_row
        == 6
    )


def test_lone_repeated_word_is_data():
    cells, merges, _ = layout()
    cells.update({(4, 1): 1, (5, 1): "Sales Report", (6, 1): 2})
    result = discover(cells, merges, headers=("Sales Report", "Online"))
    assert len(result.matches) == 1 and result.matches[0].data_end_row == 6


def test_different_column_footprint_does_not_cap_body():
    cells, merges = repeated()
    cells = {(r, c + 4 if r >= 7 else c): v for (r, c), v in cells.items()}
    merges = tuple(
        Rectangle(m.top, m.left + 4, m.bottom, m.right + 4) if m.top >= 7 else m for m in merges
    )
    cells[12, 1] = 999
    result = discover(cells, merges, width=7, body=BodyPolicy.last_populated())
    assert len(result.matches) == 2
    assert result.matches[0].data_end_row == 12


def test_overlapping_windows_are_not_silently_chosen():
    cells = {(1, 1): "Parent", (2, 1): "Code", (3, 1): "Data", (4, 1): 1}
    result = discover(cells, depth=2, headers=("Code",), width=1, bottom=4)
    assert len(result.matches) == 2
    assert [m.header_bounds.top for m in result.matches] == [1, 2]


def test_body_policies_and_empty_body():
    cells, merges, _ = layout()
    assert discover(cells, merges).matches[0].is_empty
    cells.update({(4, 1): 1, (6, 1): 2, (9, 1): 3})
    assert discover(cells, merges).matches[0].data_end_row == 6
    assert discover(cells, merges, body=BodyPolicy.last_populated()).matches[0].data_end_row == 9
    assert discover(cells, merges, body=BodyPolicy.through_row(7)).matches[0].data_end_row == 7


def test_explicit_bottom_cannot_cross_repetition_or_scope():
    for bottom in (8, 20):
        result = discover(*repeated(), body=BodyPolicy.through_row(bottom))
        assert any(
            d.code is DiagnosticCode.MULTILEVEL_BODY_BOUNDARY_CONFLICT for d in result.diagnostics
        )
        assert all(m.header_bounds.top != 1 for m in result.matches)


def test_cancellation_propagates():
    count = 0

    def checkpoint():
        nonlocal count
        count += 1
        if count == 20:
            raise RuntimeError("cancelled")

    with pytest.raises(RuntimeError, match="cancelled"):
        discover(*repeated(), checkpoint=checkpoint)


@pytest.mark.parametrize("phase", ["resolve_band", "_body_bottom"])
def test_cancellation_inside_resolution_and_body_inference(monkeypatch, phase):
    original = getattr(engine, phase)
    active = False
    count = 0

    def enter_phase(*args, **kwargs):
        nonlocal active
        active = True
        return original(*args, **kwargs)

    def checkpoint():
        nonlocal count
        if active:
            count += 1
            if count == 2:
                raise RuntimeError("cancelled inside " + phase)

    monkeypatch.setattr(engine, phase, enter_phase)
    with pytest.raises(RuntimeError, match="cancelled inside " + phase):
        discover(*repeated(), checkpoint=checkpoint)
    assert count == 2


def test_repetition_compares_vertical_merge_and_blank_as_same_path():
    cells, merges = repeated()
    # The first Store is plain C2; the second is vertically merged C8:C9.
    merges += (Rectangle(8, 3, 9, 3),)
    result = discover(cells, merges, body=BodyPolicy.last_populated())
    assert [(match.data_start_row, match.data_end_row) for match in result.matches] == [
        (4, 6),
        (10, 12),
    ]
