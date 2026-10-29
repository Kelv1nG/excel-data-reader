from excel_data_reader import Coordinate, DiagnosticCode, MultiLevelTableQuery, Rectangle
from excel_data_reader._multilevel import resolve_band


def layout(kind="plain"):
    if kind == "four":
        return (
            {
                (1, 1): "Sales Report",
                (2, 1): "Region",
                (3, 1): "Online",
                (3, 3): "Store",
                (4, 1): "New",
                (4, 2): "Returning",
                (4, 3): "New",
                (4, 4): "Returning",
            },
            (
                Rectangle(1, 1, 1, 4),
                Rectangle(2, 1, 2, 4),
                Rectangle(3, 1, 3, 2),
                Rectangle(3, 3, 3, 4),
            ),
            Rectangle(1, 1, 4, 4),
        )
    width = 5 if kind == "mixed" else 4 if kind == "wide" else 3
    cells = {
        (1, 1): "Sales Report",
        (2, 1): "col1",
        (2, 2): "Online",
        (3, 2): "New",
        (2, 3): "Store",
    }
    merges = [Rectangle(1, 1, 1, width), Rectangle(2, 1, 3, 1)]
    if kind == "merged":
        merges.append(Rectangle(2, 3, 3, 3))
    elif kind in ("wide", "mixed"):
        merges.append(Rectangle(2, 3, 2, width))
    if kind == "mixed":
        cells[3, 3] = "New"
    return cells, tuple(merges), Rectangle(1, 1, 3, width)


def resolve(cells, merges, window, headers=("Sales Report",)):
    return resolve_band(
        sheet="Report",
        header_window=window,
        merged_ranges=merges,
        value_at=lambda row, col: cells.get((row, col)),
        query=MultiLevelTableQuery(headers, window.height),
        checkpoint=lambda: None,
    )


def test_plain_and_vertical_merge_have_identical_names_and_distinct_provenance():
    plain, _ = resolve(*layout())
    merged, _ = resolve(*layout("merged"))
    assert plain is not None and merged is not None
    expected = ["Sales Report__col1", "Sales Report__Online__New", "Sales Report__Store"]
    assert [c.name for c in plain.selected_columns] == expected
    assert [c.name for c in merged.selected_columns] == expected
    assert plain.selected_columns[-1].parts[-1].source_bounds == Rectangle(2, 3, 2, 3)
    assert merged.selected_columns[-1].parts[-1].source_bounds == Rectangle(2, 3, 3, 3)
    assert plain.selected_columns[0].parts[-1].source_coordinate == Coordinate(2, 1)


def test_wide_and_mixed_groups_number_by_physical_position():
    for kind, expected in [("wide", ["col1", "col2"]), ("mixed", ["New", "col2", "col3"])]:
        band, errors = resolve(*layout(kind), headers=("Store",))
        assert band is not None and not errors
        assert [c.path[-1] for c in band.selected_columns] == expected
        assert band.selected_columns[-1].parts[-1].is_synthetic
        assert band.header_bounds.left == 1  # Preserve the Sales Report ancestor.


def test_generated_names_are_not_search_anchors():
    band, errors = resolve(*layout("wide"), headers=("col1", "Store"))
    assert not errors and band is not None
    assert [c.source_column for c in band.selected_columns] == [1, 3, 4]


def test_group_selection_overlap_and_query_order():
    band, _ = resolve(*layout(), headers=(("Sales Report", "Store"), "Sales Report"))
    assert band is not None
    assert [c.source_column for c in band.selected_columns] == [3, 1, 2]


def test_names_can_repeat_in_distinct_authored_nodes():
    band, _ = resolve(
        {(1, 1): "Same", (2, 1): "Same"}, (), Rectangle(1, 1, 2, 1), (("Same", "Same"),)
    )
    assert band is not None
    assert band.selected_columns[0].name == "Same__Same"


def test_ambiguous_bare_name_and_full_path_resolution():
    cells = {(1, 1): "A", (1, 2): "B", (2, 1): "New", (2, 2): "New"}
    band, errors = resolve(cells, (), Rectangle(1, 1, 2, 2), ("New",))
    assert band is None and errors[0].code is DiagnosticCode.AMBIGUOUS_MULTILEVEL_HEADER
    band, errors = resolve(cells, (), Rectangle(1, 1, 2, 2), (("B", "New"),))
    assert not errors and band is not None
    assert band.selected_columns[0].source_column == 2


def test_independent_roots_and_blank_gap_are_query_defined():
    cells = {
        (1, 1): "Sales",
        (1, 4): "Costs",
        (2, 1): "Code",
        (2, 2): "Amount",
        (2, 4): "Tax",
        (2, 5): "Total",
    }
    band, _ = resolve(
        cells,
        (Rectangle(1, 1, 1, 2), Rectangle(1, 4, 1, 5)),
        Rectangle(1, 1, 2, 5),
        ("Code", "Tax"),
    )
    assert band is not None and band.header_bounds == Rectangle(1, 1, 2, 5)
    assert [c.source_column for c in band.selected_columns] == [1, 4]
    assert [c.source_column for c in band.all_columns] == [1, 2, 4, 5]


def test_unrelated_top_label_does_not_qualify_candidate():
    band, _ = resolve({(1, 26): "Notes", (2, 1): "Code"}, (), Rectangle(1, 1, 2, 26), ("Code",))
    assert band is None


def test_unmerged_parent_does_not_fill_sideways():
    band, _ = resolve(
        {(1, 1): "Parent", (2, 1): "A", (2, 2): "B"}, (), Rectangle(1, 1, 2, 2), ("Parent",)
    )
    assert band is not None
    assert [c.name for c in band.selected_columns] == ["Parent__A"]


def test_group_numbering_restarts_and_supports_four_levels():
    cells = {(1, 1): "Root", (2, 1): "Region", (3, 1): "A", (3, 3): "B"}
    merges = (
        Rectangle(1, 1, 1, 4),
        Rectangle(2, 1, 2, 4),
        Rectangle(3, 1, 3, 2),
        Rectangle(3, 3, 3, 4),
    )
    band, _ = resolve(cells, merges, Rectangle(1, 1, 4, 4), ("Root",))
    assert band is not None
    assert [c.name for c in band.selected_columns] == [
        f"Root__Region__{g}__col{n}" for g in ("A", "B") for n in (1, 2)
    ]


def test_clipped_and_crossing_merges_are_malformed():
    for cells, merges, window in [
        ({(1, 1): "A"}, (Rectangle(1, 1, 1, 3),), Rectangle(1, 1, 2, 2)),
        (
            {(1, 1): "A", (2, 2): "B"},
            (Rectangle(1, 1, 1, 2), Rectangle(2, 2, 2, 3)),
            Rectangle(1, 1, 2, 3),
        ),
        ({(1, 1): "A"}, (Rectangle(1, 1, 3, 1),), Rectangle(1, 1, 2, 1)),
    ]:
        band, errors = resolve(cells, merges, window, ("A",))
        assert band is None and errors[0].code is DiagnosticCode.MALFORMED_MULTILEVEL_HEADER


def test_authored_synthetic_collision_is_reported():
    band, errors = resolve(
        {(1, 1): "Store", (2, 1): "col2"},
        (Rectangle(1, 1, 1, 2),),
        Rectangle(1, 1, 2, 2),
        ("Store",),
    )
    assert band is None and errors[0].code is DiagnosticCode.DUPLICATE_MULTILEVEL_HEADER


def test_clipped_ancestor_reports_location_even_when_anchor_is_outside_scope():
    band, errors = resolve(
        {(1, 1): "Root", (2, 2): "Child"},
        (Rectangle(1, 1, 1, 3),),
        Rectangle(1, 2, 3, 3),
        ("Child",),
    )
    assert band is None
    assert errors[0].code is DiagnosticCode.MALFORMED_MULTILEVEL_HEADER
    assert errors[0].sheet == "Report"


def test_huge_clipped_merge_does_not_expand_reads_beyond_window():
    reads = []

    def value_at(row, column):
        reads.append((row, column))
        return "Root" if (row, column) == (1, 1) else None

    band, errors = resolve_band(
        sheet="Report",
        header_window=Rectangle(1, 1, 3, 3),
        merged_ranges=(Rectangle(1, 1, 1, 16384),),
        value_at=value_at,
        query=MultiLevelTableQuery(("Root",), 3),
        checkpoint=lambda: None,
    )
    assert band is None
    assert errors[0].code is DiagnosticCode.MALFORMED_MULTILEVEL_HEADER
    assert len(reads) <= 9
    assert all(1 <= row <= 3 and 1 <= column <= 3 for row, column in reads)


def test_overlapping_merges_are_malformed():
    band, errors = resolve(
        {(1, 1): "Root", (2, 2): "Leaf"},
        (Rectangle(1, 1, 2, 2), Rectangle(2, 2, 3, 3)),
        Rectangle(1, 1, 3, 3),
        ("Root",),
    )
    assert band is None
    assert errors[0].code is DiagnosticCode.MALFORMED_MULTILEVEL_HEADER
