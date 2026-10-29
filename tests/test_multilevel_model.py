from dataclasses import replace

import pytest

from excel_data_reader import (
    Coordinate,
    Diagnostic,
    DiagnosticCode,
    ExcelDataReaderError,
    MultiLevelColumn,
    MultiLevelHeaderPart,
    MultiLevelTableData,
    MultiLevelTableMatch,
    MultiLevelTableMatchSet,
    MultiLevelTableQuery,
    Rectangle,
)


def part(label, row=1, column=1):
    return MultiLevelHeaderPart(
        label, label, Coordinate(row, column), Rectangle(row, column, row, column)
    )


def match(columns=()):
    return MultiLevelTableMatch("Report", Rectangle(1, 1, 3, 3), columns, 4, 3)


@pytest.mark.parametrize(
    "depth,error", [(0, ValueError), (-1, ValueError), (True, TypeError), (2.0, TypeError)]
)
def test_depth_validation(depth, error):
    with pytest.raises(error):
        MultiLevelTableQuery(("Store",), depth)


@pytest.mark.parametrize("headers", [(), (" ",), (("", "Store"),), (("Store",), 3)])
def test_empty_or_invalid_selectors_rejected(headers):
    with pytest.raises((TypeError, ValueError)):
        MultiLevelTableQuery(headers, 3)


def test_query_freezes_selectors():
    source = ["Store", ("Sales", "Online")]
    query = MultiLevelTableQuery(source, 3)
    source.append("Later")
    assert query.required_headers == ("Store", ("Sales", "Online"))


def test_paths_and_empty_bounds():
    column = MultiLevelColumn(
        3,
        (part("Sales Report"), part("Store", 2, 3), MultiLevelHeaderPart("col2", None, None, None)),
    )
    assert column.path == ("Sales Report", "Store", "col2")
    assert column.name == "Sales Report__Store__col2"
    assert column.parts[-1].is_synthetic
    value = match((column,))
    assert value.is_empty
    assert value.bounds == value.header_bounds
    assert MultiLevelTableData(value, ()).records() == ()


def test_require_one_respects_incomplete_and_ambiguous_diagnostics():
    single = match()
    assert MultiLevelTableMatchSet((single,)).require_one() == single
    for result, code in [
        (MultiLevelTableMatchSet(()), DiagnosticCode.TABLE_NOT_FOUND),
        (MultiLevelTableMatchSet((single, single)), DiagnosticCode.AMBIGUOUS_TABLE),
        (MultiLevelTableMatchSet((single,), is_complete=False), DiagnosticCode.SCAN_LIMIT_EXCEEDED),
        (
            MultiLevelTableMatchSet(
                (single,), (Diagnostic(DiagnosticCode.AMBIGUOUS_MULTILEVEL_HEADER, "duplicate"),)
            ),
            DiagnosticCode.AMBIGUOUS_MULTILEVEL_HEADER,
        ),
    ]:
        with pytest.raises(ExcelDataReaderError) as exc:
            result.require_one()
        assert exc.value.diagnostics[0].code == code


@pytest.mark.parametrize(
    "first,second",
    [
        (("A__B", "C"), ("A", "B__C")),
        (("Store", "New"), ("store", "new")),
        (("Store", "col2"), ("Store", "col2")),
    ],
)
def test_records_refuse_colliding_names(first, second):
    columns = tuple(
        MultiLevelColumn(i, tuple(part(label, j, i) for j, label in enumerate(path, 1)))
        for i, path in enumerate((first, second), 1)
    )
    with pytest.raises(ExcelDataReaderError, match="DUPLICATE_MULTILEVEL_HEADER"):
        MultiLevelTableData(replace(match(), columns=columns), ()).records()
