from __future__ import annotations

import json
import os
from collections.abc import Mapping
from pathlib import Path
from typing import Any

import pytest

from excel_data_reader import (
    BodyPolicy,
    BodyPolicyMode,
    ExcelReader,
    MatrixData,
    MatrixQuery,
    TableData,
    TableQuery,
)

SEED_MANIFEST = Path(__file__).parent / "acceptance" / "manifest.json"


def _load_cases() -> tuple[tuple[Path, Mapping[str, Any]], ...]:
    manifests = [SEED_MANIFEST]
    if local_manifest := os.environ.get("EXCEL_DATA_READER_ACCEPTANCE_MANIFEST"):
        manifests.append(Path(local_manifest))

    cases: list[tuple[Path, Mapping[str, Any]]] = []
    for manifest in manifests:
        payload = json.loads(manifest.read_text(encoding="utf-8"))
        if payload.get("schema_version") != 1:
            raise ValueError(f"unsupported acceptance manifest schema: {manifest}")
        cases.extend((manifest, case) for case in payload["cases"])
    return tuple(cases)


def _body(value: Mapping[str, Any] | None) -> BodyPolicy:
    if value is None:
        return BodyPolicy()
    mode = BodyPolicyMode(value["mode"])
    if mode is BodyPolicyMode.BLANK_ROWS:
        return BodyPolicy.until_blank_rows(int(value.get("blank_rows", 2)))
    if mode is BodyPolicyMode.LAST_POPULATED:
        return BodyPolicy.last_populated()
    return BodyPolicy.through_row(int(value["bottom_row"]))


def _execute(reader: ExcelReader, case: Mapping[str, Any]) -> TableData | MatrixData:
    operation = case["operation"]
    if operation == "native_table":
        return reader.get_table(str(case["name"]))
    if operation == "named_range":
        return reader.get_named_range(str(case["name"]))
    if operation == "range":
        return reader.read_range(
            str(case["sheet"]),
            str(case["range"]),
            header=case.get("header", 0),
        )
    if operation == "query":
        value = case["query"]
        query = TableQuery(
            required_headers=tuple(value["required_headers"]),
            optional_headers=tuple(value.get("optional_headers", ())),
            aliases=value.get("aliases", {}),
            sheet=value.get("sheet"),
            allow_non_adjacent_columns=value.get("allow_non_adjacent_columns", True),
            body=_body(value.get("body")),
            near=value.get("near"),
            within=value.get("within"),
        )
        return reader.extract(reader.query_tables(query).require_one())
    if operation == "matrix":
        value = case["query"]
        query = MatrixQuery(
            sections=tuple(value["sections"]),
            header_level_names=tuple(value.get("header_level_names", ("group", "attribute"))),
            aliases=value.get("aliases", {}),
            sheet=value.get("sheet"),
            within=value.get("within"),
            body=_body(value.get("body")),
            header_rows=value.get("header_rows"),
            identifier_column=value.get("identifier_column"),
        )
        match = reader.find_matrices(query).require_section(str(case["section"]))
        return reader.extract_matrix(match)
    raise ValueError(f"unknown acceptance operation: {operation!r}")


@pytest.mark.parametrize(
    ("manifest", "case"),
    _load_cases(),
    ids=lambda value: value.get("id", "manifest") if isinstance(value, Mapping) else None,
)
def test_workbook_acceptance_case(manifest: Path, case: Mapping[str, Any]) -> None:
    workbook = (manifest.parent / str(case["workbook"])).resolve()
    expected = case["expected"]

    with ExcelReader.open(workbook) as reader:
        table = _execute(reader, case)

    if isinstance(table, MatrixData):
        assert table.match.section == expected["section"]
        assert table.match.range == expected["range"]
        assert table.match.boundary_source.value == expected["boundary_source"]
        assert table.match.identifier_column == expected["identifier_column"]
        assert list(table.match.header_rows) == expected["header_rows"]
        assert [list(header.labels) for header in table.match.headers] == expected["headers"]
        cells = {item.cell.address: item.cell.value for item in table.values}
        for address, value in expected["cells"].items():
            assert cells[address] == value
        return

    assert table.match.source.value == expected["source"]
    assert table.match.range == expected["range"]
    assert len(table.rows) == expected["row_count"]
    assert [column.source_column for column in table.columns] == expected["source_columns"]
    if "requested_headers" in expected:
        assert [column.requested_header for column in table.columns] == expected[
            "requested_headers"
        ]

    cells = {cell.address: cell.value for row in table.rows for cell in row.cells}
    for address, value in expected["cells"].items():
        assert cells[address] == value
