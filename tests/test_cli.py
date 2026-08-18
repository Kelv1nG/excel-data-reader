from __future__ import annotations

import json
from pathlib import Path

from openpyxl import Workbook

from excel_data_reader.cli import main

ROOT = Path(__file__).parents[1]
WORKBOOKS = ROOT / "examples" / "workbooks"


def test_inspect_command_prints_workbook_structure(capsys) -> None:
    exit_code = main(["inspect", str(WORKBOOKS / "native_table.xlsx")])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "Sheets (1):" in captured.out
    assert "OrdersTable" in captured.out


def test_inspect_command_supports_json(capsys) -> None:
    exit_code = main(["inspect", str(WORKBOOKS / "named_and_headerless.xlsx"), "--json"])
    captured = capsys.readouterr()
    payload = json.loads(captured.out)

    assert exit_code == 0
    assert [sheet["name"] for sheet in payload["sheets"]] == ["Inventory", "Raw Import"]
    assert payload["named_ranges"][0]["name"] == "InventoryData"


def test_find_command_outputs_an_explainable_json_report(capsys) -> None:
    exit_code = main(
        [
            "find",
            str(WORKBOOKS / "scattered_headers.xlsx"),
            "--headers",
            "account number,amount",
            "--optional",
            "invoice date,owner,purchase order",
            "--alias",
            "account number=customer id|client no",
            "--sheet",
            "Scattered Orders",
            "--within",
            "A4:G20",
            "--near",
            "A4",
            "--json",
        ]
    )
    captured = capsys.readouterr()
    payload = json.loads(captured.out)

    assert exit_code == 0
    assert payload["scans"][0]["cells_considered"] == 35
    assert payload["candidates"][0]["selected"] is True
    assert len(payload["result"]["matches"]) == 1


def test_find_command_returns_two_when_nothing_matches(capsys) -> None:
    exit_code = main(
        [
            "find",
            str(WORKBOOKS / "scattered_headers.xlsx"),
            "--headers",
            "missing header",
            "--sheet",
            "Scattered Orders",
        ]
    )
    captured = capsys.readouterr()

    assert exit_code == 2
    assert "TABLE_NOT_FOUND" in captured.out


def test_find_command_rejects_malformed_aliases(capsys) -> None:
    exit_code = main(
        [
            "find",
            str(WORKBOOKS / "scattered_headers.xlsx"),
            "--headers",
            "customer id",
            "--alias",
            "broken",
        ]
    )
    captured = capsys.readouterr()

    assert exit_code == 2
    assert "expected FIELD=ALIAS|ALIAS" in captured.err


def test_find_command_returns_three_for_ambiguity(tmp_path: Path, capsys) -> None:
    path = tmp_path / "ambiguous.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["ID", "Amount"])
    sheet.append(["A", 1])
    sheet["A5"] = "ID"
    sheet["B5"] = "Amount"
    sheet["A6"] = "B"
    sheet["B6"] = 2
    workbook.save(path)
    workbook.close()

    exit_code = main(["find", str(path), "--headers", "id,amount"])
    captured = capsys.readouterr()

    assert exit_code == 3
    assert "Matches (2):" in captured.out


def test_find_command_supports_legacy_xls(capsys) -> None:
    exit_code = main(
        [
            "find",
            str(WORKBOOKS / "legacy_scattered.xls"),
            "--headers",
            "customer id,invoice date,amount",
            "--sheet",
            "Legacy Orders",
        ]
    )
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "Legacy Orders!A4:G8" in captured.out
