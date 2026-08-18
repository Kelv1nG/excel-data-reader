from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from examples.build_workbooks import build_all
from excel_data_reader import ExcelReader, MatchSource

ROOT = Path(__file__).parents[1]
WORKBOOKS = ROOT / "examples" / "workbooks"


def test_example_generator_builds_all_workbooks(tmp_path: Path) -> None:
    paths = build_all(tmp_path)

    assert [path.name for path in paths] == [
        "native_table.xlsx",
        "scattered_headers.xlsx",
        "named_and_headerless.xlsx",
        "legacy_scattered.xls",
    ]
    assert all(path.exists() for path in paths)


def test_checked_in_native_table_example() -> None:
    with ExcelReader.open(WORKBOOKS / "native_table.xlsx") as reader:
        table = reader.get_table("OrdersTable")

    assert table.match.source is MatchSource.NATIVE_TABLE
    assert table.match.range == "A4:F9"
    assert len(table.rows) == 4
    assert table.rows[0].cells[5].value == "=D5*E5"


def test_checked_in_scattered_header_example() -> None:
    with ExcelReader.open(WORKBOOKS / "scattered_headers.xlsx") as reader:
        match = reader.find_tables(
            ["amount", "customer id", "invoice date"],
            sheet="Scattered Orders",
            max_blank_rows=2,
        ).require_one()
        table = reader.extract(match)

    assert match.range == "A4:G8"
    assert [column.source_column for column in match.columns] == [7, 1, 4]
    assert table.rows[2].values == (None, None, None)


def test_checked_in_named_and_headerless_example() -> None:
    with ExcelReader.open(WORKBOOKS / "named_and_headerless.xlsx") as reader:
        inventory = reader.get_named_range("InventoryData")
        raw = reader.read_range("Raw Import", "C5:F8", header=None)

    assert len(inventory.rows) == 4
    assert [column.name for column in raw.columns] == [
        "column_1",
        "column_2",
        "column_3",
        "column_4",
    ]
    assert raw.rows[0].cells[0].address == "C5"


def test_checked_in_legacy_xls_example() -> None:
    with ExcelReader.open(WORKBOOKS / "legacy_scattered.xls") as reader:
        match = reader.find_tables(
            ["customer id", "invoice date", "amount"],
            sheet="Legacy Orders",
        ).require_one()

    assert match.range == "A4:G8"
    assert [column.source_column for column in match.columns] == [1, 4, 7]


def test_all_example_scripts_run() -> None:
    for script in (
        "01_native_table.py",
        "02_scattered_headers.py",
        "03_named_and_headerless.py",
        "04_table_query.py",
        "05_explain_query.py",
        "06_platform_upload.py",
        "07_legacy_xls.py",
    ):
        completed = subprocess.run(
            [sys.executable, str(ROOT / "examples" / script)],
            cwd=ROOT,
            check=True,
            capture_output=True,
            text=True,
        )
        assert completed.stdout.strip()
