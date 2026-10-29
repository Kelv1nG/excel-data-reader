import runpy
from pathlib import Path

from examples.multilevel_workbook import build_multilevel
from excel_data_reader import ExcelReader, MultiLevelTableQuery


def test_public_example_can_be_regenerated(tmp_path):
    path = build_multilevel(tmp_path)
    with ExcelReader.open(path) as reader:
        wide = reader.find_multilevel_tables(
            MultiLevelTableQuery(("Sales Report",), 3, sheet="Wide Store")
        ).require_one()
        assert wide.columns[-1].name == "Sales Report__Store__col2"
        four = reader.find_multilevel_tables(
            MultiLevelTableQuery(("Sales Report",), 4, sheet="Four Levels")
        ).require_one()
        assert four.columns[0].name == "Sales Report__Region__Online__New"
        repeats = reader.find_multilevel_tables(
            MultiLevelTableQuery(("Sales Report",), 3, sheet="Repeated")
        )
        assert [(m.data_start_row, m.data_end_row) for m in repeats.matches] == [(4, 6), (10, 12)]


def test_multilevel_example_script(capsys):
    script = Path(__file__).parents[1] / "examples/09_multilevel_tables.py"
    runpy.run_path(str(script), run_name="__main__")
    output = capsys.readouterr().out
    assert "Sales Report__Store__col2" in output
    assert "A7:C12" in output
