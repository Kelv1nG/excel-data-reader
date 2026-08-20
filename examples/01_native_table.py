"""Use an authored native Excel Table in several common ways."""

from pathlib import Path

from excel_data_reader import ExcelReader, FormulaValue, ValueMode

WORKBOOK = Path(__file__).parent / "workbooks" / "native_table.xlsx"


def list_available_tables() -> None:
    """Inspect native table names and authored ranges before choosing one."""

    with ExcelReader.open(WORKBOOK) as reader:
        matches = reader.find_native_tables(sheet="Orders")

    print("Available native tables:")
    for match in matches.matches:
        print(match.name, match.sheet, match.range)


def usual_case() -> None:
    """Extract a known native table by name and return mapping-style rows."""

    with ExcelReader.open(WORKBOOK) as reader:
        table = reader.get_table("OrdersTable")

    print("Native table records:")
    for record in table.records():
        print(dict(record))


def select_columns_in_logical_order() -> None:
    """Project selected native-table columns in caller-requested order."""

    with ExcelReader.open(WORKBOOK) as reader:
        match = reader.find_tables(
            ("amount", "customer", "order id"),
            sheet="Orders",
        ).require_one()
        table = reader.extract(match)

    print("Projected native table:")
    print("source:", match.source, "columns:", [column.name for column in table.columns])
    print("first row:", table.rows[0].values)


def read_formula_and_cached_value() -> None:
    """Read both formula text and Excel's last stored cached result."""

    with ExcelReader.open(WORKBOOK, value_mode=ValueMode.BOTH) as reader:
        table = reader.get_table("OrdersTable")

    amount = table.rows[0].cells[-1].value
    print("Formula and cached value:")
    if isinstance(amount, FormulaValue):
        print("formula:", amount.formula, "cached:", amount.cached)
    else:
        print(amount)


def main() -> None:
    """Run every native-table usage example."""

    list_available_tables()
    usual_case()
    select_columns_in_logical_order()
    read_formula_and_cached_value()


if __name__ == "__main__":
    main()
