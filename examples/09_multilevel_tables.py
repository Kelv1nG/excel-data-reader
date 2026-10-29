"""Discover merged header paths and independent repeated tables."""

from pathlib import Path

from excel_data_reader import ExcelReader, MultiLevelTableQuery

WORKBOOK = Path(__file__).parent / "workbooks/multilevel_tables.xlsx"


def main() -> None:
    """Demonstrate group expansion, subtree selection, and repeated matches."""

    with ExcelReader.open(WORKBOOK) as reader:
        wide = reader.find_multilevel_tables(
            MultiLevelTableQuery(
                ("col1", "Sales Report"),
                3,
                sheet="Wide Store",
            )
        ).require_one()
        print("Wide Store:", dict(reader.extract_multilevel_table(wide).records()[0]))
        online = reader.find_multilevel_tables(
            MultiLevelTableQuery(
                (("Sales Report", "Online"),),
                3,
                sheet="Plain Store",
            )
        ).require_one()
        print("Online subtree:", dict(reader.extract_multilevel_table(online).records()[0]))
        matches = reader.find_multilevel_tables(
            MultiLevelTableQuery(
                ("Sales Report",),
                3,
                sheet="Repeated",
            )
        )
        if not matches.is_complete:
            raise RuntimeError("discovery did not complete")
        for index, match in enumerate(matches.matches, 1):
            print(f"Table {index}: {match.range}", reader.extract_multilevel_table(match).values)


if __name__ == "__main__":
    main()
