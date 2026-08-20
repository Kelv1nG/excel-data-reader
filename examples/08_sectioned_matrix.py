"""Discover merged and unmerged matrix sections beneath repeated grouped headers."""

from pathlib import Path

from excel_data_reader import ExcelReader, MatrixQuery

workbook_path = Path(__file__).parent / "workbooks" / "sectioned_matrix.xlsx"
query = MatrixQuery(
    sections=("Country Identifier", "Sector Identifier"),
    header_level_names=("group", "attribute"),
    sheet="Sectioned Matrix",
)

with ExcelReader.open(workbook_path) as reader:
    matches = reader.find_matrices(query)
    for section in query.sections:
        match = matches.require_section(section)
        data = reader.extract_matrix(match)
        print(
            section,
            match.range,
            match.boundary_source,
            "long rows:",
            len(data.long_records()),
        )
        print("first long record:", dict(data.long_records()[0]))
        print("first wide record:", dict(data.wide_records()[0]))
