"""Use matrix discovery with automatic, explicit, and aliased layouts."""

from pathlib import Path

from excel_data_reader import BodyPolicy, ExcelReader, MatrixData, MatrixMatch, MatrixQuery

WORKBOOK = Path(__file__).parent / "workbooks" / "sectioned_matrix.xlsx"


def _extract_section(query: MatrixQuery, section: str) -> tuple[MatrixMatch, MatrixData]:
    """Discover and extract one section for the focused examples below."""

    with ExcelReader.open(WORKBOOK) as reader:
        match = reader.find_matrices(query).require_section(section)
        return match, reader.extract_matrix(match)


def usual_case() -> None:
    """Automatically infer headers, identifier column, and section boundaries."""

    query = MatrixQuery(
        sections=("Country Identifier", "Sector Identifier"),
        header_level_names=("group", "attribute"),
        sheet="Sectioned Matrix",
    )

    with ExcelReader.open(WORKBOOK) as reader:
        matches = reader.find_matrices(query)

    print("Automatic matrix discovery:")
    for section in query.sections:
        match = matches.require_section(section)
        print(section, match.range, match.boundary_source)


def use_explicit_layout_hints() -> None:
    """Supply header rows and identifier column when inference would be ambiguous."""

    query = MatrixQuery(
        sections=("Country Identifier",),
        header_level_names=("group", "attribute"),
        sheet="Sectioned Matrix",
        within="A1:L8",
        header_rows=(1, 4),
        identifier_column="D",
    )
    match, data = _extract_section(query, "Country Identifier")

    print("Explicit matrix layout hints:")
    print("header rows:", match.header_rows, "identifier column:", match.identifier_column)
    print("range:", match.range, "values:", len(data.values))


def use_a_logical_section_alias() -> None:
    """Map an authored section label to the application's preferred name."""

    query = MatrixQuery(
        sections=("countries",),
        aliases={"countries": ("Country Identifier",)},
        header_level_names=("group", "attribute"),
        sheet="Sectioned Matrix",
    )
    match, data = _extract_section(query, "countries")

    print("Aliased matrix section:")
    print("logical:", match.section, "authored:", match.raw_section)
    print("record section:", data.long_records()[0]["section"])


def choose_a_body_policy_for_an_unmerged_section() -> None:
    """Use selected-column population to bound an unmerged section."""

    query = MatrixQuery(
        sections=("Sector Identifier",),
        header_level_names=("group", "attribute"),
        sheet="Sectioned Matrix",
        within="A1:L18",
        body=BodyPolicy.last_populated(),
    )
    match, _ = _extract_section(query, "Sector Identifier")

    print("Last-populated matrix body:", match.range, match.boundary_source)


def produce_long_records() -> None:
    """Produce stable analytical records with one row per matrix value."""

    query = MatrixQuery(("Country Identifier",), sheet="Sectioned Matrix")
    _, data = _extract_section(query, "Country Identifier")
    records = data.long_records()

    print("Long matrix records:", len(records))
    print(dict(records[0]))


def produce_wide_records() -> None:
    """Produce one row per identifier with flattened grouped headers."""

    query = MatrixQuery(("Sector Identifier",), sheet="Sectioned Matrix")
    _, data = _extract_section(query, "Sector Identifier")
    records = data.wide_records(separator="__")

    print("Wide matrix records:", len(records))
    print(dict(records[0]))


def main() -> None:
    """Run every sectioned-matrix usage example."""

    usual_case()
    use_explicit_layout_hints()
    use_a_logical_section_alias()
    choose_a_body_policy_for_an_unmerged_section()
    produce_long_records()
    produce_wide_records()


if __name__ == "__main__":
    main()
