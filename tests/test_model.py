from __future__ import annotations

import pytest

from excel_data_reader import BodyPolicy, BodyPolicyMode, Coordinate, Rectangle, TableQuery


def test_coordinate_and_rectangle_a1_notation() -> None:
    assert Coordinate(7, 28).a1 == "AB7"
    assert Rectangle(2, 2, 4, 5).a1 == "B2:E4"
    assert Rectangle(2, 2, 2, 2).a1 == "B2"


@pytest.mark.parametrize(
    "bounds",
    [
        (0, 1, 1, 1),
        (2, 1, 1, 1),
        (1, 3, 1, 2),
    ],
)
def test_rectangle_rejects_invalid_bounds(bounds: tuple[int, int, int, int]) -> None:
    with pytest.raises(ValueError):
        Rectangle(*bounds)


def test_rectangle_containment() -> None:
    outer = Rectangle(1, 2, 10, 8)

    assert outer.contains_rectangle(Rectangle(2, 3, 9, 7))
    assert not outer.contains_rectangle(Rectangle(1, 1, 9, 7))


def test_body_policy_factories() -> None:
    assert BodyPolicy.until_blank_rows(3).blank_rows == 3
    assert BodyPolicy.last_populated().mode is BodyPolicyMode.LAST_POPULATED
    assert BodyPolicy.through_row(12).bottom_row == 12

    with pytest.raises(ValueError):
        BodyPolicy.until_blank_rows(0)


def test_table_query_freezes_input_collections() -> None:
    aliases = {"id": ["client id"]}
    query = TableQuery(["id"], aliases=aliases)
    aliases["id"].append("account id")

    assert query.required_headers == ("id",)
    assert query.aliases["id"] == ("client id",)
