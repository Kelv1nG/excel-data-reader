from __future__ import annotations

import pytest

from excel_data_reader import Coordinate, Rectangle


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
