"""Tests for :class:`martes.Coordinate`."""

# pylint: disable=missing-function-docstring

from __future__ import annotations

import pytest

from martes import Coordinate


def test_from_str_single_cell() -> None:
    assert Coordinate.from_str('A2') == Coordinate(column='A', row=2)


def test_from_str_upper_cases_the_column() -> None:
    assert Coordinate.from_str('bc7') == Coordinate(column='BC', row=7)


def test_from_str_column_only() -> None:
    assert Coordinate.from_str('A') == Coordinate(column='A', row=None)


def test_from_str_row_only() -> None:
    assert Coordinate.from_str('12') == Coordinate(column=None, row=12)


def test_from_str_range_returns_both_ends() -> None:
    assert Coordinate.from_str('A2:G50') == (
        Coordinate(column='A', row=2),
        Coordinate(column='G', row=50),
    )


def test_from_str_whole_column_range() -> None:
    assert Coordinate.from_str('A:A') == (
        Coordinate(column='A', row=None),
        Coordinate(column='A', row=None),
    )


def test_from_str_rejects_non_strings() -> None:
    with pytest.raises(TypeError):
        Coordinate.from_str(2)  # type: ignore[arg-type]


def test_from_str_rejects_two_colons() -> None:
    with pytest.raises(ValueError):
        Coordinate.from_str('A1:B2:C3')


@pytest.mark.parametrize(
    ('index', 'letter'),
    [
        (1, 'A'),
        (26, 'Z'),
        (27, 'AA'),
        (52, 'AZ'),
        (53, 'BA'),
        (702, 'ZZ'),
        (703, 'AAA'),
        (18278, 'ZZZ'),
    ],
)
def test_get_column_letter(index: int, letter: str) -> None:
    assert Coordinate.get_column_letter(index) == letter


@pytest.mark.parametrize('index', [1, 26, 27, 702, 703, 18278])
def test_column_letter_and_index_round_trip(index: int) -> None:
    letter = Coordinate.get_column_letter(index)

    assert Coordinate.get_column_index(letter) == index


@pytest.mark.parametrize('index', [0, -1, 18279])
def test_get_column_letter_rejects_out_of_range(index: int) -> None:
    with pytest.raises(ValueError):
        Coordinate.get_column_letter(index)


def test_get_column_index_of_none_is_none() -> None:
    assert Coordinate.get_column_index(None) is None


def test_get_column_index_rejects_non_letters() -> None:
    with pytest.raises(ValueError):
        Coordinate.get_column_index('A1')


def test_column_index_property() -> None:
    assert Coordinate.from_cell('C4').column_index == 3


def test_column_index_property_of_a_rowless_coordinate() -> None:
    assert Coordinate.from_cell('7').column_index is None


@pytest.mark.parametrize(
    ('coordinate', 'expected'),
    [('A2', 1), ('AB12', 2), ('2', 0), ('A', None)],
)
def test_first_digit(coordinate: str, expected: int | None) -> None:
    assert Coordinate.first_digit(coordinate) == expected
