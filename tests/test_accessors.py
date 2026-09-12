"""Tests for the ``.xl`` DataFrame accessor."""

# pylint: disable=missing-function-docstring,redefined-outer-name

from __future__ import annotations

import pandas as pd
import pytest

import martes  # noqa: F401  # pylint: disable=unused-import


@pytest.fixture
def frame() -> pd.DataFrame:
    """A 3x3 grid already labelled with Excel coordinates."""
    return pd.DataFrame(
        [[1, 2, 3], [4, 5, 6], [7, 8, 9]],
        index=[1, 2, 3],
        columns=['A', 'B', 'C'],
    )


def test_getitem_single_cell_returns_a_scalar(frame: pd.DataFrame) -> None:
    assert frame.xl['B2'] == 5


def test_getitem_range_returns_a_frame(frame: pd.DataFrame) -> None:
    result = frame.xl['A1:B2']

    assert result.shape == (2, 2)
    assert result.values.tolist() == [[1, 2], [4, 5]]


def test_getitem_whole_column(frame: pd.DataFrame) -> None:
    result = frame.xl['B:B']

    assert result.values.ravel().tolist() == [2, 5, 8]


def test_getitem_is_case_insensitive(frame: pd.DataFrame) -> None:
    assert frame.xl['b2'] == 5


def test_getitem_rejects_a_coordinate_with_no_row(
    frame: pd.DataFrame,
) -> None:
    with pytest.raises(ValueError):
        frame.xl['B']  # pylint: disable=pointless-statement


def test_setitem_single_cell(frame: pd.DataFrame) -> None:
    frame.xl['B2'] = 50

    assert frame.xl['B2'] == 50


def test_setitem_range(frame: pd.DataFrame) -> None:
    frame.xl['A1:B2'] = 0

    assert frame.xl['A1:B2'].values.tolist() == [[0, 0], [0, 0]]


def test_contains_an_existing_cell(frame: pd.DataFrame) -> None:
    assert 'C3' in frame.xl


def test_does_not_contain_a_cell_past_the_edge(frame: pd.DataFrame) -> None:
    assert 'D4' not in frame.xl


def test_setitem_past_the_edge_expands_the_frame(
    frame: pd.DataFrame,
) -> None:
    frame.xl['E5'] = 99

    assert frame.shape == (5, 5)
    assert frame.xl['E5'] == 99


def test_expanding_preserves_the_existing_cells(
    frame: pd.DataFrame,
) -> None:
    # Regression: the old expand() started at the last existing column
    # and overwrote it with NaN.
    frame.xl['E5'] = 99

    assert frame.xl['C3'] == 9
    assert frame.xl['A1'] == 1


def test_expanding_fills_new_cells_with_nan(frame: pd.DataFrame) -> None:
    frame.xl['E5'] = 99

    assert pd.isna(frame.xl['D4'])


def test_reading_past_the_edge_expands_the_frame(
    frame: pd.DataFrame,
) -> None:
    assert pd.isna(frame.xl['D4'])
    assert frame.shape == (4, 4)


def test_expanding_to_a_whole_column_does_not_raise(
    frame: pd.DataFrame,
) -> None:
    # Regression: a rowless coordinate used to be compared against an
    # int, raising TypeError.
    result = frame.xl['E:E']

    assert result.isna().all().all()


def test_rename_columns_relabels_a_plain_frame() -> None:
    plain = pd.DataFrame([[1, 2], [3, 4]])

    result = plain.xl()  # type: ignore[operator]

    assert list(result.columns) == ['A', 'B']
    assert list(result.index) == [1, 2]


def test_calling_the_accessor_is_rename_columns() -> None:
    plain = pd.DataFrame([[1, 2], [3, 4]])

    assert plain.xl().equals(  # type: ignore[operator]
        plain.xl.rename_columns()
    )


def test_expanding_to_a_whole_row_adds_rows_only(
    frame: pd.DataFrame,
) -> None:
    result = frame.xl['5:5']

    assert list(frame.columns) == ['A', 'B', 'C']
    assert result.isna().all().all()
