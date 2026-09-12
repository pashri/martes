"""Tests for :class:`martes.PandasWorkbook`."""

# pylint: disable=missing-function-docstring,redefined-outer-name

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from martes import PandasWorkbook


@pytest.fixture(scope='module')
def workbook_path(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """A two-sheet workbook with no header row."""
    path = tmp_path_factory.mktemp('martes') / 'book.xlsx'
    sheets = {
        'Addresses': pd.DataFrame(
            [['Name', 'Street'],
             ['Ada', '1 Pi Lane'],
             ['Bob', '2 Tau Road']]
        ),
        'Numbers': pd.DataFrame([[1, 2], [3, 4]]),
    }

    with pd.ExcelWriter(path) as writer:
        for name, frame in sheets.items():
            frame.to_excel(
                writer, sheet_name=name, index=False, header=False,
            )

    return path


@pytest.fixture
def workbook(workbook_path: Path) -> PandasWorkbook:
    """A freshly loaded workbook, so mutations do not leak between tests."""
    return PandasWorkbook(workbook_path)


def test_sheets_lists_every_sheet(workbook: PandasWorkbook) -> None:
    assert workbook.sheets == ['Addresses', 'Numbers']


def test_getitem_sheet_name_returns_a_frame(
    workbook: PandasWorkbook,
) -> None:
    result = workbook['Addresses']

    assert list(result.columns) == ['A', 'B']
    assert list(result.index) == [1, 2, 3]


def test_getitem_qualified_cell(workbook: PandasWorkbook) -> None:
    assert workbook['Addresses!A2'] == 'Ada'


def test_getitem_qualified_range(workbook: PandasWorkbook) -> None:
    result = workbook['Addresses!A2:B3']

    assert result.values.tolist() == [
        ['Ada', '1 Pi Lane'],
        ['Bob', '2 Tau Road'],
    ]


def test_getitem_qualified_column(workbook: PandasWorkbook) -> None:
    result = workbook['Addresses!A:A']

    assert result.values.ravel().tolist() == ['Name', 'Ada', 'Bob']


def test_accessor_on_a_retrieved_sheet(workbook: PandasWorkbook) -> None:
    assert workbook['Numbers'].xl['B2'] == 4


def test_setitem_qualified_cell(workbook: PandasWorkbook) -> None:
    workbook['Numbers!A1'] = 100

    assert workbook['Numbers!A1'] == 100


def test_setitem_without_a_sheet_name_raises(
    workbook: PandasWorkbook,
) -> None:
    with pytest.raises(ValueError):
        workbook['Numbers'] = 1


def test_read_kwargs_are_passed_through(workbook_path: Path) -> None:
    result = PandasWorkbook(workbook_path, kwargs={'Numbers': {'nrows': 1}})

    assert result['Numbers'].shape == (1, 2)


def test_unknown_sheet_raises(workbook: PandasWorkbook) -> None:
    with pytest.raises(KeyError):
        workbook['Nope']  # pylint: disable=pointless-statement
