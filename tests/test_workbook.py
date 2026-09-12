"""Tests for :class:`martes.PandasWorkbook`."""

# pylint: disable=missing-function-docstring,redefined-outer-name

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


@pytest.fixture(scope='module')
def awkward_path(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """A workbook whose sheet names need Excel-style quoting."""
    path = tmp_path_factory.mktemp('martes') / 'awkward.xlsx'
    sheets = {
        'My Sheet': pd.DataFrame([[1, 2], [3, 4]]),
        'Q1!Draft': pd.DataFrame([[5, 6]]),
        "Bob's Data": pd.DataFrame([[7, 8]]),
    }

    with pd.ExcelWriter(path) as writer:
        for name, frame in sheets.items():
            frame.to_excel(
                writer, sheet_name=name, index=False, header=False,
            )

    return path


@pytest.fixture
def awkward(awkward_path: Path) -> PandasWorkbook:
    """A freshly loaded workbook with awkward sheet names."""
    return PandasWorkbook(awkward_path)


def test_unquoted_sheet_name_with_a_space(awkward: PandasWorkbook) -> None:
    assert awkward['My Sheet!A1'] == 1


def test_quoted_sheet_name(awkward: PandasWorkbook) -> None:
    assert awkward["'My Sheet'!A1"] == 1


def test_quoted_sheet_name_containing_a_bang(
    awkward: PandasWorkbook,
) -> None:
    # Unquoted, the split would take the ! inside the sheet name.
    assert awkward["'Q1!Draft'!A1"] == 5


def test_quoted_sheet_name_with_a_doubled_apostrophe(
    awkward: PandasWorkbook,
) -> None:
    assert awkward["'Bob''s Data'!A1"] == 7


def test_quoted_sheet_name_without_a_coordinate(
    awkward: PandasWorkbook,
) -> None:
    result = awkward["'My Sheet'"]

    assert list(result.columns) == ['A', 'B']


def test_quoted_sheet_name_accepts_a_range(
    awkward: PandasWorkbook,
) -> None:
    result = awkward["'My Sheet'!A1:B2"]

    assert result.values.tolist() == [[1, 2], [3, 4]]


def test_setting_through_a_quoted_sheet_name(
    awkward: PandasWorkbook,
) -> None:
    awkward["'My Sheet'!A1"] = 99

    assert awkward["'My Sheet'!A1"] == 99


def test_unterminated_quote_raises(awkward: PandasWorkbook) -> None:
    with pytest.raises(ValueError):
        awkward["'My Sheet!A1"]  # pylint: disable=pointless-statement


def test_unknown_sheet_names_the_available_sheets(
    awkward: PandasWorkbook,
) -> None:
    with pytest.raises(KeyError) as error:
        awkward['Nope!A1']  # pylint: disable=pointless-statement

    assert 'My Sheet' in str(error.value)
    assert 'Nope' in str(error.value)


@pytest.mark.parametrize(
    ('key', 'expected'),
    [
        ('Sheet1!A2', ('Sheet1', 'A2')),
        ('Sheet1', ('Sheet1', None)),
        ("'My Sheet'!A2", ('My Sheet', 'A2')),
        ("'My Sheet'", ('My Sheet', None)),
        ("'Bob''s'!A2", ("Bob's", 'A2')),
        ("'Q1!Draft'!A2", ('Q1!Draft', 'A2')),
    ],
)
def test_split_reference(key: str, expected: tuple[str, str | None]) -> None:
    assert PandasWorkbook.split_reference(key) == expected
