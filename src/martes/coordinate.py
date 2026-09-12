"""Represent an Excel coordinate as a column letter and a row number."""

import string
from collections.abc import Iterator
from functools import reduce
from typing import Final, NamedTuple

MAX_COLUMN_INDEX: Final[int] = 18278
"""The index of column ``ZZZ``, the last column this module supports."""


class Coordinate(NamedTuple):
    """An Excel coordinate, split into its column and row parts.

    Either part may be ``None``, which is how whole-column references
    such as ``'A:A'`` and whole-row references such as ``'1:1'`` are
    represented.

    Attributes
    ----------
    column : str or None
        The column letters, upper-cased, or ``None`` if the coordinate
        names no column.
    row : int or None
        The 1-based row number, or ``None`` if the coordinate names no
        row.
    """

    column: str | None
    row: int | None

    @classmethod
    def from_str(
            cls,
            coordinate: str,
    ) -> 'Coordinate | tuple[Coordinate, Coordinate]':
        """Parse a coordinate naming either a single cell or a range.

        Parameters
        ----------
        coordinate : str
            A coordinate such as ``'A2'``, ``'A2:G50'`` or ``'A:A'``.

        Returns
        -------
        Coordinate or tuple of Coordinate
            A single coordinate, or the two endpoints of a range.
        """

        parts = cls.split_parts(coordinate)

        if len(parts) == 2:
            return (cls.from_cell(parts[0]), cls.from_cell(parts[1]))

        return cls.from_cell(parts[0])

    @staticmethod
    def split_parts(coordinate: str) -> list[str]:
        """Split a coordinate string on its colon, if it has one.

        Parameters
        ----------
        coordinate : str
            A coordinate such as ``'A2'`` or ``'A2:G50'``.

        Returns
        -------
        list of str
            One part for a cell reference, two for a range.

        Raises
        ------
        TypeError
            If `coordinate` is not a string.
        ValueError
            If `coordinate` contains more than one colon.
        """

        if not isinstance(coordinate, str):
            raise TypeError(f'Expected `str`, received `{type(coordinate)}`')

        parts = coordinate.split(':')

        if len(parts) > 2:
            raise ValueError(f'Invalid coordinate: "{coordinate}"')

        return parts

    @classmethod
    def from_cell(cls, coordinate: str) -> 'Coordinate':
        """Parse a single-cell coordinate string.

        Parameters
        ----------
        coordinate : str
            A coordinate naming at most one cell, such as ``'A2'``,
            ``'A'`` or ``'2'``.

        Returns
        -------
        Coordinate
            The parsed coordinate.
        """

        index = cls.first_digit(coordinate)
        column = coordinate[:index].upper() or None
        row = int(coordinate[index:]) if index is not None else None

        return cls(column=column, row=row)

    @staticmethod
    def first_digit(coordinate: str) -> int | None:
        """Find the position at which the row number begins.

        Parameters
        ----------
        coordinate : str
            A coordinate naming at most one cell.

        Returns
        -------
        int or None
            The index of the first digit, or ``None`` if there is none.
        """

        for index, character in enumerate(coordinate):
            if character in string.digits:
                return index

        return None

    @property
    def column_index(self) -> int | None:
        """Get the 1-based number of this coordinate's column.

        Returns
        -------
        int or None
            The column number, or ``None`` if this coordinate names no
            column.
        """

        return self.get_column_index(self.column)

    @staticmethod
    def get_column_letter(col_idx: int) -> str:
        """Convert a column number into a column letter (3 -> ``'C'``).

        Parameters
        ----------
        col_idx : int
            A 1-based column number, up to `MAX_COLUMN_INDEX`.

        Returns
        -------
        str
            The corresponding column letters.

        Raises
        ------
        ValueError
            If `col_idx` falls outside the range of allowed columns.
        """

        if not 1 <= col_idx <= MAX_COLUMN_INDEX:
            raise ValueError(f'Invalid column index {col_idx}')

        return ''.join(reversed(list(Coordinate.iter_letters(col_idx))))

    @staticmethod
    def iter_letters(col_idx: int) -> Iterator[str]:
        """Yield a column's letters, least significant first.

        Excel columns are bijective base 26, so there is no zero digit
        and each place is taken from ``1..26`` rather than ``0..25``.

        Parameters
        ----------
        col_idx : int
            A 1-based column number.

        Yields
        ------
        str
            One letter of the column, from the right-hand end.
        """

        while col_idx > 0:
            col_idx, remainder = divmod(col_idx - 1, 26)
            yield chr(remainder + ord('A'))

    @staticmethod
    def get_column_index(col_letter: str | None) -> int | None:
        """Convert a column letter into a column number (``'C'`` -> 3).

        Parameters
        ----------
        col_letter : str or None
            The column letters, upper-cased.

        Returns
        -------
        int or None
            The 1-based column number, or ``None`` if `col_letter` is
            ``None``.

        Raises
        ------
        ValueError
            If `col_letter` contains anything but upper-case letters.
        """

        if col_letter is None:
            return None

        if not all(letter in string.ascii_uppercase for letter in col_letter):
            raise ValueError(f'Invalid column "{col_letter}"')

        return reduce(Coordinate.fold_letter, col_letter, 0)

    @staticmethod
    def fold_letter(index: int, letter: str) -> int:
        """Add one letter to a running bijective base-26 total.

        Parameters
        ----------
        index : int
            The total accumulated from the letters so far.
        letter : str
            The next column letter, upper-cased.

        Returns
        -------
        int
            The running total, with `letter` folded in.
        """

        return index * 26 + (ord(letter) - ord('A') + 1)
