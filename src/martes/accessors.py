"""Access Excel ranges on a pandas DataFrame through an ``.xl`` accessor."""

from __future__ import annotations

from typing import Any

import pandas as pd

from martes.coordinate import Coordinate


@pd.api.extensions.register_dataframe_accessor('xl')
class ExcelCoordAccessor:
    """Index a DataFrame by Excel coordinate.

    Registered by importing :mod:`martes`, after which every DataFrame
    exposes it as ``.xl``.

    Parameters
    ----------
    pandas_obj : pandas.DataFrame
        The DataFrame this accessor wraps.
    """

    def __init__(self, pandas_obj: pd.DataFrame) -> None:

        self._obj = pandas_obj

    def __getitem__(self, key: str) -> Any:

        if key not in self:
            self.expand(self._obj, key)

        return self._obj.loc[self._range(key)]

    def __setitem__(self, key: str, value: Any) -> None:

        if key not in self:
            self.expand(self._obj, key)

        self._obj.loc[self._range(key)] = value

    def __contains__(self, key: str) -> bool:

        furthest = Coordinate.from_cell(key.rsplit(':', 1)[-1])

        return (
            (
                furthest.column is None
                or furthest.column in self._obj.columns
            )
            and (
                furthest.row is None
                or furthest.row in self._obj.index
            )
        )

    def __call__(self) -> pd.DataFrame:

        return self.rename_columns()

    def rename_columns(self) -> pd.DataFrame:
        """Relabel the axes with Excel column letters and row numbers.

        Returns
        -------
        pandas.DataFrame
            A copy of the frame with columns ``A``, ``B``, ... and rows
            numbered from 1.
        """

        return (self._obj
            .rename_axis('rows', axis='index')
            .rename_axis('columns', axis='columns')
            .set_axis(range(len(self._obj.columns)), axis='columns')
            .reset_index(drop=True)
            .rename(
                columns=lambda c: Coordinate.get_column_letter(c + 1),
                index=lambda r: r + 1,
            )
        )

    @classmethod
    def expand(cls, frame: pd.DataFrame, to_key: str) -> None:
        """Grow a frame in place until it reaches a coordinate.

        New cells are filled with ``NaN``. Existing cells are left
        alone.

        Parameters
        ----------
        frame : pandas.DataFrame
            The frame to grow. Modified in place.
        to_key : str
            The coordinate the frame must be large enough to contain.
        """

        furthest = Coordinate.from_cell(to_key.rsplit(':', 1)[-1])

        cls.expand_columns(frame, to_index=furthest.column_index)
        cls.expand_rows(frame, to_row=furthest.row)

    @staticmethod
    def expand_columns(frame: pd.DataFrame, to_index: int | None) -> None:
        """Append empty columns to a frame in place.

        Parameters
        ----------
        frame : pandas.DataFrame
            The frame to grow. Modified in place.
        to_index : int or None
            The 1-based column number to grow to. ``None`` is a no-op.
        """

        if to_index is None:
            return

        for col_num in range(frame.shape[1] + 1, to_index + 1):
            frame[Coordinate.get_column_letter(col_num)] = float('NaN')

    @staticmethod
    def expand_rows(frame: pd.DataFrame, to_row: int | None) -> None:
        """Append empty rows to a frame in place.

        Parameters
        ----------
        frame : pandas.DataFrame
            The frame to grow. Modified in place.
        to_row : int or None
            The 1-based row number to grow to. ``None`` is a no-op.
        """

        if to_row is None:
            return

        for row in range(frame.shape[0] + 1, to_row + 1):
            frame.loc[row] = float('NaN')

    @classmethod
    def _range(
            cls,
            key: str,
    ) -> tuple[int, str] | tuple[slice, slice]:
        """Turn a coordinate string into something ``.loc`` understands.

        Parameters
        ----------
        key : str
            A coordinate such as ``'A2'`` or ``'A2:G50'``.

        Returns
        -------
        tuple
            A ``(row, column)`` pair for a single cell, or a pair of
            slices for a range.
        """

        coordinate = Coordinate.from_str(key)

        if isinstance(coordinate, Coordinate):
            return cls._range_index(coordinate)

        return cls._range_slice(coordinate)

    @classmethod
    def _range_index(cls, coordinate: Coordinate) -> tuple[int, str]:
        """Turn a single coordinate into a ``(row, column)`` pair.

        Parameters
        ----------
        coordinate : Coordinate
            A coordinate naming exactly one cell.

        Returns
        -------
        tuple of (int, str)
            The row number and column letters.

        Raises
        ------
        ValueError
            If `coordinate` is missing its row or its column, and so
            names a range rather than a cell.
        """

        if coordinate.row is None or coordinate.column is None:
            msg = 'Invalid coordinate: "{col}{row}"'
            raise ValueError(
                msg.format(
                    col=coordinate.column or '',
                    row=coordinate.row or '',
                )
            )

        return (coordinate.row, coordinate.column)

    @classmethod
    def _range_slice(
            cls,
            coordinates: tuple[Coordinate, Coordinate],
    ) -> tuple[slice, slice]:
        """Turn a pair of coordinates into a pair of ``.loc`` slices.

        Parameters
        ----------
        coordinates : tuple of Coordinate
            The top-left and bottom-right ends of the range.

        Returns
        -------
        tuple of (slice, slice)
            A slice of rows and a slice of columns.
        """

        top_left, bottom_right = coordinates

        return (
            slice(top_left.row, bottom_right.row),
            slice(top_left.column, bottom_right.column),
        )
