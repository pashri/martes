"""Accessors for accessing Excel ranges in pandas"""

from typing import Any, Tuple, Union
import pandas as pd
from martes.coordinate import Coordinate

@pd.api.extensions.register_dataframe_accessor('xl')
class ExcelCoordAccessor:
    """A class for accessing Excel ranges in pandas"""

    def __init__(self, pandas_obj):

        self._obj = pandas_obj

    def __getitem__(self, key: str) -> Union[pd.DataFrame, Any]:

        if key not in self:
            self._obj.pipe(self.expand, key)

        return self._obj.loc[self._range(key)]

    def __setitem__(self, key: str, value: Any) -> None:

        if key not in self:
            self._obj.pipe(self.expand, key)

        self._obj.loc[self._range(key)] = value

    def __contains__(self, key: str) -> bool:

        furthest_coordinate = Coordinate.from_str(key.rsplit(':', 1)[-1])

        return (
            (
                furthest_coordinate.column is None
                or furthest_coordinate.column in self._obj.columns
            )
            and (
                furthest_coordinate.row is None
                or furthest_coordinate.row in self._obj.index
            )
        )

    def __call__(self):

        return self.rename_columns()

    def rename_columns(self):
        """Renames columns to excel format"""

        return (self._obj
            .rename_axis('rows', axis='index')
            .rename_axis('columns', axis='columns')
            .set_axis(range(len(self._obj.columns)), axis='columns')
            .reset_index(drop=True)
            .rename(
                columns=lambda c: Coordinate.get_column_letter(c+1),
                index=lambda r: r+1,
            )
        )

    @staticmethod
    def expand(frame: pd.DataFrame, to_key: str):
        """Expand the DataFrame to a certain size"""

        nrows, ncols = frame.shape
        furthest_coordinate = Coordinate.from_str(to_key.rsplit(':', 1)[-1])

        if furthest_coordinate.column_index is not None:
            for col_num in range(ncols, furthest_coordinate.column_index+1):
                col = Coordinate.get_column_letter(col_num)
                frame[col] = float('NaN')

        if furthest_coordinate.row > nrows:
            for row in range(nrows+1, furthest_coordinate.row+1):
                frame.loc[row] = float('NaN')

    @classmethod
    def _range(cls, key: str) -> Union[Tuple[int, str], Tuple[slice, slice]]:
        """Returns a slice or tuple of slices, given a coordinate string"""

        coordinate = Coordinate.from_str(key)

        if isinstance(coordinate, Coordinate):
            _range = cls._range_index(coordinate)
        elif isinstance(coordinate, tuple):
            _range = cls._range_slice(coordinate)

        return _range

    @classmethod
    def _range_index(cls, coordinate: str) -> Tuple[int, str]:
        """Returns an index and column, given a coordinate string"""

        if coordinate.row is None or coordinate.column is None:
            msg = 'Invalid coordinate: "{col}{row}"'
            raise ValueError(
                msg.format(
                    col=coordinate.column or '',
                    row=coordinate.row or '',
                )
            )

        _range = (coordinate.row, coordinate.column)

        return _range

    @classmethod
    def _range_slice(
            cls,
            coordinates: Tuple[Coordinate, Coordinate],
    ) -> Tuple[slice, slice]:
        """Returns a slice of indexes and a slice of columns, as a tuple,
        given a coordinate string
        """

        top_left, bottom_right = coordinates

        _range = (
            slice(top_left.row, bottom_right.row),
            slice(top_left.column, bottom_right.column),
        )

        return _range
