"""Hold a whole Excel workbook as a dictionary of pandas DataFrames."""

import re
from typing import Any, Final

import pandas as pd

from martes.accessors import ExcelCoordAccessor, xl

QUOTED_REFERENCE: Final[re.Pattern[str]] = re.compile(
    r"^'((?:[^']|'')*)'(?:!(?P<coordinate>.*))?$"
)
"""Matches an Excel reference whose sheet name is quoted."""


class PandasWorkbook:
    """An Excel workbook, read into pandas and indexed by coordinate.

    Every sheet is read with ``header=None`` and then relabelled with
    Excel column letters and 1-based row numbers, so that coordinates
    in the file and coordinates in the frame agree.

    Parameters
    ----------
    path_or_buffer : str, path-like, file-like or pandas.ExcelFile
        Anything :class:`pandas.ExcelFile` accepts.
    kwargs : dict of str to dict, optional
        Extra keyword arguments for :func:`pandas.read_excel`, keyed by
        sheet name.
    """

    def __init__(
            self,
            path_or_buffer: Any,
            kwargs: dict[str, dict[str, Any]] | None = None,
    ) -> None:

        self.path_or_buffer = path_or_buffer
        self.kwargs = kwargs or {}
        self.workbook = pd.ExcelFile(self.path_or_buffer)
        names = [str(name) for name in self.workbook.sheet_names]
        self.frames = {name: self.read_sheet(name) for name in names}

    def read_sheet(self, sheet_name: str) -> pd.DataFrame:
        """Read one sheet and relabel it with Excel coordinates.

        Parameters
        ----------
        sheet_name : str
            The name of the sheet to read.

        Returns
        -------
        pandas.DataFrame
            The sheet, with columns ``A``, ``B``, ... and rows numbered
            from 1.
        """

        frame = pd.read_excel(
            io=self.workbook,
            sheet_name=sheet_name,
            header=None,
            **self.kwargs.get(sheet_name, {}),
        )

        return ExcelCoordAccessor(frame).rename_columns()

    @property
    def sheets(self) -> list[str]:
        """Get the names of the sheets in the workbook.

        Returns
        -------
        list of str
            The sheet names, in the order they appear in the file.
        """

        return list(self.frames.keys())

    def __getitem__(self, key: str) -> Any:

        sheet_name, coordinate = self.split_reference(key)
        frame = self.sheet(sheet_name)

        if coordinate is None:
            return frame

        return xl(frame)[coordinate]

    def __setitem__(self, key: str, value: Any) -> None:

        sheet_name, coordinate = self.split_reference(key)

        if coordinate is None:
            raise ValueError('Can only set value within sheet')

        xl(self.sheet(sheet_name))[coordinate] = value

    def sheet(self, sheet_name: str) -> pd.DataFrame:
        """Look up one sheet by name.

        Parameters
        ----------
        sheet_name : str
            The name of the sheet, unquoted.

        Returns
        -------
        pandas.DataFrame
            The sheet.

        Raises
        ------
        KeyError
            If the workbook has no sheet of that name.
        """

        if sheet_name not in self.frames:
            available = ', '.join(repr(name) for name in self.sheets)
            raise KeyError(
                f'No sheet {sheet_name!r} in this workbook. '
                f'Available sheets: {available}'
            )

        return self.frames[sheet_name]

    @staticmethod
    def split_reference(key: str) -> tuple[str, str | None]:
        """Split a reference into its sheet name and its coordinate.

        Excel quotes a sheet name whenever it contains a space or a
        punctuation mark, and doubles any apostrophe inside it, so
        ``'Bob''s Data'!A1`` is a reference to cell A1 of the sheet
        named ``Bob's Data``. Both quoted and unquoted names are
        accepted.

        Parameters
        ----------
        key : str
            A reference such as ``'Sheet1!A2'``, ``"'My Sheet'!A2"`` or
            a bare sheet name.

        Returns
        -------
        tuple of (str, str or None)
            The sheet name, and the coordinate if the reference names
            one.

        Raises
        ------
        ValueError
            If `key` opens a quoted sheet name but never closes it.
        """

        if not key.startswith("'"):
            sheet_name, separator, coordinate = key.partition('!')
            return (sheet_name, coordinate if separator else None)

        match = QUOTED_REFERENCE.match(key)

        if match is None:
            raise ValueError(f'Unterminated sheet name: "{key}"')

        return (match.group(1).replace("''", "'"), match.group('coordinate'))
