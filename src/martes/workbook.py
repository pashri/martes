"""Hold a whole Excel workbook as a dictionary of pandas DataFrames."""

from __future__ import annotations

from typing import Any

import pandas as pd

from martes.accessors import ExcelCoordAccessor


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

        if len(key.split('!', 1)) == 2:
            sheet_name, coordinate = key.split('!', 1)
            return self.frames[sheet_name].xl[coordinate]

        return self.frames[key]

    def __setitem__(self, key: str, value: Any) -> None:

        if len(key.split('!', 1)) != 2:
            raise ValueError('Can only set value within sheet')

        sheet_name, coordinate = key.split('!', 1)
        self.frames[sheet_name].xl[coordinate] = value
