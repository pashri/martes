"""A class for holding Excel workbooks in pandas"""

from typing import Any, Dict, List, Union
import pandas as pd

class PandasWorkbook:
    """A class for holding Excel workbooks in pandas"""

    def __init__(
            self,
            path_or_buffer,
            kwargs: Dict[str, Dict[str, Any]] = None,
    ):

        self.path_or_buffer = path_or_buffer
        self.kwargs = kwargs or {}
        self.workbook = pd.ExcelFile(self.path_or_buffer)
        self.frames = {}

        for sheet_name in self.workbook.sheet_names:

            self.frames[sheet_name] = pd.read_excel(
                io=self.workbook,
                sheet_name=sheet_name,
                header=None,
                **self.kwargs.get(sheet_name, {}),
            ).xl()

    @property
    def sheets(self) -> List[str]:
        """The names of the sheets in the workbook"""

        return list(self.frames.keys())

    def __getitem__(self, key: str) -> Union[pd.DataFrame, Any]:

        if len(key.split('!', 1)) == 2:
            sheet_name, coordinate = key.split('!', 1)
            output = self.frames[sheet_name].xl[coordinate]
        else:
            output = self.frames[key]

        return output

    def __setitem__(self, key: str, value: Any):

        if len(key.split('!', 1)) == 2:
            sheet_name, coordinate = key.split('!', 1)
            self.frames[sheet_name].xl[coordinate] = value
        else:
            raise ValueError('Can only set value within sheet')
