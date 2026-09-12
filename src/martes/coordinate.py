"""Class to represent an excel coordinate"""

from typing import NamedTuple
import string

class Coordinate(NamedTuple):
    """Represents an Excel coordinate"""

    column: str
    row: str

    @classmethod
    def from_str(cls, coordinate: str) -> 'Coordinate':
        """Creates a Coordinate object from a string"""

        if not isinstance(coordinate, str):
            raise TypeError(f'Expected `str`, received `{type(coordinate)}`')

        if len(coordinate.split(':')) == 2:
            output = tuple(cls.from_str(c) for c in coordinate.split(':', 1))
        elif len(coordinate.split(':')) > 2:
            raise ValueError(f'Invalid coordinate: "{coordinate}"')
        else:
            for index, character in enumerate(coordinate):
                if character in string.digits:
                    break
            else:
                index = None

            column = coordinate[:index].upper() or None
            row = int(coordinate[index:]) if index is not None else None
            output =  Coordinate(column=column, row=row)

        return output

    @property
    def column_index(self):
        """Return the column number"""

        return self.get_column_index(self.column)

    @staticmethod
    def get_column_letter(col_idx: int) -> str:
        """Convert a column number into a column letter (3 -> 'C')

        Right shift the column col_idx by 26 to find column letters in reverse
        order.  These numbers are 1-based, and can be converted to ASCII
        ordinals by adding 64.
        """

        # Indicies correspond to A -> ZZZ and include all allowed columns
        if not 1 <= col_idx <= 18278:
            raise ValueError(f'Invalid column index {col_idx}')

        letters = []

        while col_idx > 0:

            col_idx, remainder = divmod(col_idx, 26)

            # check for exact division and borrow if needed
            if remainder == 0:
                remainder = 26
                col_idx -= 1

            letters.append(chr(remainder+64))

        return ''.join(reversed(letters))

    @staticmethod
    def get_column_index(col_letter: str) -> int:
        """Convert a column letter into a column number ('C' -> 3)"""

        if col_letter is None:
            return None

        if not all(letter in string.ascii_uppercase for letter in col_letter):
            raise ValueError(f'Invalid column "{col_letter}"')

        index = 0
        for position, letter in enumerate(reversed(col_letter)):
            index += (26**position) * (ord(letter)-64)

        return index
