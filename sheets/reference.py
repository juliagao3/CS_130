import re
from typing import Tuple, Optional

from . import base_types

location_regex = re.compile("(([A-Za-z_][A-Za-z0-9_]*|'[^']*')!)?(\$?)([A-Za-z]+)(\$?)([0-9]+)")

def from_base_26(s: str):
    result = 0
    for c in s:
        result *= 26
        result += ord(c) - ord('a') + 1
    return result

def to_base_26(index: int) -> str:
    def divmod_excel(i: int) -> Tuple[int, int]:
        a, b = divmod(i, 26)
        if b == 0:
            return a - 1, b + 26
        return a, b
    result = []
    while index > 0:
        index, rem = divmod_excel(index)
        result.append(chr(ord("a") + rem - 1))
    return "".join(reversed(result))

def unquote(s: Optional[str]) -> Optional[str]:
        if s is None:
            return None
        return s[1:-1] if s[0] == "'" else s

class Reference:

    MAX_COL = from_base_26("zzzz")

    def __init__(self,
                 sheet_name: Optional[str],
                 col: int,
                 row: int,
                 abs_col: bool = False,
                 abs_row: bool = False):
        self.sheet_name = sheet_name
        self.abs_col = abs_col
        self.abs_row = abs_row
        self.col = col
        self.row = row
    
    def max(a, b):
        assert a.sheet_name == b.sheet_name or (a.sheet_name is None and b.sheet_name is None)
        return Reference(a.sheet_name, max(a.col, b.col), max(a.row, b.row))

    def min(a, b):
        assert a.sheet_name == b.sheet_name or (a.sheet_name is None and b.sheet_name is None)
        return Reference(a.sheet_name, min(a.col, b.col), min(a.row, b.row))

    def check_bounds(self):
        if self.col <= 0 or self.col > Reference.MAX_COL or self.row <= 0 or self.row > 9999:
            raise ValueError
        return self

    def check_absolute(self):
        if self.abs_col or self.abs_row:
            raise ValueError
        return self

    def from_string(default_sheet_name: Optional[str], location_string: str):
        if location_string is None:
            raise ValueError

        m = location_regex.fullmatch(location_string)

        if m is None:
            raise ValueError

        groups = m.groups()

        sheet_name = unquote(groups[1])
        abs_col = groups[2] == "$"
        abs_row = groups[4] == "$"

        if sheet_name is None:
            sheet_name = default_sheet_name

        col = from_base_26(groups[3].lower())
        row = int(groups[5])

        return Reference(sheet_name, col, row, abs_col, abs_row)

    def moved(self, offset: Tuple[int, int]):
        sheet_name = self.sheet_name
        col = self.col
        row = self.row
        if not self.abs_col:
            col += offset[0]
        if not self.abs_row:
            row += offset[1]
        return Reference(sheet_name, col, row, self.abs_col, self.abs_row)

    def tuple(self) -> Tuple[int, int]:
        return (self.col, self.row)

    def location_string(self) -> str:
        def ds(b: bool) -> str:
            return "$" if b else ""
        return "{}{}{}{}".format(ds(self.abs_col), to_base_26(self.col), ds(self.abs_row), self.row)
        
    def __str__(self) -> str:
        def f(s: Optional[str]) -> str:
            if s is None:
                return ""
            elif base_types.sheet_name_needs_quotes(s):
                return "'" + s + "'!"
            else:
                return s + "!"

        return "{}{}".format(f(self.sheet_name), self.location_string())
