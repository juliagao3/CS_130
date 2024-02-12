import re

from typing import Tuple

location_regex = re.compile("(\$?)([A-Za-z]+)(\$?)([0-9]+)")

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

class Reference:

    def __init__(self, col: int, row: int, abs_col: bool = False, abs_row: bool = False):
        if col <= 0 or col > from_base_26("zzzz") or row <= 0 or row > 9999:
            raise ValueError

        self.abs_col = abs_col
        self.abs_row = abs_row
        self.col = col
        self.row = row

    def from_string(location_string: str):
        m = location_regex.fullmatch(location_string)

        if m is None:
            raise KeyError

        groups = m.groups()

        abs_col = groups[0] == "$"
        abs_row = groups[2] == "$"
        col = from_base_26(groups[1].lower())
        row = int(groups[3])

        return Reference(col, row, abs_col, abs_row)

    def moved(self, offset: Tuple[int, int]):
        col = self.col
        row = self.row
        if not self.abs_col:
            col += offset[0]
        if not self.abs_row:
            row += offset[1]
        return Reference(col, row, self.abs_col, self.abs_row)

    def tuple(self) -> Tuple[int, int]:
        return (self.col, self.row)

    def __str__(self) -> str:
        def qm(b: bool) -> str:
            return "?" if b else ""
        return "{}{}{}{}".format(qm(self.abs_col), to_base_26(self.col), qm(self.abs_row), self.row)
