#! /usr/bin/env python3
import unittest
import unittest.mock

import sys
import sheets
import decimal
import traceback
import re
import io

from typing import Tuple

def to_excel_column(index: int) -> str:
    def divmod_excel(i: int) -> Tuple[int, int]:
        a, b = divmod(i, 26)
        if b == 0:
            return a - 1, b + 26
        return a, b
    result = []
    while index > 0:
        index, rem = divmod_excel(index)
        result.append(chr(ord("A") + rem - 1))
    return "".join(reversed(result))

def from_excel_column(s: str) -> int:
    s = s.upper()
    result = 0
    for c in s:
        result *= 26
        result += ord(c) - ord('A') + 1
    return result

def to_sheet_location(pos: Tuple[int, int]) -> str:
    return "{}{}".format(to_excel_column(pos[0]), pos[1])

def from_sheet_location(s: str) -> Tuple[int, int]:
    i = 0
    while not s[i].isdigit():
        i += 1
    return (from_excel_column(s[:i]), int(s[i:]))

class TestClass(unittest.TestCase):

    def test_sort_reversed_reversed(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        row_count = 10

        for row in range(1, row_count + 1):
            wb.set_cell_contents(n, "A" + str(row), f"={row}")

        wb.sort_region(n, "A1", f"A{row_count}", [-1])

        for new_row in range(1, row_count + 1):
            self.assertEqual(wb.get_cell_value(n, "A" + str(new_row)), decimal.Decimal(row_count+1-new_row))

    def test_sort_reversed(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        row_count = 10

        for row in range(1, row_count + 1):
            wb.set_cell_contents(n, "A" + str(row), f"={row_count+1-row}")

        wb.sort_region(n, "A1", f"A{row_count}", [1])

        for new_row in range(1, row_count + 1):
            self.assertEqual(wb.get_cell_value(n, "A" + str(new_row)), decimal.Decimal(new_row))

    def test_stable(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        for row in range(1, 5):
            wb.set_cell_contents(n, "A" + str(row), f"{row}")
            wb.set_cell_contents(n, "B" + str(row), f"={row}")

        for row in range(5, 10):
            wb.set_cell_contents(n, "A" + str(row), "5")
            wb.set_cell_contents(n, "B" + str(row), f"={row}")

        for row in range(10, 15 + 1):
            wb.set_cell_contents(n, "A" + str(row), f"{row}")
            wb.set_cell_contents(n, "B" + str(row), f"={row}")

        wb.sort_region(n, "A1", f"B{15}", [1])

        for new_row in range(5, 10):
            self.assertEqual(wb.get_cell_value(n, "B" + str(new_row)), decimal.Decimal(new_row))

    def test_sort_reversed_reversed_bounds(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        row_count = 10

        for row in range(1, row_count + 1):
            wb.set_cell_contents(n, "A" + str(row), f"={row_count+1-row}")

        wb.sort_region(n, f"A{row_count}", "A1", [1])

        for new_row in range(1, row_count + 1):
            self.assertEqual(wb.get_cell_value(n, "A" + str(new_row)), decimal.Decimal(new_row))

    def test_bad_bounds(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        row_count = 10

        for row in range(1, row_count + 1):
            wb.set_cell_contents(n, "ZZZZ" + str(row), f"={row_count+1-row}")

        with self.assertRaises(ValueError):
            wb.sort_region(n, f"ZZZZ{row_count}", "ZZZZA1", [1])

    def test_bad_sort_col(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        row_count = 10

        for row in range(1, row_count + 1):
            wb.set_cell_contents(n, "ZZZZ" + str(row), f"={row_count+1-row}")

        with self.assertRaises(ValueError):
            wb.sort_region(n, f"ZZZZ{row_count}", "ZZZZ1", [0])

    def test_duplicate_sort_col(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        row_count = 10

        for row in range(1, row_count + 1):
            wb.set_cell_contents(n, "ZZZZ" + str(row), f"={row_count+1-row}")

        with self.assertRaises(ValueError):
            wb.sort_region(n, f"ZZZZ{row_count}", "ZZZZ1", [1, -1])

    def test_sort_formulas(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        row_count = 10

        for row in range(1, row_count + 1):
            wb.set_cell_contents(n, "B" + str(row), f"={row}")
            wb.set_cell_contents(n, "A" + str(row), "=B$" + str(row_count + 1 - row))

        wb.sort_region(n, "A1", f"A{row_count}", [1])

        for new_row in range(1, row_count + 1):
            self.assertEqual(wb.get_cell_contents(n, "A" + str(new_row)), f"=b${new_row}")
            #self.assertEqual(wb.get_cell_value(n, "A" + str(new_row)), decimal.Decimal(row_count + 1 - new_row))

    def test_sort_formulas_relative(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        row_count = 10

        for row in range(1, row_count + 1):
            wb.set_cell_contents(n, "B" + str(row), f"={row}")
            wb.set_cell_contents(n, "A" + str(row), f"=B{row}")

        wb.sort_region(n, "A1", f"A{row_count}", [-1])

        for new_row in range(1, row_count + 1):
            self.assertEqual(wb.get_cell_contents(n, "A" + str(new_row)), f"=b{new_row}")
            #self.assertEqual(wb.get_cell_value(n, "A" + str(new_row)), decimal.Decimal(row_count + 1 - new_row))
       
    def test_sort_blank(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        row_count = 10

        for row in range(1, row_count + 1):
            wb.set_cell_contents(n, "A" + str(row), f"={row_count+1-row}")

        wb.sort_region(n, "A1", f"A{row_count}", [1])

        for new_row in range(1, row_count + 1):
            self.assertEqual(wb.get_cell_value(n, "A" + str(new_row)), decimal.Decimal(new_row))

    def test_sort_multiple_cols(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        row_count = 10

        for row in range(1, row_count + 1):
            wb.set_cell_contents(n, "A" + str(row), f"={row_count+1-row}")
            # wb.set_cell_contents(n, "B" + str(row), f"={row_count+1-row}")

        wb.sort_region(n, "A1", f"A{row_count}", [1])

        for new_row in range(1, row_count + 1):
            self.assertEqual(wb.get_cell_value(n, "A" + str(new_row)), decimal.Decimal(new_row))
        pass

if __name__ == "__main__":
        unittest.main()