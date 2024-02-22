#! /usr/bin/env python3
import unittest
import unittest.mock

import sys
import sheets
import decimal
import traceback
import re
import io

class TestClass(unittest.TestCase):

    def test_function_parse(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.set_cell_contents(n, "A2", '=A3')

        wb.set_cell_contents(n, "A1", '=VERSION()')
        self.assertEqual(wb.get_cell_value(n, "A1"), sheets.version)

    def test_function_add(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.set_cell_contents(n, "A1", '=AND(1 < 2, 6 < 5)')
        self.assertEqual(wb.get_cell_value(n, "A1"), False)

        wb.set_cell_contents(n, "A1", '=AND(1 < 2, 4 < 5)')
        self.assertEqual(wb.get_cell_value(n, "A1"), True)

    def test_function_or(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.set_cell_contents(n, "A1", '=OR(1 < 2, 6 < 5)')
        self.assertEqual(wb.get_cell_value(n, "A1"), True)

        wb.set_cell_contents(n, "A1", '=OR(1 < 2, 4 < 5)')
        self.assertEqual(wb.get_cell_value(n, "A1"), True)

        wb.set_cell_contents(n, "A1", '=OR(1 > 2, 4 > 5)')
        self.assertEqual(wb.get_cell_value(n, "A1"), False)

    def test_function_add_or(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.set_cell_contents(n, "A1", "4")
        wb.set_cell_contents(n, "B1", "5")
        wb.set_cell_contents(n, "C1", "7")
        wb.set_cell_contents(n, "D1", "14")
        wb.set_cell_contents(n, "E1", '=OR(AND(A1 > 5, B1 < 2), AND(C1 < 6, D1 = 14))')
        self.assertEqual(wb.get_cell_value(n, "E1"), False)

    def test_function_indirect(self):
        wb = sheets.Workbook()
        _, n = wb.new_sheet()
        _, m = wb.new_sheet()

        wb.set_cell_contents(n, "A1", "'hello")
        wb.set_cell_contents(n, "A2", "A1")
        wb.set_cell_contents(n, "A3", "=INDIRECT(A2)")
        self.assertEqual(wb.get_cell_value(n, "A3"), "hello")

        wb.set_cell_contents(m, "A1", "'world")
        wb.set_cell_contents(n, "A2", f"{m}!A1")
        self.assertEqual(wb.get_cell_value(n, "A3"), "world")

        wb.set_cell_contents(n, "A2", f"invalid")
        self.assertIsInstance(wb.get_cell_value(n, "A3"), sheets.CellError)
        self.assertEqual(wb.get_cell_value(n, "A3").get_type(), sheets.CellErrorType.BAD_REFERENCE)

        wb.set_cell_contents(n, "A2", "A3")
        self.assertIsInstance(wb.get_cell_value(n, "A3"), sheets.CellError)
        self.assertEqual(wb.get_cell_value(n, "A3").get_type(), sheets.CellErrorType.BAD_REFERENCE)

if __name__ == "__main__":
        unittest.main()
