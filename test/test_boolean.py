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

    def test_literals(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.set_cell_contents(n, "A1", "faLse")
        self.assertEqual(wb.get_cell_value(n, "A1"), False)

        wb.set_cell_contents(n, "A1", "trUe")
        self.assertEqual(wb.get_cell_value(n, "A1"), True)

        wb.set_cell_contents(n, "A1", "'trUe")
        self.assertEqual(wb.get_cell_value(n, "A1"), "trUe")

    def test_formula_literals(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.set_cell_contents(n, "A1", "=faLse")
        self.assertEqual(wb.get_cell_value(n, "A1"), False)
        self.assertIsInstance(wb.get_cell_value(n, "A1"), bool)

        wb.set_cell_contents(n, "A1", "=trUe")
        self.assertEqual(wb.get_cell_value(n, "A1"), True)
        self.assertIsInstance(wb.get_cell_value(n, "A1"), bool)

    def test_literal_cmp(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.set_cell_contents(n, "A1", "=1 < 2")
        self.assertEqual(wb.get_cell_value(n, "A1"), True)

        wb.set_cell_contents(n, "A1", "=1 > 2")
        self.assertEqual(wb.get_cell_value(n, "A1"), False)

        wb.set_cell_contents(n, "A1", "=2 > 1")
        self.assertEqual(wb.get_cell_value(n, "A1"), True)

        wb.set_cell_contents(n, "A1", "=1 = 1")
        self.assertEqual(wb.get_cell_value(n, "A1"), True)

        wb.set_cell_contents(n, "A1", "=1 != 2")
        self.assertEqual(wb.get_cell_value(n, "A1"), True)

    def test_cell_cmp(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.set_cell_contents(n, "A1", "1")
        wb.set_cell_contents(n, "A2", "2")

        wb.set_cell_contents(n, "A3", "=a2 > a1")
        self.assertEqual(wb.get_cell_value(n, "A3"), True)

        wb.set_cell_contents(n, "A3", "=a1 = a1")
        self.assertEqual(wb.get_cell_value(n, "A3"), True)

        wb.set_cell_contents(n, "A3", "=a1 != a2")
        self.assertEqual(wb.get_cell_value(n, "A3"), True)

        wb.set_cell_contents(n, "A3", "=b1 = b2")
        self.assertEqual(wb.get_cell_value(n, "A3"), True)

    def test_precedence(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.set_cell_contents(n, "A2", "1")
        wb.set_cell_contents(n, "B2", "2")

        wb.set_cell_contents(n, "A1", '=a2 = b2 & " type"')
        self.assertEqual(wb.get_cell_value(n, "A1"), False)

    def test_unequal_strings(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.set_cell_contents(n, "A1", '="hello" <> "Hello"')
        self.assertEqual(wb.get_cell_value(n, "A1"), False)

        wb.set_cell_contents(n, "A1", '="hello" > "asdf"')
        self.assertEqual(wb.get_cell_value(n, "A1"), True)

        wb.set_cell_contents(n, "A1", '="hello" > "Hello"')
        self.assertEqual(wb.get_cell_value(n, "A1"), False)

        wb.set_cell_contents(n, "A1", '="hello" < "Hello"')
        self.assertEqual(wb.get_cell_value(n, "A1"), False)

    def test_compare_strings(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.set_cell_contents(n, "A1", '="hello" <> "Hello"')
        self.assertEqual(wb.get_cell_value(n, "A1"), False)

        wb.set_cell_contents(n, "A1", '="a" < "[" ')
        self.assertEqual(wb.get_cell_value(n, "A1"), False)

        wb.set_cell_contents(n, "A1", '="hello" = "Hello"')
        self.assertEqual(wb.get_cell_value(n, "A1"), True)

        wb.set_cell_contents(n, "A1", '=false < true')
        self.assertEqual(wb.get_cell_value(n, "A1"), True)

        wb.set_cell_contents(n, "A1", '=false > true')
        self.assertEqual(wb.get_cell_value(n, "A1"), False)

        wb.set_cell_contents(n, "A1", '="true" < True')
        self.assertEqual(wb.get_cell_value(n, "A1"), True)

        wb.set_cell_contents(n, "A1", '=0 < True')
        self.assertEqual(wb.get_cell_value(n, "A1"), True)

        wb.set_cell_contents(n, "A1", '="true" > 0')
        self.assertEqual(wb.get_cell_value(n, "A1"), True)

        wb.set_cell_contents(n, "A1", '="0" == 0')
        self.assertEqual(wb.get_cell_value(n, "A1"), False)

        wb.set_cell_contents(n, "A1", '="true" == true')
        self.assertEqual(wb.get_cell_value(n, "A1"), False)
        
    def test_empty_cells(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.set_cell_contents(n, "A1", "=b1 == 0")
        self.assertEqual(wb.get_cell_value(n, "A1"), True)

        wb.set_cell_contents(n, "A1", '=b1 == ""')
        self.assertEqual(wb.get_cell_value(n, "A1"), True)

        wb.set_cell_contents(n, "A1", '=b1 > "hi"')
        self.assertEqual(wb.get_cell_value(n, "A1"), False)

        wb.set_cell_contents(n, "A1", '=b1 == False')
        self.assertEqual(wb.get_cell_value(n, "A1"), True)

    def test_error_propagation(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.set_cell_contents(n, "B1", "=B1")
        wb.set_cell_contents(n, "B2", "=10+")

        wb.set_cell_contents(n, "A1", '=b1 == b2')
        self.assertIsInstance(wb.get_cell_value(n, "A1"), sheets.CellError)
        self.assertEqual(wb.get_cell_value(n, "A1").get_type(), sheets.CellErrorType.PARSE_ERROR)

    def test_implicit_conversion(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.set_cell_contents(n, "A1", '=True + 1')
        self.assertEqual(wb.get_cell_value(n, "A1"), decimal.Decimal("2"))

        wb.set_cell_contents(n, "A1", '=False + 1')
        self.assertEqual(wb.get_cell_value(n, "A1"), decimal.Decimal("1"))

        wb.set_cell_contents(n, "A1", '=True')
        wb.set_cell_contents(n, "A2", '=A1 & "Hello"')
        self.assertEqual(wb.get_cell_value(n, "A2"), "TRUEHello")

        wb.set_cell_contents(n, "A1", '=FaLSe & "Hello"')
        self.assertEqual(wb.get_cell_value(n, "A2"), "FALSEHello")        

if __name__ == "__main__":
        unittest.main()
