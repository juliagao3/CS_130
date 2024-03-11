#! /usr/bin/env python3
import unittest
import unittest.mock

import sheets
import decimal

class TestClass(unittest.TestCase):

    def test_range_parse(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.set_cell_contents(n, "A1", '=B1:B10')
        wb.set_cell_contents(n, "B1", 'HI')
        wb.set_cell_contents(n, "B2", 'Hello')

        self.assertEqual(wb.get_cell_value(n, "A1"), "HI")

    def test_range_min(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        for i in range(1,11):
            wb.set_cell_contents(n, f"B{i}", f"{i}")

        wb.set_cell_contents(n, "A1", '=MIN(B1:B9, B10)')

        self.assertEqual(wb.get_cell_value(n, "A1"), decimal.Decimal(1))

    def test_range_max(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        for i in range(1,11):
            wb.set_cell_contents(n, f"B{i}", f"{i}")

        wb.set_cell_contents(n, "A1", '=MAX(B1:B10)')

        self.assertEqual(wb.get_cell_value(n, "A1"), decimal.Decimal(10))

    def test_range_average(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        for i in range(1,11):
            wb.set_cell_contents(n, f"B{i}", f"{i}")

        wb.set_cell_contents(n, "A1", '=AVERAGE(B1:B10)')

        self.assertEqual(wb.get_cell_value(n, "A1"), decimal.Decimal("5.5"))

    def test_range_sum(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        for i in range(1,11):
            wb.set_cell_contents(n, f"B{i}", f"{i}")

        wb.set_cell_contents(n, "A1", '=SUM(B1:B10)')

        self.assertEqual(wb.get_cell_value(n, "A1"), decimal.Decimal("55"))

if __name__ == "__main__":
        unittest.main()
