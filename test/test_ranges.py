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

    def test_range_cycle(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        for i in range(2, 10):
            wb.set_cell_contents(n, f"B{i}", f"{i}")
        
        wb.set_cell_contents(n, "B1", '=SUM(B2:B11)')
        wb.set_cell_contents(n, "B9", '=B1')

        self.assertIsInstance(wb.get_cell_value(n, "B1"), sheets.CellError)
        self.assertEqual(wb.get_cell_value(n, "B1").get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        self.assertIsInstance(wb.get_cell_value(n, "B9"), sheets.CellError)
        self.assertEqual(wb.get_cell_value(n, "B9").get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

    def test_range_if(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.set_cell_contents(n, "A1", "=5>6")
        for i in range(1, 5):
            wb.set_cell_contents(n, f"B{i}", f"{i}")
            wb.set_cell_contents(n, f"C{i}", f"{i + 1}")
        
        wb.set_cell_contents(n, "D1", '=SUM(IF(A1, B1:B4, C1:C4))')
        
        self.assertEqual(wb.get_cell_value(n, "D1"), decimal.Decimal(14))

        wb.set_cell_contents(n, "A1", "=7>6")
        self.assertEqual(wb.get_cell_value(n, "D1"), decimal.Decimal(10))

    def test_range_choose(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()
        
        wb.set_cell_contents(n, "A1", "10")
        wb.set_cell_contents(n, "A2", "5")

        wb.set_cell_contents(n, "B1", "=SUM(A1:A2)")
        wb.set_cell_contents(n, "B2", "=AVERAGE(A1:A2)")
        
        wb.set_cell_contents(n, "C1", "=CHOOSE(1, B1, B2)")
        
        self.assertEqual(wb.get_cell_value(n, "C1"), decimal.Decimal(15))
        
        wb.set_cell_contents(n, "D1", '=SUM(CHOOSE(2, A1:A2, B1:B2))')
        
        self.assertEqual(wb.get_cell_value(n, "D1"), decimal.Decimal(22.5))
    
    def test_range_indirect(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.set_cell_contents(n, "A1", "10")
        wb.set_cell_contents(n, "A2", "20")
        wb.set_cell_contents(n, "B1", "")
        pass

if __name__ == "__main__":
        unittest.main()
