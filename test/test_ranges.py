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
        wb.set_cell_contents(n, "C1", '=MIN(INDIRECT("A1:A2"))')

        self.assertEqual(wb.get_cell_value(n, "C1"), decimal.Decimal(10))
        
        i, new_sheet = wb.new_sheet("B1")

        #wb.set_cell_contents(new_sheet, "E1", "=IF(7<6, 1+1, 2+2)")

        for i in range(1, 10):
            wb.set_cell_contents(new_sheet, f"E{i}", f"={i}")
            wb.set_cell_contents(new_sheet, f"F{i}", f'=E{i}')
        wb.set_cell_contents(new_sheet, "D1", f'=INDIRECT("{new_sheet}" & "!F1:F9")')

        #self.assertEqual(wb.get_cell_value(new_sheet, "F1"), decimal.Decimal(1))
        self.assertEqual(wb.get_cell_value(new_sheet, "D1"), decimal.Decimal(1))
        
    def test_vlookup(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        for i in range(1, 10 + 1):
            wb.set_cell_contents(n, f"A{i}", f"={i}")
            wb.set_cell_contents(n, f"B{i}", f"={i**2}")
            wb.set_cell_contents(n, f"C{i}", f"={i**3}")
        wb.set_cell_contents(n, "D1", "=VLOOKUP(5, A1:C10, 3)")

        self.assertEqual(wb.get_cell_value(n, "D1"), decimal.Decimal(125))

    def test_vlookup_out_of_bounds(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        for i in range(1, 10 + 1):
            wb.set_cell_contents(n, f"A{i}", f"={i}")
            wb.set_cell_contents(n, f"B{i}", f"={i**2}")
            wb.set_cell_contents(n, f"C{i}", f"={i**3}")
        wb.set_cell_contents(n, "D1", "=VLOOKUP(5, A1:C10, 4)")

        self.assertIsInstance(wb.get_cell_value(n, "D1"), sheets.CellError)
        self.assertEqual(wb.get_cell_value(n, "D1").get_type(), sheets.CellErrorType.TYPE_ERROR)

    def test_vlookup_circular(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        for i in range(1, 10 + 1):
            wb.set_cell_contents(n, f"A{i}", f"={i}")
            wb.set_cell_contents(n, f"B{i}", f"={i**2}")
            wb.set_cell_contents(n, f"C{i}", f"={i**3}")
        wb.set_cell_contents(n, "D1", "=VLOOKUP(5, A1:D10, 4)")

        self.assertIsInstance(wb.get_cell_value(n, "D1"), sheets.CellError)
        self.assertEqual(wb.get_cell_value(n, "D1").get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        wb.set_cell_contents(n, "E1", f'=IFERROR(VLOOKUP(5, INDIRECT("{n}" & "!A1:C10"), 3), "")')
        self.assertEqual(wb.get_cell_value(n, "E1"), decimal.Decimal(125))

    def test_hlookup(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        for i in range(1, 10 + 1):
            wb.set_cell_contents(n, f"{chr(i + ord('A'))}1", f"={i}")
            wb.set_cell_contents(n, f"{chr(i + ord('A'))}2", f"={i**2}")
            wb.set_cell_contents(n, f"{chr(i + ord('A'))}3", f"={i**3}")
        wb.set_cell_contents(n, "A4", "=HLOOKUP(5, A1:J3, 3)")

        self.assertEqual(wb.get_cell_value(n, "A4"), decimal.Decimal(125))

    def test_hlookup_cycle(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        for i in range(1, 10 + 1):
            wb.set_cell_contents(n, f"{chr(i + ord('A'))}1", f"={i}")
            wb.set_cell_contents(n, f"{chr(i + ord('A'))}2", f"={i**2}")
            wb.set_cell_contents(n, f"{chr(i + ord('A'))}3", f"={i**3}")
        wb.set_cell_contents(n, "A4", "=HLOOKUP(5, A1:J4, 4)")

        self.assertIsInstance(wb.get_cell_value(n, "A4"), sheets.CellError)
        self.assertEqual(wb.get_cell_value(n, "A4").get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

    def test_hlookup_out_of_bounds(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        for i in range(1, 10 + 1):
            wb.set_cell_contents(n, f"{chr(i + ord('A'))}1", f"={i}")
            wb.set_cell_contents(n, f"{chr(i + ord('A'))}2", f"={i**2}")
            wb.set_cell_contents(n, f"{chr(i + ord('A'))}3", f"={i**3}")
        wb.set_cell_contents(n, "A4", "=HLOOKUP(5, A1:J3, 4)")

        self.assertIsInstance(wb.get_cell_value(n, "A4"), sheets.CellError)
        self.assertEqual(wb.get_cell_value(n, "A4").get_type(), sheets.CellErrorType.TYPE_ERROR)

    def test_update_on_copy_sheet(self):
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.set_cell_contents(n, 'A1', '1')
        wb.set_cell_contents(n, 'A2', '2')

        j, m = wb.new_sheet()

        wb.set_cell_contents(m, 'A1', f'=SUM({n}_1!A1:B2)')

        k, o = wb.copy_sheet(n)

        self.assertEqual(wb.get_cell_value(m, "A1"), decimal.Decimal(3))


if __name__ == "__main__":
        unittest.main()
