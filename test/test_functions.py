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

    def test_function_and(self):
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

    def test_function_and_or(self):
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
        wb.set_cell_contents(n, "A2", f"A1")
        wb.set_cell_contents(n, "A3", f"=INDIRECT(A2)")
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

        wb.set_cell_contents(n, "B1", "1")
        wb.set_cell_contents(n, "B2", "2")

    # not, xor, exact, if, iferror, choose, isblank, iserror, version
    def test_function_not(self):
        wb = sheets.Workbook()
        _, name1 = wb.new_sheet()
        _, name2 = wb.new_sheet()

        wb.set_cell_contents(name1, "A1", "=10 == 10")
        wb.set_cell_contents(name1, "A2", "=0 > 1")
        # false
        wb.set_cell_contents(name2, "A1", f"=NOT({name1}!A1)")
        # true
        wb.set_cell_contents(name2, "A2", f"=NOT({name1}!A2)")
        # true
        wb.set_cell_contents(name2, "A3", f"=NOT(AND({name1}!A1, {name1}!A2))")
        # false
        wb.set_cell_contents(name2, "A4", f"=NOT(OR({name1}!A1, {name1}!A2))")
    
        self.assertEqual(wb.get_cell_value(name1, "A1"), True)
        self.assertEqual(wb.get_cell_value(name1, "A2"), False)

        self.assertEqual(wb.get_cell_value(name2, "A1"), False)
        self.assertEqual(wb.get_cell_value(name2, "A2"), True)

        self.assertEqual(wb.get_cell_value(name2, "A3"), True)
        self.assertEqual(wb.get_cell_value(name2, "A4"), False)
    
    def test_function_xor(self):
        wb = sheets.Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "A1", "1")
        wb.set_cell_contents(name, "A2", "5")
        wb.set_cell_contents(name, "A3", "10")

        # 1 true
        wb.set_cell_contents(name, "B1", "=XOR(A1 < 4, A2 < 4, A3 < 4)")
        self.assertEqual(wb.get_cell_value(name, "B1"), True)

        # 2 true
        wb.set_cell_contents(name, "B2", "=XOR(A1 < 5, A2 > 5, A3 > 5)")
        self.assertEqual(wb.get_cell_value(name, "B2"), False)
    
    def test_function_exact(self):
        # test exact with errors/wrong types
        wb = sheets.Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "A1", "hello")
        wb.set_cell_contents(name, "A2", "'hello")
        wb.set_cell_contents(name, "A3", "Hello")
        wb.set_cell_contents(name, "A4", "12345")
        wb.set_cell_contents(name, "A5", "#REF!")

        wb.set_cell_contents(name, "B1", "=EXACT(A1, A2)")
        wb.set_cell_contents(name, "B2", "=EXACT(A1, A3)")
        wb.set_cell_contents(name, "B3", "=EXACT(A1, A4)")
        wb.set_cell_contents(name, "B4", "=EXACT(A1, A5)")

        str1 = "string"
        str2 = "string"
        wb.set_cell_contents(name, "B5", "=EXACT({str1}, {str2})")
        self.assertTrue(wb.get_cell_value(name, "B5"))

        self.assertEqual(wb.get_cell_value(name, "B1"), True)
        self.assertEqual(wb.get_cell_value(name, "B2"), False)
        self.assertEqual(wb.get_cell_value(name, "B3"), False)
        self.assertIsInstance(wb.get_cell_value(name, "B4"), sheets.CellError)
        self.assertEqual(wb.get_cell_value(name, "B4").get_type(), sheets.CellErrorType.BAD_REFERENCE)
                
    def test_if(self):
        wb = sheets.Workbook()
        _, name = wb.new_sheet()
        
        wb.set_cell_contents(name, "A1", "True")
        wb.set_cell_contents(name, "A2", "2")
        wb.set_cell_contents(name, "A3", '=IF(A2 < A1, "True > 2", "2 is smaller than True")')
        self.assertEqual(wb.get_cell_value(name, "A3"), "True > 2")
        
        wb.set_cell_contents(name, "A4", '=IF(4 < 2, "4 > 2")')
        self.assertEqual(wb.get_cell_value(name, "A4"), False)

        wb.set_cell_contents(name, "A5", '=IF(4 < 2, "4 > 2", "4 is greater than 2")')
        self.assertEqual(wb.get_cell_value(name, "A5"), "4 is greater than 2")
    
    def test_function_iferror(self):
        wb = sheets.Workbook()
        _, name = wb.new_sheet()
        
        wb.set_cell_contents(name, "A1", "5 > 6")
        wb.set_cell_contents(name, "A2", "#REF!")
        
        # test with 1 arg
        wb.set_cell_contents(name, "B1", "=IFERROR(A1)")
        self.assertEqual(wb.get_cell_value(name, "B1"), "5 > 6")
        
        wb.set_cell_contents(name, "B2", "=IFERROR(A2)")
        self.assertEqual(wb.get_cell_value(name, "B2"), "")
        
        wb.set_cell_contents(name, "C1", "=ISERROR(A1+)")
        self.assertIsInstance(wb.get_cell_value(name, "C1"), sheets.CellError)
        self.assertEqual(wb.get_cell_value(name, "C1").get_type(), sheets.CellErrorType.PARSE_ERROR)
        
        wb.set_cell_contents(name, "C2", "=C1")
        self.assertIsInstance(wb.get_cell_value(name, "C2"), sheets.CellError)
        self.assertEqual(wb.get_cell_value(name, "C2").get_type(), sheets.CellErrorType.PARSE_ERROR)
        
        wb.set_cell_contents(name, "C3", "=ISERROR(C4)")
        wb.set_cell_contents(name, "C4", "=ISERROR(C3)")
        self.assertIsInstance(wb.get_cell_value(name, "C3"), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value(name, "C4"), sheets.CellError)
        self.assertEqual(wb.get_cell_value(name, "C3").get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value(name, "C4").get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        
        # test with 2 args
        wb.set_cell_contents(name, "A3", "dummy value")
        wb.set_cell_contents(name, "D1", "=IFERROR(A1, A3)")
        self.assertEqual(wb.get_cell_value(name, "B1"), "5 > 6")

        wb.set_cell_contents(name, "D2", "=IFERROR(A2, A3)")
        self.assertEqual(wb.get_cell_value(name, "D2"), "dummy value")
        
    # check that isblank passes ruff
    def test_is_blank(self):
        wb = sheets.Workbook()
        _, n = wb.new_sheet()
        
        wb.set_cell_contents(n, "A1", '=ISBLANK("")')
        self.assertEqual(wb.get_cell_value(n, "A1"), False)
        
        wb.set_cell_contents(n, "A2", '=ISBLANK(FALSE)')
        self.assertEqual(wb.get_cell_value(n, "A2"), False)
        
        wb.set_cell_contents(n, "A3", '=ISBLANK(0)')
        self.assertEqual(wb.get_cell_value(n, "A3"), False)
        
        wb.set_cell_contents(n, "A4", "=ISBLANK(A5)")
        self.assertEqual(wb.get_cell_value(n, "A4"), True)

        wb.set_cell_contents(n, "A4", "=ISBLANK(sheet2!A5)")
        self.assertIsInstance(wb.get_cell_value(n, "A4"), sheets.CellError)
        self.assertEqual(wb.get_cell_value(n, "A4").get_type(), sheets.CellErrorType.BAD_REFERENCE)

        _, m = wb.new_sheet()
        self.assertEqual(wb.get_cell_value(n, "A4"), True)
        
    def test_choose(self):
        wb = sheets.Workbook()
        _, name = wb.new_sheet()
        
        wb.set_cell_contents(name, "A1", "4")
        wb.set_cell_contents(name, "A2", "True")
        wb.set_cell_contents(name, "A3", "=10+2")
        wb.set_cell_contents(name, "A5", "2")
        
        wb.set_cell_contents(name, "A4", '=CHOOSE(A5, A1, A2, A3)')
        self.assertEqual(wb.get_cell_value(name, "A4"), True)
        
    def test_iserror(self):
        wb = sheets.Workbook()
        _, name = wb.new_sheet()
        _, name2 = wb.new_sheet()
        
        wb.set_cell_contents(name, "A1", "#REF!")
        wb.set_cell_contents(name, "A2", "=ISERROR(A1)")
        self.assertEqual(wb.get_cell_value(name, "A2"), True)
        
        wb.set_cell_contents(name, "A3", "5")
        wb.set_cell_contents(name, "A4", "=ISERROR(A3)")
        self.assertEqual(wb.get_cell_value(name, "A4"), False)
        
        wb.set_cell_contents(name2, "A1", "#CIRCREF!")
        wb.set_cell_contents(name, "A5", f'=ISERROR({name2}!A1)')
        self.assertEqual(wb.get_cell_value(name, "A5"), True)
        

if __name__ == "__main__":
        unittest.main()
