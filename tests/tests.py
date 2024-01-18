#! /usr/bin/env python3
import unittest

import sys
import sheets
import decimal
import traceback
import re

class TestClass(unittest.TestCase):

        def test_default_sheet_name(self):
                wb = sheets.Workbook("wb")
                sheet_num, sheet_name = wb.new_sheet(None)
                self.assertEqual(sheet_name, "Sheet1")            
                
        #def test_special_name(self):
        #        wb = sheets.Workbook("wb")
        #        with self.assertRaises(ValueError):
        #                wb.new_sheet("~Sheet")

        def test_empty_sheet(self):
                wb = sheets.Workbook("wb")
                sheet_num, sheet_name = wb.new_sheet(None)
                self.assertEqual(wb.get_cell_value(sheet_name, "A1"), None)

        def test_quoted_contents(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet(None)
                wb.set_cell_contents(sheet_name, "A1", "'1.000")
                self.assertEqual(wb.get_cell_value(sheet_name, "A1"), "1.000")

        def test_bad_location(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet(None)
                with self.assertRaises(KeyError):
                        wb.set_cell_contents(sheet_name, "A 1", "'1.000")
                        
        def test_out_of_bounds(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet(None)
                with self.assertRaises(ValueError):
                        wb.set_cell_contents(sheet_name, "ZZZZ99999", "1")

        def test_one_plus_one(self):
                wb = sheets.Workbook("wb")
                sheet_num, sheet_name = wb.new_sheet(None)

                wb.set_cell_contents(sheet_name, "A1", "=1+1")
                self.assertEqual(wb.get_cell_value(sheet_name, "A1"), decimal.Decimal(2))
                
        def test_one_plus_string(self):
              wb = sheets.Workbook("wb")
              sheet_num, sheet_name = wb.new_sheet(None)

              wb.set_cell_contents(sheet_name, "A1", '=1+ "hello"')
              a1 = wb.get_cell_value(sheet_name, "A1")
              self.assertIsInstance(a1, sheets.CellError)
              self.assertEqual(a1.get_type(), sheets.CellErrorType.TYPE_ERROR)
                
        def test_one_minus_unary_string(self):
                wb = sheets.Workbook("wb")
                sheet_num, sheet_name = wb.new_sheet(None)

                wb.set_cell_contents(sheet_name, "A1", '=1 - -"hello"')
                a1 = wb.get_cell_value(sheet_name, "A1")
                self.assertIsInstance(a1, sheets.CellError)
                self.assertEqual(a1.get_type(), sheets.CellErrorType.TYPE_ERROR)

        def test_one_plus_empty(self):
                wb = sheets.Workbook("wb")
                sheet_num, sheet_name = wb.new_sheet(None)

                wb.set_cell_contents(sheet_name, "A1", "1")
                wb.set_cell_contents(sheet_name, "A3", "=A1+A2")

                self.assertEqual(wb.get_cell_value(sheet_name, "A1"), decimal.Decimal(1))
                self.assertEqual(wb.get_cell_value(sheet_name, "A3"), decimal.Decimal(1))

        def test_one_plus_one_cells(self):
                wb = sheets.Workbook("wb")
                sheet_num, sheet_name = wb.new_sheet(None)

                wb.set_cell_contents(sheet_name, "A1", "1")
                wb.set_cell_contents(sheet_name, "A2", "=1")
                wb.set_cell_contents(sheet_name, "A3", "=A1+A2")

                self.assertEqual(wb.get_cell_value(sheet_name, "A1"), decimal.Decimal(1))
                self.assertEqual(wb.get_cell_value(sheet_name, "A2"), decimal.Decimal(1))
                self.assertEqual(wb.get_cell_value(sheet_name, "A3"), decimal.Decimal(2))

        def test_string_arithmetic_cells(self):
                wb = sheets.Workbook("wb")
                sheet_num, sheet_name = wb.new_sheet(None)

                wb.set_cell_contents(sheet_name, "A1", "1")
                wb.set_cell_contents(sheet_name, "A2", "hello")
                wb.set_cell_contents(sheet_name, "A3", "=A1+A2")

                self.assertEqual(wb.get_cell_value(sheet_name, "A1"), decimal.Decimal(1))

                a3 = wb.get_cell_value(sheet_name, "A3")

                self.assertIsInstance(a3, sheets.CellError)
                self.assertEqual(a3.get_type(), sheets.CellErrorType.TYPE_ERROR)

        def test_cell_update(self):
                wb = sheets.Workbook("wb")
                sheet_num, sheet_name = wb.new_sheet(None)

                wb.set_cell_contents(sheet_name, "A1", "1")
                wb.set_cell_contents(sheet_name, "A2", "1")
                wb.set_cell_contents(sheet_name, "A3", "=A1+A2")

                self.assertEqual(wb.get_cell_value(sheet_name, "A3"), decimal.Decimal(2))

                wb.set_cell_contents(sheet_name, "A2", "2")

                self.assertEqual(wb.get_cell_value(sheet_name, "A3"), decimal.Decimal(3))
                
        def test_cell_update_multiple(self):
                wb = sheets.Workbook("wb")
                sheet_num, sheet_name = wb.new_sheet(None)

                wb.set_cell_contents(sheet_name, "A1", "1")
                wb.set_cell_contents(sheet_name, "A2", "=A1")
                wb.set_cell_contents(sheet_name, "A3", "=A1+A2")

                self.assertEqual(wb.get_cell_value(sheet_name, "A3"), decimal.Decimal(2))

                wb.set_cell_contents(sheet_name, "A1", "2")

                self.assertEqual(wb.get_cell_value(sheet_name, "A3"), decimal.Decimal(4))

        def test_circular_refs(self):
                wb = sheets.Workbook("wb")
                sheet_num, sheet_name = wb.new_sheet(None)

                wb.set_cell_contents(sheet_name, "A1", "=A2")
                wb.set_cell_contents(sheet_name, "A2", "=A1")

                a1 = wb.get_cell_value(sheet_name, "A1")
                a2 = wb.get_cell_value(sheet_name, "A1")

                self.assertIsInstance(a1, sheets.CellError)
                self.assertIsInstance(a2, sheets.CellError)

                self.assertEqual(a1.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
                self.assertEqual(a2.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        def test_circular_refs_with_tail(self):
                wb = sheets.Workbook("wb")
                sheet_num, sheet_name = wb.new_sheet(None)

                wb.set_cell_contents(sheet_name, "A1", "=A2")
                wb.set_cell_contents(sheet_name, "A2", "=A1+A4")
                wb.set_cell_contents(sheet_name, "A3", "=\"Hello \" & A1 & \"!\"")

                a1 = wb.get_cell_value(sheet_name, "A1")
                a2 = wb.get_cell_value(sheet_name, "A2")
                a3 = wb.get_cell_value(sheet_name, "A3")

                self.assertIsInstance(a1, sheets.CellError)
                self.assertIsInstance(a2, sheets.CellError)
                self.assertIsInstance(a3, str)

                self.assertEqual(a1.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
                self.assertEqual(a2.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        def test_grow_and_shrink_extent(self):
                wb = sheets.Workbook("wb")
                sheet_num, sheet_name = wb.new_sheet(None)

                wb.set_cell_contents(sheet_name, "A1", "first")
                wb.set_cell_contents(sheet_name, "B2", "second")

                init_extent = wb.get_sheet_extent(sheet_name)
                self.assertEqual(init_extent, (2, 2))

                wb.set_cell_contents(sheet_name, "C4", "third")
                new_extent = wb.get_sheet_extent(sheet_name)
                self.assertEqual(new_extent, (3, 4))

                wb.set_cell_contents(sheet_name, "C4", "")
                self.assertEqual(wb.get_sheet_extent(sheet_name), init_extent)
                
        def test_multiple_sheets(self):
                wb = sheets.Workbook("wb")
                sheet_num1, sheet_name1 = wb.new_sheet("Sheet1")
                sheet_num2, sheet_name2 = wb.new_sheet("Sheet2")

                wb.set_cell_contents(sheet_name1, "A1", "10")
                wb.set_cell_contents(sheet_name2, "A1", "=Sheet1!A1+5")

                self.assertEqual(wb.get_cell_value(sheet_name2, "A1"), decimal.Decimal(15))

                wb.set_cell_contents(sheet_name1, "A1", "20")

                self.assertEqual(wb.get_cell_value(sheet_name2, "A1"), decimal.Decimal(25))

                wb.set_cell_contents(sheet_name1, "A1", "=Sheet2!A1")

                sheet1_a1 = wb.get_cell_value(sheet_name1, "A1")
                sheet2_a1 = wb.get_cell_value(sheet_name2, "A1")

                self.assertIsInstance(sheet1_a1, sheets.CellError)
                self.assertIsInstance(sheet2_a1, sheets.CellError)

                self.assertEqual(sheet1_a1.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
                self.assertEqual(sheet2_a1.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        def test_delete_add(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name1 = wb.new_sheet("Sheet1")
                sheet_num, sheet_name2 = wb.new_sheet("Sheet2")
                
                wb.set_cell_contents(sheet_name1, "A1", "1.0")
                wb.set_cell_contents(sheet_name2, "A1", "=Sheet1!A1 + 2.0")
                
                self.assertEqual(wb.get_cell_value(sheet_name2, "A1"), decimal.Decimal(3))
                
                wb.del_sheet(sheet_name1)
                
                a1 = wb.get_cell_value(sheet_name2, "A1")
                self.assertIsInstance(a1, sheets.CellError)
                self.assertEqual(a1.get_type(), sheets.CellErrorType.BAD_REFERENCE)

                wb.new_sheet(sheet_name1)
                
                self.assertEqual(wb.get_cell_value(sheet_name2, "A1"), decimal.Decimal(2))

        def test_div_by_zero(self):
                wb = sheets.Workbook()
                sheet_num1, sheet_name1 = wb.new_sheet()
                sheet_num2, sheet_name2 = wb.new_sheet()

                wb.set_cell_contents(sheet_name1, "A1", "100")
                wb.set_cell_contents(sheet_name2, "B2", "=Sheet1!A1/0")
                
                b2 = wb.get_cell_value(sheet_name2, "B2")
                
                self.assertIsInstance(b2, sheets.CellError)
                self.assertEqual(b2.get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)
        
        def test_parse_error(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet()
                
                wb.set_cell_contents(sheet_name, "A1", "=10+")
                a1 = wb.get_cell_value(sheet_name, "A1")

                self.assertIsInstance(a1, sheets.CellError)
                self.assertEqual(a1.get_type(), sheets.CellErrorType.PARSE_ERROR)
                
        def test_bad_reference(self):
                wb = sheets.Workbook()
                sheet_num1, sheet_name1 = wb.new_sheet()
                sheet_num2, sheet_name2 = wb.new_sheet()

                wb.set_cell_contents(sheet_name1, "A1", "1000.000")
                
                wb.set_cell_contents(sheet_name2, "A1", "=Sheet3!A1")
                value_1 = wb.get_cell_value(sheet_name2, "A1")
                self.assertIsInstance(value_1, sheets.CellError)
                self.assertEqual(value_1.get_type(), sheets.CellErrorType.BAD_REFERENCE)

                wb.set_cell_contents(sheet_name2, "A2", "=Sheet1!AAAAAAA1")
                value_2 = wb.get_cell_value(sheet_name2, "A2")
                self.assertIsInstance(value_2, sheets.CellError)
                self.assertEqual(value_2.get_type(), sheets.CellErrorType.BAD_REFERENCE)

                wb.set_cell_contents(sheet_name2, "A3", "=Sheet3!AAAAAAA1")
                value_2 = wb.get_cell_value(sheet_name2, "A3")
                self.assertIsInstance(value_2, sheets.CellError)
                self.assertEqual(value_2.get_type(), sheets.CellErrorType.BAD_REFERENCE)                

        # def test_bad_name():
        #         pass

        def test_error_string_reps(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet()
                
                wb.set_cell_contents(sheet_name, "A1", "#DIV/0!")
                wb.set_cell_contents(sheet_name, "A2", "'#DIV/0!")

                a1 = wb.get_cell_value(sheet_name, "A1")
                a2 = wb.get_cell_value(sheet_name, "A2")

                self.assertEqual(type(a1), sheets.CellError)
                self.assertEqual(a1.get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)
                self.assertEqual(a2, "#DIV/0!")

                wb.set_cell_contents(sheet_name, "A1", "#ERROR!")
                a1 = wb.get_cell_value(sheet_name, "A1")
                self.assertEqual(type(a1), sheets.CellError)
                self.assertEqual(a1.get_type(), sheets.CellErrorType.PARSE_ERROR)

                wb.set_cell_contents(sheet_name, "A1", "#Circref!")
                a1 = wb.get_cell_value(sheet_name, "A1")
                self.assertEqual(type(a1), sheets.CellError)
                self.assertEqual(a1.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

                wb.set_cell_contents(sheet_name, "A1", "#ref!")
                a1 = wb.get_cell_value(sheet_name, "A1")
                self.assertEqual(type(a1), sheets.CellError)
                self.assertEqual(a1.get_type(), sheets.CellErrorType.BAD_REFERENCE)

        def test_error_literals_in_formulas(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet()
                
                wb.set_cell_contents(sheet_name, "A1", "=#REF!+5")

                a1 = wb.get_cell_value(sheet_name, "A1")

                self.assertIsInstance(a1, sheets.CellError)
                self.assertEqual(a1.get_type(), sheets.CellErrorType.BAD_REFERENCE)


        def test_error_priorities(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet()
                
                wb.set_cell_contents(sheet_name, "A1", "=B1+5")
                wb.set_cell_contents(sheet_name, "B1", "=10.0/0")

                a1 = wb.get_cell_value(sheet_name, "A1")
                b1 = wb.get_cell_value(sheet_name, "B1")

                self.assertIsInstance(a1, sheets.CellError)
                self.assertEqual(a1.get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)
                self.assertIsInstance(b1, sheets.CellError)
                self.assertEqual(b1.get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

                wb.set_cell_contents(sheet_name, "A2", "=B2")
                wb.set_cell_contents(sheet_name, "B2", "=C2")
                wb.set_cell_contents(sheet_name, "C2", "=B2/0")

                # should prioritize circref error over div/0 error
                a2 = wb.get_cell_value(sheet_name, "A2")
                b2 = wb.get_cell_value(sheet_name, "B2")
                c2 = wb.get_cell_value(sheet_name, "C2")

                self.assertIsInstance(a2, sheets.CellError)
                self.assertEqual(a2.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
                self.assertIsInstance(b2, sheets.CellError)
                self.assertEqual(b2.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
                self.assertIsInstance(c2, sheets.CellError)
                self.assertEqual(c2.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

                # parse error should have the highest priority (?)
                wb.set_cell_contents(sheet_name, "A3", "=#REF!+")
                wb.set_cell_contents(sheet_name, "B3", "=A3+1")

                a3 = wb.get_cell_value(sheet_name, "A3")
                b3 = wb.get_cell_value(sheet_name, "B3")
                
                self.assertIsInstance(a3, sheets.CellError)
                self.assertEqual(a3.get_type(), sheets.CellErrorType.PARSE_ERROR)
                self.assertIsInstance(b3, sheets.CellError)
                self.assertEqual(b3.get_type(), sheets.CellErrorType.PARSE_ERROR)
                
                wb.set_cell_contents(sheet_name, "A5", "=#error!")
                wb.set_cell_contents(sheet_name, "B5", "=A5+#Ref!")
                b5 = wb.get_cell_value(sheet_name, "B5")
                self.assertIsInstance(b5, sheets.CellError)

        def test_parentheses_in_formulas(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet("sheet1")
                
                wb.set_cell_contents(sheet_name, "A1", "=5*(1+1+(2*4))")
                self.assertEqual(wb.get_cell_value(sheet_name, "A1"), decimal.Decimal(50))
                
        def test_formula_in_parenthesis(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name1 = wb.new_sheet("sheet1")
                sheet_num, sheet_name2 = wb.new_sheet("sheet2")
                
                wb.set_cell_contents(sheet_name1, "A1", "=6")
                wb.set_cell_contents(sheet_name2, "A1", "= 5 * (sheet1!A1 + 1)")
                
                self.assertEqual(wb.get_cell_value(sheet_name2, "A1"), decimal.Decimal(35))

        def test_trailing_zeros_formula(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name1 = wb.new_sheet("sheet1")
                
                wb.set_cell_contents(sheet_name1, "A1", "=1.00001")
                wb.set_cell_contents(sheet_name1, "A2", "=0.00001")
                wb.set_cell_contents(sheet_name1, "A3", "=A1-A2")
                
                self.assertEqual(wb.get_cell_value(sheet_name1, "A3"), decimal.Decimal(1))
                
        def concat_with_error(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet("sheet1")
                
                wb.set_cell_contents(sheet_name, "A1", "=1")
                wb.set_cell_contents(sheet_name, "A2", "=#Ref!")
                wb.set_cell_contents(sheet_name, "A3", "=A1&A2")
                
                self.assertIsInstance(wb.get_cell_value(sheet_name, "A3"), sheets.CellError)
                
        def get_empty_cell(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet()
                
                wb.set_cell_contents(sheet_name, "A1", "2")
                with self.assertRaises(KeyError):
                        wb.get_cell_value(sheet_name, "A2")

        def test_spec(self):
                self.assertEqual(sheets.version, "1.0")
                wb = sheets.Workbook()
                (index, name) = wb.new_sheet()
                self.assertEqual(index, 0)
                self.assertEqual(name, "Sheet1")
                wb.set_cell_contents(name, 'a1', '12')
                wb.set_cell_contents(name, 'b1', '34')
                wb.set_cell_contents(name, 'c1', '=a1+b1')
                value = wb.get_cell_value(name, 'c1')
                self.assertEqual(value, decimal.Decimal('46'))
                wb.set_cell_contents(name, 'd3', '=nonexistent!b4')
                value = wb.get_cell_value(name, 'd3')
                self.assertEqual(value.get_type(), sheets.CellErrorType.BAD_REFERENCE)
                self.assertIsInstance(value, sheets.CellError)
                wb.set_cell_contents(name, 'e1', '#div/0!')
                wb.set_cell_contents(name, 'e2', '=e1+5')
                value = wb.get_cell_value(name, 'e2')
                self.assertIsInstance(value, sheets.CellError)
                self.assertEqual(value.get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        def test_spec_2(self):
                wb = sheets.Workbook()
                (index, name) = wb.new_sheet("July Totals")
                # with self.assertRaises(ValueError):
                #      wb.new_sheet("july totals")
                wb.set_cell_contents(name, 'a1', "'=")
                self.assertEqual(wb.get_cell_value(name, 'a1'), "=")
                wb.set_cell_contents(name, 'a2', "''")
                self.assertEqual(wb.get_cell_value(name, 'a2'), "'")
                wb.set_cell_contents(name, 'a3', "'    hello")
                self.assertEqual(wb.get_cell_value(name, 'a3'), "    hello")
                wb.set_cell_contents(name, 'a4', "=3*'abc'")
                value = wb.get_cell_value(name, 'a4')
                self.assertIsInstance(value, sheets.CellError)
                self.assertEqual(value.get_type(), sheets.CellErrorType.PARSE_ERROR)
                wb.set_cell_contents(name, 'b1', "'123")
                wb.set_cell_contents(name, 'b2', "5.3")
                wb.set_cell_contents(name, 'b3', "=b1*B2")
                value = wb.get_cell_value(name, 'b3')
                self.assertEqual(value, decimal.Decimal('651.9'))
                wb.set_cell_contents(name, 'c1', "'    123")
                wb.set_cell_contents(name, 'c2', "5.3")
                wb.set_cell_contents(name, 'c3', "=C1*c2")
                value = wb.get_cell_value(name, 'c3')
                self.assertEqual(value, decimal.Decimal('651.9')) 
                # wb.set_cell_contents(name, 'd1', None)
                # wb.set_cell_contents(name, 'd2', '=D1')
                # self.assertEqual(wb.get_cell_value(name, 'd2'), decimal.Decimal(0))

if __name__ == "__main__":
        unittest.main(verbosity=2)
