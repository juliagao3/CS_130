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

        def test_default_sheet_name(self):
                wb = sheets.Workbook("wb")
                sheet_num, sheet_name = wb.new_sheet(None)
                self.assertEqual(sheet_name, "Sheet1")            
                
        def test_remove_sheet(self):
                wb = sheets.Workbook()
                i, n = wb.new_sheet()
                j, m = wb.new_sheet()
                self.assertEqual(wb.list_sheets(), ["Sheet1", "Sheet2"])
                wb.del_sheet(n)
                self.assertEqual(wb.list_sheets(), ["Sheet2"])
                
        def test_remove_and_add_sheet(self):
                wb = sheets.Workbook()
                i, n = wb.new_sheet()
                j, m = wb.new_sheet()
                self.assertEqual(wb.list_sheets(), [n,m])
                wb.del_sheet(n)
                self.assertEqual(wb.list_sheets(), [m])
                i, n = wb.new_sheet()
                self.assertEqual(wb.list_sheets(), [m,n])
                k, o = i, "something"
                wb.rename_sheet(n,o)
                self.assertEqual(wb.list_sheets(), [m,o])
                i, n = wb.copy_sheet(o)
                self.assertEqual(wb.list_sheets(), [m,o,n])
                wb.move_sheet(n, 0)
                self.assertEqual(wb.list_sheets(), [n,m,o])
                wb.move_sheet(o, 1)
                self.assertEqual(wb.list_sheets(), [n,o,m])
                wb.del_sheet(n)
                self.assertEqual(wb.list_sheets(), [o,m])
                wb.del_sheet(m)
                self.assertEqual(wb.list_sheets(), [o])
                wb.del_sheet(o)
                self.assertEqual(wb.list_sheets(), [])

        def test_overlapping_sheets(self):
                wb = sheets.Workbook()
                i, n = wb.new_sheet("Sheet1")
                j, m = wb.new_sheet()

                self.assertEqual(wb.list_sheets(), ["Sheet1", "Sheet2"])

        def test_empty_sheet_name(self):
                wb = sheets.Workbook()
                with self.assertRaises(ValueError):
                        wb.new_sheet("\t")
                with self.assertRaises(ValueError):
                        wb.new_sheet("\n")
                with self.assertRaises(ValueError):
                        wb.new_sheet("")
                with self.assertRaises(ValueError):
                        wb.new_sheet("  ")
                with self.assertRaises(ValueError):
                        wb.new_sheet("~Sheet")

        def test_empty_sheet(self):
                wb = sheets.Workbook("wb")
                sheet_num, sheet_name = wb.new_sheet(None)
                self.assertEqual(wb.get_cell_value(sheet_name, "A1"), None)

        def test_quoted_contents(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet(None)
                wb.set_cell_contents(sheet_name, "A1", "'1.000")
                self.assertEqual(wb.get_cell_value(sheet_name, "A1"), "1.000")

                wb.set_cell_contents(sheet_name, "A1", "'   1.000")
                self.assertEqual(wb.get_cell_value(sheet_name, "A1"), "   1.000")

                wb.set_cell_contents(sheet_name, "A1", "   '   1.000")
                self.assertEqual(wb.get_cell_value(sheet_name, "A1"), "   1.000")

        def test_bad_location(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet(None)
                with self.assertRaises(ValueError):
                        wb.set_cell_contents(sheet_name, "A 1", "'1.000")

                with self.assertRaises(ValueError):
                        wb.set_cell_contents(sheet_name, "a2a2", "test")

                with self.assertRaises(ValueError):
                        wb.set_cell_contents(sheet_name, "$A$A1", "invalid")
                        
        def test_out_of_bounds(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet(None)
                with self.assertRaises(ValueError):
                        wb.set_cell_contents(sheet_name, "ZZZZ99999", "1")

        def test_empty_ref(self):
                wb = sheets.Workbook("wb")
                sheet_num, sheet_name = wb.new_sheet(None)

                wb.set_cell_contents(sheet_name, "A2", "=A1")
                self.assertEqual(wb.get_cell_value(sheet_name, "A2"), decimal.Decimal(0))

        def test_concat_empty(self):
                wb = sheets.Workbook("wb")
                sheet_num, sheet_name = wb.new_sheet(None)

                wb.set_cell_contents(sheet_name, "A2", "hello")
                wb.set_cell_contents(sheet_name, "A3", "=A1 & A2")
                self.assertEqual(wb.get_cell_value(sheet_name, "A3"), "hello")

        def test_leading_whitespace(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet()

                wb.set_cell_contents(sheet_name, "A1","'      grrr  grrr   ")
                a1 = wb.get_cell_value(sheet_name, "A1")

                self.assertEqual(a1, "      grrr  grrr")

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

        def test_basic_unary_ops(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet()

                wb.set_cell_contents(sheet_name, "A1", "=-20.000")
                a1 = wb.get_cell_value(sheet_name, "A1")
                self.assertEqual(a1, decimal.Decimal(-20))

                wb.set_cell_contents(sheet_name, "A2", "=-A1")
                a2 = wb.get_cell_value(sheet_name, "A2")
                self.assertEqual(a2, decimal.Decimal(20))
                
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

        def test_basic_division(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet()

                wb.set_cell_contents(sheet_name, "A1", "=  10.0   /2.0  ")
                a1 = wb.get_cell_value(sheet_name, "A1")

                self.assertEqual(a1, decimal.Decimal(5))
                
        def test_multiplication(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet()
                
                wb.set_cell_contents(sheet_name, "A1", "= 5")
                wb.set_cell_contents(sheet_name, "A2", " =A1 *    6")
                
                self.assertEqual(wb.get_cell_value(sheet_name, "A2"), decimal.Decimal(30))
         
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

        @unittest.mock.patch('sys.stdout', new_callable=io.StringIO)
        def assert_stdout_for_object(self, obj, expected_print, mock_print):
                print(obj)
                self.assertEqual(mock_print.getvalue(), expected_print)

        #def test_print_cell(self):
        #        wb = sheets.Workbook()
        #        sheet_num, sheet_name = wb.new_sheet()

        #        wb.set_cell_contents(sheet_name, "A1", "testing print")
        #        a1 = wb.get_cell(sheet_name, "A1")
        #        self.assert_stdout_for_object(a1, "testing print\n") 

        #        a2 = wb.get_cell(sheet_name, "A2")
        #        self.assert_stdout_for_object(a2, "None\n")

        def test_non_finite_value(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet()

                wb.set_cell_contents(sheet_name, "A1", "-Infinity")
                a1 = wb.get_cell_value(sheet_name, "A1")
                self.assertEqual(a1, "-Infinity")

                wb.set_cell_contents(sheet_name, "A2", "NaN")
                a2 = wb.get_cell_value(sheet_name, "A2")
                self.assertEqual(a2, "NaN")

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
                
                wb.set_cell_contents(sheet_name, "A1", "=10+  ")
                a1 = wb.get_cell_value(sheet_name, "A1")

                self.assertIsInstance(a1, sheets.CellError)
                self.assertEqual(a1.get_type(), sheets.CellErrorType.PARSE_ERROR)

                # TODO: ??? "The formulas that are tested here are bad bc
                # they are missing one or more quotes for strings"
                wb.set_cell_contents(sheet_name, "A2", "=\"string     ")
                a2 = wb.get_cell_value(sheet_name, "A2")

                self.assertIsInstance(a2, sheets.CellError)
                self.assertEqual(a2.get_type(), sheets.CellErrorType.PARSE_ERROR)

                wb.set_cell_contents(sheet_name, "A3", "=  A1 +  5   ")
                a3 = wb.get_cell_value(sheet_name, "A3")

                self.assertIsInstance(a3, sheets.CellError)
                self.assertEqual(a3.get_type(), sheets.CellErrorType.PARSE_ERROR)

                wb.set_cell_contents(sheet_name, "A4", "=-A3")
                a4 = wb.get_cell_value(sheet_name, "A4")

                self.assertIsInstance(a4, sheets.CellError)
                self.assertEqual(a4.get_type(), sheets.CellErrorType.PARSE_ERROR)

                wb.set_cell_contents(sheet_name, "A4", "=-")
                a4 = wb.get_cell_value(sheet_name, "A4")

                self.assertIsInstance(a4, sheets.CellError)
                self.assertEqual(a4.get_type(), sheets.CellErrorType.PARSE_ERROR)
                
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
                
                def test_error_type(s: str, e: sheets.CellErrorType):
                        nonlocal wb
                        nonlocal sheet_name
                        wb.set_cell_contents(sheet_name, "A1", "=-" + s + "")
                        a1 = wb.get_cell_value(sheet_name, "A1")
                        self.assertIsInstance(a1, sheets.CellError)
                        self.assertEqual(a1.get_type(), e)
                        
                test_error_type("#ERROR!", sheets.CellErrorType.PARSE_ERROR)
                test_error_type("#REF!", sheets.CellErrorType.BAD_REFERENCE)
                test_error_type("#CIRCREF!", sheets.CellErrorType.CIRCULAR_REFERENCE)
                test_error_type("#NAME?", sheets.CellErrorType.BAD_NAME)
                test_error_type("#VALUE!", sheets.CellErrorType.TYPE_ERROR)
                test_error_type("#DIV/0!", sheets.CellErrorType.DIVIDE_BY_ZERO)

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

        def test_get_cell_contents(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet()

                a1 = wb.get_cell_contents(sheet_name, "A1")
                self.assertEqual(a1, None)

                wb.set_cell_contents(sheet_name, "A2", "contents")
                a2 = wb.get_cell_contents(sheet_name, "A2")
                self.assertEqual(a2, "contents")

        def test_num_sheets_and_list_sheets(self):
                wb = sheets.Workbook()
                sheet_num1, sheet_name1 = wb.new_sheet()
                sheet_num2, sheet_name2 = wb.new_sheet()
                sheet_num3, sheet_name3 = wb.new_sheet()
                
                self.assertEqual(wb.num_sheets(), 3)

                self.assertEqual(wb.list_sheets(), [sheet_name1, sheet_name2, sheet_name3])

                wb.del_sheet(sheet_name3)
                self.assertEqual(wb.num_sheets(), 2)
                self.assertEqual(wb.list_sheets(), [sheet_name1, sheet_name2])

        def test_double_cycle(self):
                wb = sheets.Workbook()
                (index, name) = wb.new_sheet()

                wb.set_cell_contents(name, "A1", "=A2")
                wb.set_cell_contents(name, "A3", "=A2")
                wb.set_cell_contents(name, "A2", "=A1+A3")

                value = wb.get_cell_value(name, "A1")
                self.assertEqual(type(value), sheets.CellError)
                self.assertEqual(value.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

                value = wb.get_cell_value(name, "A2")
                self.assertEqual(type(value), sheets.CellError)
                self.assertEqual(value.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

                value = wb.get_cell_value(name, "A3")
                self.assertEqual(type(value), sheets.CellError)
                self.assertEqual(value.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

                wb.set_cell_contents(name, "A1", "1")

                self.assertEqual(wb.get_cell_value(name, "A1"), decimal.Decimal(1.0))

                value = wb.get_cell_value(name, "A2")
                self.assertEqual(type(value), sheets.CellError)
                self.assertEqual(value.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

                value = wb.get_cell_value(name, "A3")
                self.assertEqual(type(value), sheets.CellError)
                self.assertEqual(value.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

                wb.set_cell_contents(name, "A3", "2")

                self.assertEqual(wb.get_cell_value(name, "A1"), decimal.Decimal(1.0))
                self.assertEqual(wb.get_cell_value(name, "A2"), decimal.Decimal(3.0))
                self.assertEqual(wb.get_cell_value(name, "A3"), decimal.Decimal(2.0))

        def test_spec(self):
                self.assertEqual(sheets.version, "1.3")
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
                with self.assertRaises(ValueError):
                     wb.new_sheet("july totals")
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
                wb.set_cell_contents(name, 'd1', None)
                wb.set_cell_contents(name, 'd2', '=D1')
                self.assertEqual(wb.get_cell_value(name, 'd2'), decimal.Decimal(0))

        def test_bad_ref_missing_sheet(self):
                wb = sheets.Workbook()
                sheet_num1, sheet_name1 = wb.new_sheet()

                wb.set_cell_contents(sheet_name1, "A1", "='New Sheet'!A1 + 10 ")
                wb.set_cell_contents(sheet_name1, "A2", "='New Sheet'!A2")

                sheet1_a1 = wb.get_cell_value(sheet_name1, "A1")
                self.assertIsInstance(sheet1_a1, sheets.CellError)
                self.assertEqual(sheet1_a1.get_type(), sheets.CellErrorType.BAD_REFERENCE)

                sheet1_a2 = wb.get_cell_value(sheet_name1, "A2")
                self.assertIsInstance(sheet1_a2, sheets.CellError)
                self.assertEqual(sheet1_a2.get_type(), sheets.CellErrorType.BAD_REFERENCE)

                sheet_num2, sheet_name2 = wb.new_sheet("New Sheet")

                wb.set_cell_contents(sheet_name2, "A1", "15")
                sheet1_a1 = wb.get_cell_value(sheet_name1, "A1")
                self.assertEqual(sheet1_a1, decimal.Decimal(25))

                sheet1_a2 = wb.get_cell_value(sheet_name1, "A2")
                self.assertEqual(sheet1_a2, decimal.Decimal(0))

                wb.del_sheet(sheet_name2)
                sheet1_a1 = wb.get_cell_value(sheet_name1, "A1")
                self.assertIsInstance(sheet1_a1, sheets.CellError)
                self.assertEqual(sheet1_a1.get_type(), sheets.CellErrorType.BAD_REFERENCE)
                
                sheet1_a2 = wb.get_cell_value(sheet_name1, "A2")
                self.assertIsInstance(sheet1_a2, sheets.CellError)
                self.assertEqual(sheet1_a2.get_type(), sheets.CellErrorType.BAD_REFERENCE)

        def test_add_cell_into_cycle(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet()
                
                wb.set_cell_contents(sheet_name, "A1", "=B1")
                wb.set_cell_contents(sheet_name, "C1", "=8/0 + #REF! + A1")

                value_2 = wb.get_cell_value(sheet_name, "C1")
                self.assertIsInstance(value_2, sheets.CellError)
                
                wb.set_cell_contents(sheet_name, "B1", "=C1")
                
                value_1 = wb.get_cell_value(sheet_name, "A1")
                self.assertIsInstance(value_1, sheets.CellError)
                self.assertEqual(value_1.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
                
                value_2 = wb.get_cell_value(sheet_name, "B1")
                self.assertIsInstance(value_2, sheets.CellError)
                self.assertEqual(value_2.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
                
                value_3 = wb.get_cell_value(sheet_name, "C1")
                self.assertIsInstance(value_3, sheets.CellError)
                self.assertEqual(value_3.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
                
                wb.set_cell_contents(sheet_name, "D1", "=A1")
                value_4 = wb.get_cell_value(sheet_name, "D1")
                self.assertIsInstance(value_4, sheets.CellError)
                self.assertEqual(value_4.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
                
        def test_removing_trailing_spaces_after_operation(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet()
                
                wb.set_cell_contents(sheet_name, "A1", "=1.0001")
                wb.set_cell_contents(sheet_name, "A3", "=0.0001")
                wb.set_cell_contents(sheet_name, "A4", "=A1 - A3")
                self.assertEqual(str(wb.get_cell_value(sheet_name, "A4")), '1')

                wb.set_cell_contents(sheet_name, "A2", "=A4 & \"word\"")
                self.assertEqual(wb.get_cell_value(sheet_name, "A2"), "1word")
                
        def test_error_setting(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet()
                
                wb.set_cell_contents(sheet_name, "A1", " #REF!")

                a1 = wb.get_cell_value(sheet_name, "A1")
                self.assertIsInstance(a1, sheets.CellError)
                self.assertEqual(a1.get_type(), sheets.CellErrorType.BAD_REFERENCE)
                
        def test_error_circular_reference(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet()
                
                wb.set_cell_contents(sheet_name, "A1", "=A1/0")
                a1 = wb.get_cell_value(sheet_name, "A1")
                self.assertIsInstance(a1, sheets.CellError)
                self.assertEqual(a1.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
                
                #fail this test, should be circular error instead of bad reference
                wb.set_cell_contents(sheet_name, "A2", "=A2 + Nonexistent!A2")
                a2 = wb.get_cell_value(sheet_name, "A2")
                self.assertIsInstance(a2, sheets.CellError)
                self.assertEqual(a2.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        def test_concat_trailing_zeros(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet()

                wb.set_cell_contents(sheet_name, "A1", '=0.000 & "hi"')
                self.assertEqual(wb.get_cell_value(sheet_name, 'A1'), '0hi')

                wb.set_cell_contents(sheet_name, "A2", '="hi" & 0.000')
                self.assertEqual(wb.get_cell_value(sheet_name, 'A2'), 'hi0')

                wb.set_cell_contents(sheet_name, "A3", "0.000")
                wb.set_cell_contents(sheet_name, "A4", '=A3 & "0.020"')
                wb.set_cell_contents(sheet_name, "A5", '=A4 + 0.0')
                self.assertEqual(wb.get_cell_value(sheet_name, 'A4'), '00.020')
                self.assertEqual(wb.get_cell_value(sheet_name, 'A5'), decimal.Decimal("0.02"))

                wb.set_cell_contents(sheet_name, "A5", "0.000")
                wb.set_cell_contents(sheet_name, "A6", '="hi" & A3')
                self.assertEqual(wb.get_cell_value(sheet_name, 'A6'), 'hi0')

                wb.set_cell_contents(sheet_name, "A6", '="hi" & (1.0 + 2.0)')
                self.assertEqual(wb.get_cell_value(sheet_name, 'A6'), 'hi3')

                wb.set_cell_contents(sheet_name, "A7", '="0.0"')
                self.assertEqual(wb.get_cell_value(sheet_name, 'A7'), '0.0')
                
                wb.set_cell_contents(sheet_name, "A8", '="hi" & (1.0 / 2.0)')
                self.assertEqual(wb.get_cell_value(sheet_name, 'A8'), 'hi0.5')
                
if __name__ == "__main__":
        unittest.main()
