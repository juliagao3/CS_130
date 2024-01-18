#! /usr/bin/env python3

import sys
import sheets
import decimal
import traceback

def fail(message):
        assert False, (message, traceback.extract_stack(limit=2)[-2])

def assert_eq(a, b):
        assert a == b, (f'{str(a)} != {str(b)}', traceback.extract_stack(limit=2)[-2])

def test_default_sheet_name():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)
        assert_eq(sheet_name, "Sheet1")

def test_empty_sheet():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)
        assert_eq(wb.get_cell_value(sheet_name, "A1"), None)

def test_quoted_contents():
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet(None)
        wb.set_cell_contents(sheet_name, "A1", "'1.000")
        assert_eq(wb.get_cell_value(sheet_name, "A1"), "1.000")

def test_bad_location():
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet(None)
        try:
            wb.set_cell_contents(sheet_name, "A 1", "'1.000")
            fail("Expected ValueError")
        except ValueError:
            pass

def test_one_plus_one():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", "=1+1")
        assert_eq(wb.get_cell_value(sheet_name, "A1"), decimal.Decimal(2))
        
def test_one_plus_string():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", '=1+ "hello"')
        a1 = wb.get_cell_value(sheet_name, "A1")
        assert_eq(type(a1), sheets.CellError)
        assert_eq(a1.get_type(), sheets.CellErrorType.TYPE_ERROR)
        
def test_one_minus_unary_string():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", '=1 - -"hello"')
        a1 = wb.get_cell_value(sheet_name, "A1")
        assert_eq(type(a1), sheets.CellError)
        assert_eq(a1.get_type(), sheets.CellErrorType.TYPE_ERROR)

def test_one_plus_empty():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", "1")
        wb.set_cell_contents(sheet_name, "A3", "=A1+A2")

        assert_eq(wb.get_cell_value(sheet_name, "A1"), decimal.Decimal(1))
        assert_eq(wb.get_cell_value(sheet_name, "A3"), decimal.Decimal(1))

def test_one_plus_one_cells():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", "1")
        wb.set_cell_contents(sheet_name, "A2", "=1")
        wb.set_cell_contents(sheet_name, "A3", "=A1+A2")

        assert_eq(wb.get_cell_value(sheet_name, "A1"), decimal.Decimal(1))
        assert_eq(wb.get_cell_value(sheet_name, "A2"), decimal.Decimal(1))
        assert_eq(wb.get_cell_value(sheet_name, "A3"), decimal.Decimal(2))

def test_string_arithmetic_cells():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", "1")
        wb.set_cell_contents(sheet_name, "A2", "hello")
        wb.set_cell_contents(sheet_name, "A3", "=A1+A2")

        assert_eq(wb.get_cell_value(sheet_name, "A1"), decimal.Decimal(1))

        a3 = wb.get_cell_value(sheet_name, "A3")

        assert_eq(type(a3), sheets.CellError)
        assert_eq(a3.get_type(), sheets.CellErrorType.TYPE_ERROR)

def test_cell_update():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", "1")
        wb.set_cell_contents(sheet_name, "A2", "1")
        wb.set_cell_contents(sheet_name, "A3", "=A1+A2")

        assert_eq(wb.get_cell_value(sheet_name, "A3"), decimal.Decimal(2))

        wb.set_cell_contents(sheet_name, "A2", "2")

        assert_eq(wb.get_cell_value(sheet_name, "A3"), decimal.Decimal(3))
        
def test_cell_update_multiple():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", "1")
        wb.set_cell_contents(sheet_name, "A2", "=A1")
        wb.set_cell_contents(sheet_name, "A3", "=A1+A2")

        assert_eq(wb.get_cell_value(sheet_name, "A3"), decimal.Decimal(2))

        wb.set_cell_contents(sheet_name, "A1", "2")

        assert_eq(wb.get_cell_value(sheet_name, "A3"), decimal.Decimal(4))

def test_circular_refs():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", "=A2")
        wb.set_cell_contents(sheet_name, "A2", "=A1")

        a1 = wb.get_cell_value(sheet_name, "A1")
        a2 = wb.get_cell_value(sheet_name, "A1")

        assert_eq(type(a1), sheets.CellError)
        assert_eq(type(a2), sheets.CellError)

        assert_eq(a1.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        assert_eq(a2.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

def test_circular_refs_with_tail():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", "=A2")
        wb.set_cell_contents(sheet_name, "A2", "=A1+A4")
        wb.set_cell_contents(sheet_name, "A3", "=\"Hello \" & A1 & \"!\"")

        a1 = wb.get_cell_value(sheet_name, "A1")
        a2 = wb.get_cell_value(sheet_name, "A2")
        a3 = wb.get_cell_value(sheet_name, "A3")

        assert_eq(type(a1), sheets.CellError)
        assert_eq(type(a2), sheets.CellError)
        assert_eq(type(a3), str)

        assert_eq(a1.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        assert_eq(a2.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

def test_grow_and_shrink_extent():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", "first")
        wb.set_cell_contents(sheet_name, "B2", "second")

        init_extent = wb.get_sheet_extent(sheet_name)
        assert_eq(init_extent, (2, 2))

        wb.set_cell_contents(sheet_name, "C4", "third")
        new_extent = wb.get_sheet_extent(sheet_name)
        assert_eq(new_extent, (3, 4))

        wb.set_cell_contents(sheet_name, "C4", "")
        assert_eq(wb.get_sheet_extent(sheet_name), init_extent)
        
def test_multiple_sheets():
        wb = sheets.Workbook("wb")
        sheet_num1, sheet_name1 = wb.new_sheet("Sheet1")
        sheet_num2, sheet_name2 = wb.new_sheet("Sheet2")

        wb.set_cell_contents(sheet_name1, "A1", "10")
        wb.set_cell_contents(sheet_name2, "A1", "=Sheet1!A1+5")

        assert_eq(wb.get_cell_value(sheet_name2, "A1"), decimal.Decimal(15))

        wb.set_cell_contents(sheet_name1, "A1", "20")

        assert_eq(wb.get_cell_value(sheet_name2, "A1"), decimal.Decimal(25))

        wb.set_cell_contents(sheet_name1, "A1", "=Sheet2!A1")

        sheet1_a1 = wb.get_cell_value(sheet_name1, "A1")
        sheet2_a1 = wb.get_cell_value(sheet_name2, "A1")

        assert_eq(type(sheet1_a1), sheets.CellError)
        assert_eq(type(sheet2_a1), sheets.CellError)

        assert_eq(sheet1_a1.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        assert_eq(sheet2_a1.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

def test_delete_add():
        wb = sheets.Workbook()
        sheet_num, sheet_name1 = wb.new_sheet("Sheet1")
        sheet_num, sheet_name2 = wb.new_sheet("Sheet2")
        
        wb.set_cell_contents(sheet_name1, "A1", "1.0")
        wb.set_cell_contents(sheet_name2, "A1", "=Sheet1!A1 + 2.0")
        
        assert_eq(wb.get_cell_value(sheet_name2, "A1"), decimal.Decimal(3))
        
        wb.del_sheet(sheet_name1)
        
        a1 = wb.get_cell_value(sheet_name2, "A1")
        assert_eq(type(a1), sheets.CellError)
        assert_eq(a1.get_type(), sheets.CellErrorType.BAD_REFERENCE)

        wb.new_sheet(sheet_name1)
        
        assert_eq(wb.get_cell_value(sheet_name2, "A1"), decimal.Decimal(2))

def test_div_by_zero():
        wb = sheets.Workbook()
        sheet_num1, sheet_name1 = wb.new_sheet()
        sheet_num2, sheet_name2 = wb.new_sheet()

        wb.set_cell_contents(sheet_name1, "A1", "100")
        wb.set_cell_contents(sheet_name2, "B2", "=Sheet1!A1/0")
        
        b2 = wb.get_cell_value(sheet_name2, "B2")
        
        assert_eq(type(b2), sheets.CellError)
        assert_eq(b2.get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)
       
def test_parse_error():
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()
        
        wb.set_cell_contents(sheet_name, "A1", "=10+")
        a1 = wb.get_cell_value(sheet_name, "A1")

        assert_eq(type(a1), sheets.CellError)
        assert_eq(a1.get_type(), sheets.CellErrorType.PARSE_ERROR)
        
def test_bad_reference():
        wb = sheets.Workbook()
        sheet_num1, sheet_name1 = wb.new_sheet()
        sheet_num2, sheet_name2 = wb.new_sheet()

        wb.set_cell_contents(sheet_name1, "A1", "1000.000")
        
        wb.set_cell_contents(sheet_name2, "A1", "=Sheet3!A1")
        value_1 = wb.get_cell_value(sheet_name2, "A1")
        assert_eq(type(value_1), sheets.CellError)
        assert_eq(value_1.get_type(), sheets.CellErrorType.BAD_REFERENCE)

        wb.set_cell_contents(sheet_name2, "A2", "=Sheet1!AAAAAAA1")
        value_2 = wb.get_cell_value(sheet_name2, "A2")
        assert_eq(type(value_2), sheets.CellError)
        assert_eq(value_2.get_type(), sheets.CellErrorType.BAD_REFERENCE)

        wb.set_cell_contents(sheet_name2, "A3", "=Sheet3!AAAAAAA1")
        value_2 = wb.get_cell_value(sheet_name2, "A3")
        assert_eq(type(value_2), sheets.CellError)
        assert_eq(value_2.get_type(), sheets.CellErrorType.BAD_REFERENCE)                

def test_bad_name():
        pass

def test_error_string_reps():
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()
        
        wb.set_cell_contents(sheet_name, "A1", "#DIV/0!")
        wb.set_cell_contents(sheet_name, "A2", "'#DIV/0!")

        a1 = wb.get_cell_value(sheet_name, "A1")
        a2 = wb.get_cell_value(sheet_name, "A2")

        assert_eq(type(a1), sheets.CellError)
        assert_eq(a1.get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)
        assert_eq(a2, "#DIV/0!")

        wb.set_cell_contents(sheet_name, "A1", "#ERROR!")
        a1 = wb.get_cell_value(sheet_name, "A1")
        assert_eq(type(a1), sheets.CellError)
        assert_eq(a1.get_type(), sheets.CellErrorType.PARSE_ERROR)

        wb.set_cell_contents(sheet_name, "A1", "#Circref!")
        a1 = wb.get_cell_value(sheet_name, "A1")
        assert_eq(type(a1), sheets.CellError)
        assert_eq(a1.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        wb.set_cell_contents(sheet_name, "A1", "#ref!")
        a1 = wb.get_cell_value(sheet_name, "A1")
        assert_eq(type(a1), sheets.CellError)
        assert_eq(a1.get_type(), sheets.CellErrorType.BAD_REFERENCE)

def test_error_literals_in_formulas():
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()
        
        wb.set_cell_contents(sheet_name, "A1", "=#REF!+5")

        a1 = wb.get_cell_value(sheet_name, "A1")

        assert_eq(type(a1), sheets.CellError)
        assert_eq(a1.get_type(), sheets.CellErrorType.BAD_REFERENCE)


def test_error_priorities():
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()
        
        wb.set_cell_contents(sheet_name, "A1", "=B1+5")
        wb.set_cell_contents(sheet_name, "B1", "=10.0/0")

        a1 = wb.get_cell_value(sheet_name, "A1")
        b1 = wb.get_cell_value(sheet_name, "B1")

        assert_eq(type(a1), sheets.CellError)
        assert_eq(a1.get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)
        assert_eq(type(b1), sheets.CellError)
        assert_eq(b1.get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        wb.set_cell_contents(sheet_name, "A2", "=B2")
        wb.set_cell_contents(sheet_name, "B2", "=C2")
        wb.set_cell_contents(sheet_name, "C2", "=B2/0")

        # should prioritize circref error over div/0 error
        a2 = wb.get_cell_value(sheet_name, "A2")
        b2 = wb.get_cell_value(sheet_name, "B2")
        c2 = wb.get_cell_value(sheet_name, "C2")

        assert_eq(type(a2), sheets.CellError)
        assert_eq(a2.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        assert_eq(type(b2), sheets.CellError)
        assert_eq(b2.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        assert_eq(type(c2), sheets.CellError)
        assert_eq(c2.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        # parse error should have the highest priority (?)
        wb.set_cell_contents(sheet_name, "A3", "=#REF!+")
        wb.set_cell_contents(sheet_name, "B3", "=A3+1")

        a3 = wb.get_cell_value(sheet_name, "A3")
        b3 = wb.get_cell_value(sheet_name, "B3")
        
        assert_eq(type(a3), sheets.CellError)
        assert_eq(a3.get_type(), sheets.CellErrorType.PARSE_ERROR)
        assert_eq(type(b3), sheets.CellError)
        assert_eq(b3.get_type(), sheets.CellErrorType.PARSE_ERROR)
        
        wb.set_cell_contents(sheet_name, "A5", "=#error!")
        wb.set_cell_contents(sheet_name, "B5", "=A5+#Ref!")
        b5 = wb.get_cell_value(sheet_name, "B5")
        assert_eq(type(b5), sheets.CellError)
        #assert_eq(b5.get_type(), sheets.CellErrorType.PARSE_ERROR)

def test_parentheses_in_formulas():
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet("sheet1")
        
        wb.set_cell_contents(sheet_name, "A1", "=5*(1+1+(2*4))")
        assert_eq(wb.get_cell_value(sheet_name, "A1"), decimal.Decimal(50))
        
def test_formula_in_parenthesis():
        wb = sheets.Workbook()
        sheet_num, sheet_name1 = wb.new_sheet("sheet1")
        sheet_num, sheet_name2 = wb.new_sheet("sheet2")
        
        wb.set_cell_contents(sheet_name1, "A1", "=6")
        wb.set_cell_contents(sheet_name2, "A1", "= 5 * (sheet1!A1 + 1)")
        
        assert_eq(wb.get_cell_value(sheet_name2, "A1"), decimal.Decimal(35))

def test_all():
        tests = [
                test_default_sheet_name,
                test_quoted_contents,
                test_bad_location,
                test_one_plus_one,
                test_one_plus_string,
                test_one_minus_unary_string,
                test_one_plus_one_cells,
                test_one_plus_empty,
                test_string_arithmetic_cells,
                test_circular_refs,
                test_circular_refs_with_tail,
                test_grow_and_shrink_extent,
                test_cell_update,
                test_cell_update_multiple,
                test_multiple_sheets, 
                test_delete_add,
                test_div_by_zero,
                test_parse_error,
                test_bad_reference,
                test_error_string_reps,
                test_error_literals_in_formulas,
                test_error_priorities,
                test_parentheses_in_formulas,
                test_formula_in_parenthesis
        ]

        GREEN  = "\033[32m"
        RED    = "\033[31m"
        NORMAL = "\033[0m"

        for t in tests:
                try:
                        t()
                        print(GREEN, "PASSED ", NORMAL, t.__name__)
                except AssertionError as e:
                        print(RED, "FAILED ", NORMAL, t.__name__)
                        print("\t", "line", e.args[0][1][1])
                        print("\t", e.args[0][1][3])
                        print("\t", e.args[0][0])
                except RecursionError as e:
                        print(RED, "FAILED ", NORMAL, t.__name__)
                        print("RECURSION")
                except Exception as e:
                        print(RED, "FAILED ", NORMAL, t.__name__)
                        raise e

def __main__():
    test_all()

if __name__ == "__main__":
        test_all()
