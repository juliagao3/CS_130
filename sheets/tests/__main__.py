#! /usr/bin/env python3

import sys
import sheets
import decimal
import traceback

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

def test_all():
        tests = [
                test_default_sheet_name,
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
                test_cell_update_multiple
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
