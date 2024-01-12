#! /usr/bin/env python3

import sheets
import decimal

def test_default_sheet_name():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)
        assert sheet_name == "Sheet1"

def test_empty_sheet():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)
        assert wb.get_cell_value(sheet_name, "A1") == None

def test_one_plus_one():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", "=1+1")
        assert wb.get_cell_value(sheet_name, "A1") == decimal.Decimal(2)
        
def test_one_plus_string():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", '=1+ "hello"')
        a1 = wb.get_cell_value(sheet_name, "A1")
        assert type(a1) == sheets.CellError
        assert a1.get_type() == sheets.CellErrorType.TYPE_ERROR
        
def test_one_minus_unary_string():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", '=1 - -"hello"')
        a1 = wb.get_cell_value(sheet_name, "A1")
        assert type(a1) == sheets.CellError
        assert a1.get_type() == sheets.CellErrorType.TYPE_ERROR

def test_one_plus_empty():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", "1")
        wb.set_cell_contents(sheet_name, "A3", "=A1+A2")

        assert wb.get_cell_value(sheet_name, "A1") == decimal.Decimal(1)
        assert wb.get_cell_value(sheet_name, "A3") == decimal.Decimal(1)

def test_one_plus_one_cells():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", "1")
        wb.set_cell_contents(sheet_name, "A2", "1")
        wb.set_cell_contents(sheet_name, "A3", "=A1+A2")

        assert wb.get_cell_value(sheet_name, "A1") == decimal.Decimal(1)
        assert wb.get_cell_value(sheet_name, "A2") == decimal.Decimal(1)
        assert wb.get_cell_value(sheet_name, "A3") == decimal.Decimal(2)

def test_string_arithmetic_cells():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", "1")
        wb.set_cell_contents(sheet_name, "A2", "hello")
        wb.set_cell_contents(sheet_name, "A3", "=A1+A2")

        assert wb.get_cell_value(sheet_name, "A1") == decimal.Decimal(1)

        a3 = wb.get_cell_value(sheet_name, "A3")

        assert type(a3) == sheets.CellError
        assert a3.get_type() == sheets.CellErrorType.TYPE_ERROR

def test_cell_update():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", "1")
        wb.set_cell_contents(sheet_name, "A2", "1")
        wb.set_cell_contents(sheet_name, "A3", "=A1+A2")

        assert wb.get_cell_value(sheet_name, "A3") == decimal.Decimal(2)

        wb.set_cell_contents(sheet_name, "A2", "2")

        assert wb.get_cell_value(sheet_name, "A3") == decimal.Decimal(3)

def test_circular_refs():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", "=A2")
        wb.set_cell_contents(sheet_name, "A2", "=A1")

        a1 = wb.get_cell_value(sheet_name, "A1")
        a2 = wb.get_cell_value(sheet_name, "A1")

        assert type(a1) == sheets.CellError
        assert type(a2) == sheets.CellError

        assert a1.get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE
        assert a2.get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE

def test_circular_refs_with_tail():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", "=A2")
        wb.set_cell_contents(sheet_name, "A2", "=A1+A4")
        wb.set_cell_contents(sheet_name, "A3", "=\"Hello \" & A1 & \"!\"")

        a1 = wb.get_cell_value(sheet_name, "A1")
        a2 = wb.get_cell_value(sheet_name, "A2")
        a3 = wb.get_cell_value(sheet_name, "A3")

        assert type(a1) == sheets.CellError
        assert type(a2) == sheets.CellError
        assert type(a3) == str

        assert a1.get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE
        assert a2.get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE
        assert a3.startswith("Hello")

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
                test_cell_update,
        ]

        for t in tests:
                t()
        
        print("All tests pass!")

def __main__():
    test_all()

if __name__ == "__main__":
        test_all()
