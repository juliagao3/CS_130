import sheets
import decimal

from sheets import CellError, CellErrorType

import lark
from lark.visitors import visit_children_decor

def test_default_sheet_name():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)
        assert sheet_name == "Sheet1"

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
        assert type(a1) == CellError
        assert a1.get_type() == CellErrorType.TYPE_ERROR
        
def test_one_minus_unary_string():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", '=1 - -"hello"')
        a1 = wb.get_cell_value(sheet_name, "A1")
        assert type(a1) == CellError
        assert a1.get_type() == CellErrorType.TYPE_ERROR

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
        assert wb.get_cell_value(sheet_name, "A2") == decimal.Decimal(1)

        a3 = wb.get_cell_value(sheet_name, "A3")

        assert type(a3) == sheets.CellError
        assert a3.get_type() == sheets.CellErrorType.TYPE_ERROR

def test_all():
        tests = [
                test_default_sheet_name,
                test_one_plus_one,
                test_one_plus_string,
                test_one_minus_unary_string
        ]

        for t in tests:
                t()

if __name__ == "__main__":
        test_all()

