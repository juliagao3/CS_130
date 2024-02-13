import unittest
import decimal
import sheets

from typing import Tuple

def to_excel_column(index: int) -> str:
    def divmod_excel(i: int) -> Tuple[int, int]:
        a, b = divmod(i, 26)
        if b == 0:
            return a - 1, b + 26
        return a, b
    result = []
    while index > 0:
        index, rem = divmod_excel(index)
        result.append(chr(ord("A") + rem - 1))
    return "".join(reversed(result))

def from_excel_column(s: str) -> int:
    s = s.upper()
    result = 0
    for c in s:
        result *= 26
        result += ord(c) - ord('A') + 1
    return result

def to_sheet_location(pos: Tuple[int, int]) -> str:
    return "{}{}".format(to_excel_column(pos[0]), pos[1])

def from_sheet_location(s: str) -> Tuple[int, int]:
    i = 0
    while not s[i].isdigit():
        i += 1
    return (from_excel_column(s[:i]), int(s[i:]))

def test_copy(self, wb: sheets.Workbook, sheet_name: str, to_sheet: str, start_tuple: Tuple[int, int], end_tuple: Tuple[int, int], to_tuple: Tuple[int, int]):
    if to_sheet is None:
        to_sheet = sheet_name

    size = (end_tuple[0] - start_tuple[0], end_tuple[1] - start_tuple[1])
    to_end_tuple = (to_tuple[0] + size[0], to_tuple[1] + size[1])

    for col in range(start_tuple[0], end_tuple[0] + 1):
        for row in range(start_tuple[1], end_tuple[1] + 1):
            location = to_sheet_location((col, row))
            wb.set_cell_contents(sheet_name, location, str((col - start_tuple[0]) + (row - start_tuple[1]) * size[0]))
            
    start_location = to_sheet_location(start_tuple)
    end_location = to_sheet_location(end_tuple)
    to_location = to_sheet_location(to_tuple)

    wb.copy_cells(sheet_name, start_location, end_location, to_location, to_sheet)

    for col in range(to_tuple[0], to_end_tuple[0] + 1):
        for row in range(to_tuple[1], to_end_tuple[1] + 1):
            location = to_sheet_location((col, row))
            contents = wb.get_cell_value(to_sheet, location)
            self.assertEqual(wb.get_cell_value(to_sheet, location), decimal.Decimal((col - to_tuple[0]) + (row - to_tuple[1]) * size[0]))

    for col in range(start_tuple[0], end_tuple[0] + 1):
        for row in range(start_tuple[1], end_tuple[1] + 1):
            if col <= to_end_tuple[0] and col >= to_tuple[0] and row <= to_end_tuple[1] and row >= to_tuple[1]:
                 continue
            location = to_sheet_location((col, row))
            self.assertEqual(wb.get_cell_value(sheet_name, location), decimal.Decimal((col - start_tuple[0]) + (row - start_tuple[1]) * size[0]))
            
def test_move(self, wb: sheets.Workbook, sheet_name: str, to_sheet: str, start_tuple: Tuple[int, int], end_tuple: Tuple[int, int], to_tuple: Tuple[int, int]):
    if to_sheet is None:
        to_sheet = sheet_name

    size = (end_tuple[0] - start_tuple[0], end_tuple[1] - start_tuple[1])
    to_end_tuple = (to_tuple[0] + size[0], to_tuple[1] + size[1])

    for col in range(start_tuple[0], end_tuple[0] + 1):
        for row in range(start_tuple[1], end_tuple[1] + 1):
            location = to_sheet_location((col, row))
            wb.set_cell_contents(sheet_name, location, str((col - start_tuple[0]) + (row - start_tuple[1]) * size[0]))
            
    start_location = to_sheet_location(start_tuple)
    end_location = to_sheet_location(end_tuple)
    to_location = to_sheet_location(to_tuple)

    wb.move_cells(sheet_name, start_location, end_location, to_location, to_sheet)

    for col in range(to_tuple[0], to_end_tuple[0] + 1):
        for row in range(to_tuple[1], to_end_tuple[1] + 1):
            location = to_sheet_location((col, row))
            contents = wb.get_cell_value(to_sheet, location)
            self.assertEqual(wb.get_cell_value(to_sheet, location), decimal.Decimal((col - to_tuple[0]) + (row - to_tuple[1]) * size[0]))

    for col in range(start_tuple[0], end_tuple[0] + 1):
        for row in range(start_tuple[1], end_tuple[1] + 1):
            if col <= to_end_tuple[0] and col >= to_tuple[0] and row <= to_end_tuple[1] and row >= to_tuple[1]:
                 continue
            location = to_sheet_location((col, row))
            self.assertEqual(wb.get_cell_value(sheet_name, location), None)

class TestClass(unittest.TestCase):

    def test_copy_many(self):
        start_location = (1, 1)
        end_location = (3, 3)
        to_location = (6, 1)

        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()

        wb.set_cell_contents(sheet_name, to_sheet_location(start_location), "1")
        for row in range(start_location[0], end_location[0] + 1):
            for col in range(start_location[1], end_location[1] + 1):
                test_move(self, wb, sheet_name, None, start_location, start_location, (row, col))

    def test_copy_cells_right(self):
        start_location = (1, 1)
        end_location = (3, 3)
        to_location = (6, 1)

        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()
        test_copy(self, wb, sheet_name, None, start_location, end_location, to_location)

        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()
        test_move(self, wb, sheet_name, None, start_location, end_location, to_location)
        
    def test_copy_cells_left(self):
        start_location = (10, 10)
        end_location = (13, 13)
        to_location = (1, 1)

        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()
        test_copy(self, wb, sheet_name, None, start_location, end_location, to_location)
        
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()
        test_move(self, wb, sheet_name, None, start_location, end_location, to_location)
        
    def test_copy_cells_no_move(self):
        start_location = (10, 10)
        end_location = (13, 13)
        to_location = (10, 10)

        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()
        test_copy(self, wb, sheet_name, None, start_location, end_location, to_location)
        
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()
        test_move(self, wb, sheet_name, None, start_location, end_location, to_location)
        
    def test_copy_cells_overlapping_right(self):
        start_location = (1, 1)
        end_location = (10, 10)
        to_location = (5, 5)

        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()
        test_copy(self, wb, sheet_name, None, start_location, end_location, to_location)
        
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()
        test_move(self, wb, sheet_name, None, start_location, end_location, to_location)
        
    def test_copy_cells_overlapping_left(self):
        start_location = (10, 10)
        end_location = (20, 20)
        to_location = (5, 5)

        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()
        test_copy(self, wb, sheet_name, None, start_location, end_location, to_location)
        
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()
        test_move(self, wb, sheet_name, None, start_location, end_location, to_location)
        
    def test_copy_cell_two_sheet(self):
        start_location = (1, 1)
        end_location = (3, 3)
        to_location = (6, 1)

        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()
        sheet_num, sheet_name1 = wb.new_sheet()
        test_copy(self, wb, sheet_name, sheet_name1, start_location, end_location, to_location)
        
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()
        test_move(self, wb, sheet_name, None, start_location, end_location, to_location)

    def test_move_same_sheet(self):
        wb = sheets.Workbook()
        num1, name1 = wb.new_sheet()
        num1, name2 = wb.new_sheet()
        
        start_tuple = (1, 1)
        end_tuple = (1, 2)
        to_tuple = (6, 1)
        
        wb.set_cell_contents(name1, "A1", "=15")
        wb.set_cell_contents(name1, "A2", "=A1")
        # wb.set_cell_contents(name2, "A1", "=Sheet2!A2")

        # actually move the cell
        #test_copy(self, wb, name1, name1, start_tuple, end_tuple, to_tuple)
        wb.copy_cells(name1, to_sheet_location(start_tuple), to_sheet_location(end_tuple), to_sheet_location(to_tuple))

        f2 = wb.get_cell_value(name1, to_sheet_location(to_tuple))
        # a2 = wb.get_cell_value(name2, "A2")

        self.assertEqual(f2, decimal.Decimal(15))

    def test_move_same_sheet_circ_ref(self):
        wb = sheets.Workbook()
        num, sheet = wb.new_sheet()
        num2, sheet2 = wb.new_sheet()
        
        start_tuple = (1, 1)
        end_tuple = (1, 1)
        to_tuple = (1, 1)
        
        wb.set_cell_contents(sheet, "A1", "=Sheet2!A1")
        
        wb.copy_cells(sheet, to_sheet_location(start_tuple), to_sheet_location(end_tuple), to_sheet_location(to_tuple), "sheet2")
        
        a1_sheet2 = wb.get_cell_value(sheet2, to_sheet_location(to_tuple))
        self.assertIsInstance(a1_sheet2, sheets.CellError)
        self.assertEqual(a1_sheet2.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        
        a1_sheet = wb.get_cell_value(sheet, to_sheet_location(to_tuple))
        self.assertIsInstance(a1_sheet, sheets.CellError)
        self.assertEqual(a1_sheet.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        
    def test_bad_reference(self):
        wb = sheets.Workbook()
        num, sheet = wb.new_sheet()
        
        end_tuple = (1, 2)
        to_tuple = (2, 1)
        
        wb.set_cell_contents(sheet, "A1", "10")
        wb.set_cell_contents(sheet, "A2", "=A1")
        
        # a2 = wb.get_cell_value(sheet, "A2")
        # self.assertIsInstance(a2, sheets.CellError)
        # self.assertEqual(a2.get_type(), sheets.CellErrorType.BAD_REFERENCE)
        
        # num, sheet2 = wb.new_sheet() 
        # wb.set_cell_contents(sheet2, "A1", "8")
        
        # a2_sheet2 = wb.get_cell_value(sheet, "A2")
        # self.assertEqual(a2_sheet2, decimal.Decimal(8))
        
        wb.copy_cells(sheet, to_sheet_location(end_tuple), to_sheet_location(end_tuple), to_sheet_location(to_tuple), sheet)
        
        b1 = wb.get_cell_value(sheet, "B1")
        self.assertIsInstance(b1, sheets.CellError)
        self.assertEqual(b1.get_type(), sheets.CellErrorType.BAD_REFERENCE)
      
    def test_move_absolute(self):
        wb = sheets.Workbook()
        num, name = wb.new_sheet()
        
        wb.set_cell_contents(name, "A1", "=$B$2")
        wb.set_cell_contents(name, "B2", "=123.0")

        old = wb.get_cell_value(name, "A1")
        wb.move_cells(name, "A1", "A1", "F10")
        new = wb.get_cell_value(name, "F10")

        self.assertEqual(old, new)

    def test_copy_relative(self):
        wb = sheets.Workbook()
        num, name = wb.new_sheet()

        wb.set_cell_contents(name, "A1", "=10")
        wb.set_cell_contents(name, "A2", "=A1")

        wb.copy_cells(name, "A1", "A2", "B1")
        wb.set_cell_contents(name, "A1", "a string")

        self.assertEqual(wb.get_cell_value(name, "B2"), decimal.Decimal(10))
        self.assertEqual(wb.get_cell_value(name, "A2"), "a string")
        self.assertEqual(wb.get_cell_contents(name, "B2"), "=b1")

    # test $ on only row/only column?

if __name__ == "__main__":
    unittest.main()