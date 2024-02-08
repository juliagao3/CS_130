import sheets
import decimal
import unittest

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
    return "{}{}".format(to_excel_column(pos[1]), pos[0])

def from_sheet_location(s: str) -> Tuple[int, int]:
    i = 0
    while not s[i].isdigit():
        i += 1
    return (int(s[i:]), from_excel_column(s[:i]))

def test_error(test, wb, sheet_name, location, error_type):
    value = wb.get_cell_value(sheet_name, location)
    test.assertIsInstance(value, sheets.CellError)
    test.assertEqual(value.get_type(), error_type)

def tree_workbook(branching_factor, levels): 
    wb = sheets.Workbook()
    index, name = wb.new_sheet()
    
    total_nodes = (branching_factor**levels - 1) // (branching_factor - 1)
    root_index = 2
    
    wb.set_cell_contents(name, f"A{root_index}", "=A1")
    
    for i in range(root_index + 1, root_index + total_nodes - 1): 
        parent = (i - root_index - 1) // branching_factor + root_index 
        wb.set_cell_contents(name, f"A{i}", f"=A{parent}") 
        
    return (wb, index, name)

class TestClass(unittest.TestCase):    

    # def test_big_cycle(self):
    #     wb = sheets.Workbook()
    #     index, name = wb.new_sheet()

    #     for i in range(2,1000):
    #         wb.set_cell_contents(name, "A" + str(i), "=A" + str(i-1))

    #     wb.set_cell_contents(name, "A1", "1.0")

    #     self.assertEqual(wb.get_cell_value(name, "A999"), decimal.Decimal(1.0))

    #     wb.set_cell_contents(name, "A1", "=A999")

    #     a1 = wb.get_cell_value(name, "A1")
    #     self.assertEqual(type(a1), sheets.CellError)
    #     self.assertEqual(a1.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

    # def test_smol_cycles(self):
    #     wb = sheets.Workbook()
    #     index, name = wb.new_sheet()

    #     cycle_count = 100
    #     cycle_length = 3

    #     for i in range(1, 301, 3):
    #         wb.set_cell_contents(name, "A" + str(i), "=A" + str(i + 1))
    #         wb.set_cell_contents(name, "A" + str(i + 1), "=A" + str(i + 2))
    #         wb.set_cell_contents(name, "A" + str(i + 2), "=A" + str(i))
    #         test_error(self, wb, name, "A" + str(i), sheets.CellErrorType.CIRCULAR_REFERENCE)
    #         test_error(self, wb, name, "A" + str(i + 1), sheets.CellErrorType.CIRCULAR_REFERENCE)
    #         test_error(self, wb, name, "A" + str(i + 2), sheets.CellErrorType.CIRCULAR_REFERENCE)
        
    #     for i in range(1, 301, 3):
    #         wb.set_cell_contents(name, "A" + str(i), "0")
    #         self.assertEqual(wb.get_cell_value(name, "A" + str(i)), decimal.Decimal(0))
    #         self.assertEqual(wb.get_cell_value(name, "A" + str(i + 1)), decimal.Decimal(0))
    #         self.assertEqual(wb.get_cell_value(name, "A" + str(i + 2)), decimal.Decimal(0))

    #         for j in range(i + 3, 301, 3):
    #             test_error(self, wb, name, "A" + str(j), sheets.CellErrorType.CIRCULAR_REFERENCE)
                
    # def test_error_propogation(self):
    #     wb = sheets.Workbook()
    #     num, name = wb.new_sheet()
    #     num1, name1 = wb.new_sheet()
        
    #     wb.set_cell_contents(name, "A1", f"={name1}!A1")
    #     for i in range(2, 1000):
    #         wb.set_cell_contents(name, "A" + str(i), "=A1")
        
    #     wb.del_sheet(name1)
        
    #     for i in range(2, 1000):
    #         test_error(self, wb, name, "A" + str(i), sheets.CellErrorType.BAD_REFERENCE)
            
    # def test_tree(self):
    #     wb, index, name = tree_workbook(3, 2)

    #     updated = []
    
    #     def on_update(wb, locations):
    #         nonlocal updated
    #         for l in locations:
    #             updated.append(l)
        
    #     wb.notify_cells_changed(on_update)
    #     wb.set_cell_contents(name, "A1", "0")

    #     self.assertEqual(wb.get_cell_value(name, "A1"), decimal.Decimal(0))
       
    #1000 cells, precision up to 250
    # def test_fibonacci(self):
    #     wb = sheets.Workbook()
    #     num, name = wb.new_sheet()
        
    #     for i in range(3, 500):
    #         if i % 50 == 0:
    #             print("setting ", i)
    #         wb.set_cell_contents(name, "A" + str(i), "=A" + str(i-1) + " +" + " A" + str(i-2))
            
    #     wb.set_cell_contents(name, "A1", "=1")
    #     wb.set_cell_contents(name, "A2", "=1")

    #     # with open("wb.wb", "w") as f:
    #     #     wb.save_workbook(f)

    #     for i in range(3, 500):
    #         if i % 50 == 0:
    #             print("checking ", i)
    #         value_1 = wb.get_cell_value(name, "A" + str(i - 2))
    #         value_2 = wb.get_cell_value(name, "A" + str(i - 1))
    #         self.assertEqual(wb.get_cell_value(name, "A" + str(i)), value_1 + value_2)
      
    def test_pascal(self):
        wb = sheets.Workbook()
        num, name = wb.new_sheet()

        # 1
        # 1 1
        # 1 2 1
        # 1 3 3 1

        rows = 100
        for r in range(3, rows + 1):
            for c in range(2, r):
                location = to_sheet_location((r, c))
                parent1 = to_sheet_location((r-1, c))
                parent2 = to_sheet_location((r-1, c-1))
                wb.set_cell_contents(name, location, f"={parent1} + {parent2}")

        rows = 100
        wb.set_cell_contents(name, "A1", "1")
        for r in range(2, rows + 1):
            first = to_sheet_location((r, 1))
            last = to_sheet_location((r, r))
            wb.set_cell_contents(name, first, "1")
            wb.set_cell_contents(name, last, "1")
        
        with open("wb.wb", "w") as f:
            wb.save_workbook(f)

    # def test_large_formula(self):  
    #     wb = sheets.Workbook()
    #     index, name = wb.new_sheet()
    #     x = 2
    #     wb.set_cell_contents("sheet1", "a1", str(x))
    #     k = 10
    #     for i in range(2, k+1):
    #         formula = "="
    #         for j in range(1, i-1):
    #             formula += "A" + str(j) + "*"
    #         formula += "A" + str(i-1)
    #         wb.set_cell_contents("sheet1", "a" + str(i), formula)

    #     x = decimal.Decimal(x)
    #     k = decimal.Decimal(k)
    #     self.assertEqual(wb.get_cell_value(name, f"A{k}"), x**(2**(k-2)))

if __name__ == "__main__":
    unittest.main(module="stress")
