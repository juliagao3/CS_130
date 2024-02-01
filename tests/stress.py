import sheets
import decimal
import unittest

def test_error(test, wb, sheet_name, location, error_type):
    value = wb.get_cell_value(sheet_name, location)
    test.assertIsInstance(value, sheets.CellError)
    test.assertEqual(value.get_type(), error_type)

def tree_workbook(branching_factor, levels):
    wb = sheets.Workbook()
    index, name = wb.new_sheet()

    total_nodes = branching_factor ** levels - 1
    root_index = branching_factor - 1

    if root_index != 1:
        wb.set_cell_contents(name, f"A{root_index}", "=A1")

    for i in range(root_index + 1, root_index + total_nodes):
        parent = i // branching_factor
        wb.set_cell_contents(name, f"A{i}", f"=A{parent}")

    return (wb, index, name)

class TestClass(unittest.TestCase):    

    def test_big_cycle(self):
        wb = sheets.Workbook()
        index, name = wb.new_sheet()

        for i in range(2,1000):
            wb.set_cell_contents(name, "A" + str(i), "=A" + str(i-1))

        wb.set_cell_contents(name, "A1", "1.0")

        self.assertEqual(wb.get_cell_value(name, "A999"), decimal.Decimal(1.0))

        wb.set_cell_contents(name, "A1", "=A999")

        a1 = wb.get_cell_value(name, "A1")
        self.assertEqual(type(a1), sheets.CellError)
        self.assertEqual(a1.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

    def test_smol_cycles(self):
        wb = sheets.Workbook()
        index, name = wb.new_sheet()

        cycle_count = 100
        cycle_length = 3

        for i in range(1, 301, 3):
            wb.set_cell_contents(name, "A" + str(i), "=A" + str(i + 1))
            wb.set_cell_contents(name, "A" + str(i + 1), "=A" + str(i + 2))
            wb.set_cell_contents(name, "A" + str(i + 2), "=A" + str(i))
            test_error(self, wb, name, "A" + str(i), sheets.CellErrorType.CIRCULAR_REFERENCE)
            test_error(self, wb, name, "A" + str(i + 1), sheets.CellErrorType.CIRCULAR_REFERENCE)
            test_error(self, wb, name, "A" + str(i + 2), sheets.CellErrorType.CIRCULAR_REFERENCE)
        
        for i in range(1, 301, 3):
            wb.set_cell_contents(name, "A" + str(i), "0")
            self.assertEqual(wb.get_cell_value(name, "A" + str(i)), decimal.Decimal(0))
            self.assertEqual(wb.get_cell_value(name, "A" + str(i + 1)), decimal.Decimal(0))
            self.assertEqual(wb.get_cell_value(name, "A" + str(i + 2)), decimal.Decimal(0))

            for j in range(i + 3, 301, 3):
                test_error(self, wb, name, "A" + str(j), sheets.CellErrorType.CIRCULAR_REFERENCE)
                
    def test_error_propogation(self):
        wb = sheets.Workbook()
        num, name = wb.new_sheet()
        num1, name1 = wb.new_sheet()
        
        wb.set_cell_contents(name, "A1", f"={name1}!A1")
        for i in range(2, 1000):
            wb.set_cell_contents(name, "A" + str(i), "=A1")
        
        wb.del_sheet(name1)
        
        for i in range(2, 1000):
            test_error(self, wb, name, "A" + str(i), sheets.CellErrorType.BAD_REFERENCE)
            
    def test_tree(self):
        wb, index, name = tree_workbook(3, 2)

        updated = []
    
        def on_update(wb, locations):
            nonlocal updated
            for l in locations:
                updated.append(l)
        
        wb.notify_cells_changed(on_update)
        wb.set_cell_contents(name, "A1", "0")

        print(updated)
      
    def test_large_formula(self):  
        wb = sheets.Workbook()
        index, name = wb.new_sheet()
        x = 2
        wb.set_cell_contents("sheet1", "a1", str(x))
        k = 10
        for i in range(2, k+1):
            formula = "="
            for j in range(1, i-1):
                formula += "A" + str(j) + "*"
            formula += "A" + str(i-1)
            wb.set_cell_contents("sheet1", "a" + str(i), formula)

        x = decimal.Decimal(x)
        k = decimal.Decimal(k)
        self.assertEqual(wb.get_cell_value(name, f"A{k}"), x**(2**(k-2)))

if __name__ == "__main__":
    unittest.main()