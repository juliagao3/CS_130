import unittest
import sheets
import decimal


class TestClass(unittest.TestCase):
    
    def test_rename(self):
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet("sheet1")
        
        wb.rename_sheet(sheet_name, "sheet bla")
        self.assertEqual(wb.list_sheets(), ["'sheet bla'"])  
        
    def test_formula_with_quotes(self):
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet("sheet1")
        
        wb.set_cell_contents(sheet_name, "A1", "=sheet1!A2")
        new_name = "sheet bla"
        wb.rename_sheet(sheet_name, new_name) 
        
        self.assertEqual(wb.get_cell_contents(new_name, "A1"), f"='{new_name}'!A2")
        
    def test_formula_without_quotes(self):
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet("sheet1")
        
        wb.set_cell_contents(sheet_name, "A1", "=sheet1!A2")
        new_name = "sheet2"
        wb.rename_sheet(sheet_name, new_name)
        
        self.assertEqual(wb.get_cell_contents(new_name, "A1"), f"={new_name}!A2")
        
    def test_formula_add_quotes(self):
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet("sheet1")
        sheet_num1, sheet_name1 = wb.new_sheet("sheet2")
        
        wb.set_cell_contents(sheet_name, "A1", "=5")
        wb.set_cell_contents(sheet_name1, "A1", "=sheet1!A1 + sheet2!A2")
        
        new_name = "sheet bla"
        wb.rename_sheet(sheet_name1, new_name)
        
        self.assertEqual(wb.get_cell_contents(new_name, "A1"), f"='{sheet_name}'!A1 + '{new_name}'!A2")
        
    def test_formula_add_quotes(self):
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet("sheet1")
        sheet_num1, sheet_name1 = wb.new_sheet("sheet2")
        
        wb.set_cell_contents(sheet_name, "A1", "=5")
        wb.set_cell_contents(sheet_name1, "A1", "=sheet1!A1 + sheet2!A2")
        
        new_name = "sheet3"
        wb.rename_sheet(sheet_name1, new_name)
        
        self.assertEqual(wb.get_cell_contents(new_name, "A1"), f"={sheet_name}!A1 + {new_name}!A2")
        
    def test_formula_add_quotes_2(self):
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet("sheet 1")
        sheet_num1, sheet_name1 = wb.new_sheet("sheet2")
        
        wb.set_cell_contents(sheet_name, "A1", "=5")
        wb.set_cell_contents(sheet_name1, "A1", "='sheet 1'!A1 + sheet2!A2")
        
        new_name = "sheet3"
        wb.rename_sheet(sheet_name1, new_name)
        
        self.assertEqual(wb.get_cell_contents(new_name, "A1"), f"='{sheet_name}'!A1 + {new_name}!A2")
        
    def test_get_value(self):
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet("sheet 1")
        sheet_num1, sheet_name1 = wb.new_sheet("sheet2")
        
        wb.set_cell_contents(sheet_name, "A1", "=sheEt2!A1")
        wb.set_cell_contents(sheet_name, "A3", "='sheet 1'!A1 + sheet3!A1")
        wb.set_cell_contents(sheet_name1, "A1", "4")
        wb.set_cell_contents(sheet_name, "A4", "=A3")
        
        self.assertIsInstance(wb.get_cell_value(sheet_name, "A4"), sheets.CellError)
        self.assertEqual(wb.get_cell_value(sheet_name, "A4").get_type(), sheets.CellErrorType.BAD_REFERENCE)
        
        new_name = "sheet3"
        wb.rename_sheet(sheet_name1, new_name)
        self.assertEqual(wb.get_cell_value(sheet_name, "A1"), decimal.Decimal(4))
        self.assertEqual(wb.get_cell_value(sheet_name, "A3"), decimal.Decimal(8))
        self.assertEqual(wb.get_cell_value(sheet_name, "A4"), decimal.Decimal(8))
    
if __name__ == "__main__":
        unittest.main()
