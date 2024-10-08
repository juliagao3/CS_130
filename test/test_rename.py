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
        
        self.assertEqual(wb.get_cell_contents(new_name, "A1").lower(), f"='{new_name}'!a2")
        
    def test_formula_without_quotes(self):
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet("sheet1")
        
        wb.set_cell_contents(sheet_name, "A1", "=sheet1!A2")
        new_name = "sheet2"
        wb.rename_sheet(sheet_name, new_name)
        
        self.assertEqual(wb.get_cell_contents(new_name, "A1").lower(), f"={new_name}!a2")
        
    def test_formula_add_quotes(self):
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet("sheet1")
        sheet_num1, sheet_name1 = wb.new_sheet("sheet2")
        
        wb.set_cell_contents(sheet_name, "A1", "=5")
        wb.set_cell_contents(sheet_name1, "A1", "=sheet1!A1 + sheet2!A2")
        
        new_name = "sheet bla"
        wb.rename_sheet(sheet_name1, new_name)
        
        self.assertEqual(wb.get_cell_contents(new_name, "A1").lower(), f"='{sheet_name}'!a1 + '{new_name}'!a2")
        
    def test_formula_add_quotes(self):
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet("sheet1")
        sheet_num1, sheet_name1 = wb.new_sheet("sheet2")
        
        wb.set_cell_contents(sheet_name, "A1", "=5")
        wb.set_cell_contents(sheet_name1, "A1", "=sheet1!A1 + sheet2!A2")
        
        new_name = "sheet3"
        wb.rename_sheet(sheet_name1, new_name)
        
        self.assertEqual(wb.get_cell_contents(new_name, "A1").lower(), f"={sheet_name}!a1 + {new_name}!a2")
        
    def test_formula_add_quotes_2(self):
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet("sheet 1")
        sheet_num1, sheet_name1 = wb.new_sheet("sheet2")
        
        wb.set_cell_contents(sheet_name, "A1", "=5")
        wb.set_cell_contents(sheet_name1, "A1", "='sheet 1'!A1 + sheet2!A2")
        
        new_name = "sheet3"
        wb.rename_sheet(sheet_name1, new_name)
        
        self.assertEqual(wb.get_cell_contents(new_name, "A1").lower(), f"='{sheet_name}'!a1 + {new_name}!a2")
        
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
        
    def test_self_reference(self):
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()

        wb.set_cell_contents(sheet_name, "A1", f"='{sheet_name}'!A2")
        wb.set_cell_contents(sheet_name, "A2", "1")

        self.assertEqual(wb.get_cell_value(sheet_name, "A1"), decimal.Decimal(1))
        
        new_name = "hi"
        wb.rename_sheet(sheet_name, new_name)

        self.assertEqual(wb.get_cell_value(new_name, "A1"), decimal.Decimal(1))

    def test_other_reference(self):
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet()

        wb.set_cell_contents(sheet_name, "A1", f"='{sheet_name}'!A2")
        wb.set_cell_contents(sheet_name, "A2", "1")

    def test_concentation(self):
        wb = sheets.Workbook()
        sheet_num, name = wb.new_sheet("Somesheet")
        
        wb.set_cell_contents(name, "A1", "=Somesheet!A2 & \"hello\"")
        wb.rename_sheet(name, "Another")
        self.assertEqual(wb.get_cell_value("Another", "A1"), "hello")
        self.assertEqual(wb.get_cell_contents("Another", "A1").lower(), "=another!a2 & \"hello\"")
        
    def test_remove_quotes(self):
        wb = sheets.Workbook()
        sheet_num, name = wb.new_sheet()
        sheet_num2, name2 = wb.new_sheet("sheet@bla")
        
        wb.set_cell_contents(name, "A1", "='sheet@bla'!A1")
        wb.rename_sheet(name2, "sheet2")
        self.assertEqual(wb.get_cell_value(name, "A1"), decimal.Decimal(0))
    
if __name__ == "__main__":
        unittest.main()
