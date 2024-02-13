import unittest
import decimal
import sheets

class TestClass(unittest.TestCase):

    def test_correct_new_location(self):
        wb = sheets.Workbook()
        sheet_num1, sheet_name1 = wb.new_sheet("Spreadsheet 1")
        sheet_num2, sheet_name2 = wb.new_sheet("Spreadsheet 2")
        sheet_num3, sheet_name3 = wb.new_sheet("Spreadsheet 3")

        wb.set_cell_contents(sheet_name1, "C4", "=B2 + 5.0")
        wb.set_cell_contents(sheet_name1, "B2", "=10*20")
        wb.set_cell_contents(sheet_name2, "A1", "=Spreadsheet 1!B2")

        old_c4 = wb.get_cell_value(sheet_name1, "C4")
        old_b2 = wb.get_cell_value(sheet_name1, "B2")
        old_a1 = wb.get_cell_value(sheet_name2, "A1")
        
        wb.move_sheet(sheet_name1, 1)

        new_list = wb.list_sheets()
        self.assertEqual(new_list, ["'Spreadsheet 2'", "'Spreadsheet 1'", "'Spreadsheet 3'"])

        self.assertEqual(wb.get_cell_value(sheet_name1, "C4"), old_c4)
        self.assertEqual(wb.get_cell_value(sheet_name1, "B2"), old_b2)
        self.assertEqual(wb.get_cell_value(sheet_name2, "A1"), old_a1)

    def test_invalid_move_index(self):
        wb = sheets.Workbook()
        sheet_num1, sheet_name1 = wb.new_sheet("Spreadsheet 1")

        with self.assertRaises(IndexError):
            wb.move_sheet(sheet_name1, -1)

        with self.assertRaises(IndexError):
            wb.move_sheet(sheet_name1, 1)

    def test_invalid_name(self):
        wb = sheets.Workbook()
        sheet_num1, sheet_name1 = wb.new_sheet("Spreadsheet 1")

        with self.assertRaises(KeyError):
            wb.move_sheet("askdjfhalksdjfh", 0)
            
    def test_deep_copy(self):
        wb = sheets.Workbook()
        sheet_num1, sheet_name1 = wb.new_sheet("Spreadsheet 1")
        sheet_num2, sheet_name2 = wb.new_sheet("Spreadsheet 2")

        wb.set_cell_contents(sheet_name1, "A1", "=100.0")
        wb.set_cell_contents(sheet_name1, "A2", "='Spreadsheet 2'!B1 + 1.2")
        wb.set_cell_contents(sheet_name2, "B1", "=25")

        copy_num, copy_name = wb.copy_sheet(sheet_name1)
        self.assertEqual(wb.get_cell_value(copy_name, "A1"), wb.get_cell_value(sheet_name1, "A1"))
        self.assertEqual(wb.get_cell_contents(copy_name, "A1"), wb.get_cell_contents(copy_name, "A1"))
        self.assertEqual(wb.get_cell_value(copy_name, "A2"), wb.get_cell_value(sheet_name1, "A2"))
        self.assertEqual(wb.get_cell_contents(copy_name, "A2"), wb.get_cell_contents(sheet_name1, "A2"))
        self.assertEqual(wb.get_cell_value(sheet_name2, "B1"), decimal.Decimal(25))
        
        self.assertEqual(copy_num, 2)
        self.assertEqual(wb.list_sheets(), ["'Spreadsheet 1'", "'Spreadsheet 2'", "'Spreadsheet 1_1'"])

        wb.set_cell_contents(copy_name, "A1", "something different")
        self.assertNotEqual(wb.get_cell_contents(sheet_name1, "A1"), wb.get_cell_contents(copy_name, "A1"))
        self.assertNotEqual(wb.get_cell_value(sheet_name1, "A1"), wb.get_cell_value(copy_name, "A1"))

    def test_copy_names(self):
        wb = sheets.Workbook()
        sheet_num1, sheet_name1 = wb.new_sheet("Sheet1")

        copy1_num1, sheet1_copy1 = wb.copy_sheet(sheet_name1)
        copy1_num2, sheet1_copy2 = wb.copy_sheet(sheet_name1)
        copy1_copy1_num, sheet1_copy1_copy = wb.copy_sheet(sheet1_copy1)

        self.assertEqual(sheet1_copy1, "Sheet1_1")
        self.assertEqual(sheet1_copy2, "Sheet1_2")
        self.assertEqual(sheet1_copy1_copy, "Sheet1_1_1")
        
    def test_bad_reference(self):
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet("sheet1")
        
        wb.set_cell_contents(sheet_name, "A1", f"={sheet_name}_1!A1")
        
        value = wb.get_cell_value(sheet_name, "A1")
        self.assertIsInstance(value, sheets.CellError)
        self.assertEqual(value.get_type(), sheets.CellErrorType.BAD_REFERENCE)
        
        copy_num, copy_name = wb.copy_sheet(sheet_name)
        # wb.set_cell_contents(copy_name, "A1", f"={sheet_name}!A1")
        
        value1 = wb.get_cell_value(copy_name, "A1")
        self.assertIsInstance(value1, sheets.CellError)
        self.assertEqual(value1.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        
        wb.set_cell_contents(copy_name, "A1", "5")
        self.assertEqual(wb.get_cell_value(copy_name, "A1"), decimal.Decimal(5))       
        
    def test_copy_error(self):
        wb = sheets.Workbook()
        sheet_num, sheet_name = wb.new_sheet("sheet1")

        wb.set_cell_contents(sheet_name, "A2", "=A1")
        wb.set_cell_contents(sheet_name, "A1", f"={sheet_name}_1!A2")

        value1 = wb.get_cell_value(sheet_name, "A1")
        self.assertIsInstance(value1, sheets.CellError)
        self.assertEqual(value1.get_type(), sheets.CellErrorType.BAD_REFERENCE)

        copy_num, copy_name = wb.copy_sheet(sheet_name)

        value1 = wb.get_cell_value(copy_name, "A1")
        self.assertIsInstance(value1, sheets.CellError)
        self.assertEqual(value1.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        # self.assertEqual(wb.get_cell_value(copy_name, "A1"), decimal.Decimal(2))


if __name__ == "__main__":
        unittest.main()