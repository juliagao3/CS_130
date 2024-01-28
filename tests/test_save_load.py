# python -m unittest tests.test_save_load

import unittest
import unittest.mock

import sys
import sheets
import decimal
import traceback
import re
import io
import json

class TestClass(unittest.TestCase):    

        def test_save_workbook(self):
                wb = sheets.Workbook()
                sheet_num1, sheet_name1 = wb.new_sheet()
                sheet_num2, sheet_name2 = wb.new_sheet()

                wb.set_cell_contents(sheet_name1, "A1", "'123")
                wb.set_cell_contents(sheet_name1, "B1", "5.3")
                wb.set_cell_contents(sheet_name1, "C1", "=A1*B1")

                wb.set_cell_contents(sheet_name2, "A1", "Hello")

                f = open("test_file.txt", "w")
                wb.save_workbook(f)
                f.close()

                file = open("test_file.txt", "r")
                file_read = file.read()

                expected_json = {
                    "sheets": [
                        {
                            "name": "Sheet1",
                            "cell-contents": {
                                "a1": "'123",
                                "b1": "5.3",
                                "c1": "=A1*B1"
                            }
                        },
                        {
                            "name": "Sheet2",
                            "cell-contents": {
                                "a1": "Hello"
                            }
                        }
                    ]
                }
                expected = json.dumps(expected_json, indent=4)

                self.assertEqual(file_read, expected)

        def test_save_empty_workbook(self):
                wb = sheets.Workbook()
                f = open("test_file.txt", "w")
                wb.save_workbook(f)
                f.close()

                file = open("test_file.txt", "r")
                file_read = file.read()

                expected_json = {
                    "sheets": []
                        }
                expected = json.dumps(expected_json, indent=4)

                self.assertEqual(file_read, expected)
        
        def test_load_workbook(self):
                
                test_json = {
                    "sheets":[
                        {
                            "name":"Sheet1",
                            "cell-contents":{
                                "A1":"'test contents"
                            }                                        
                        }
                    ]
                }

                f = open("test_file.txt", "w")
                json.dump(test_json, f, indent=4)
                f.close()

                fp = open("test_file.txt", "r")
                wb1 = sheets.Workbook.load_workbook(fp)
                fp.close()

                wb2 = sheets.Workbook("manual")
                wb2.new_sheet("Sheet1")
                wb2.set_cell_contents("Sheet1", "A1", "'test contents")

                loaded_contents = wb1.get_cell_contents("Sheet1", "A1")
                loaded_value = wb1.get_cell_value("Sheet1", "A1")
                manual_contents = wb2.get_cell_contents("Sheet1", "A1")
                manual_value = wb2.get_cell_value("Sheet1", "A1")

                self.assertEqual(loaded_contents, manual_contents)
                self.assertEqual(loaded_value, manual_value)
                pass

        def test_key_error(self):
                test_json = {
                    "sheets":[
                        {
                            "name":"Sheet1",
                            "spelling-error":{
                                "A1":"asdkfj"
                            }                                        
                        }
                    ]
                }

                f = open("test_file.txt", "w")
                json.dump(test_json, f, indent=4)
                f.close()

                fp = open("test_file.txt", "r")

                with self.assertRaises(KeyError):
                        sheets.Workbook.load_workbook(fp)
                fp.close()

        def test_type_error(self):
                test_json = {
                    "sheets":[
                        {
                            "name":12345,
                            "cell-contents":{
                                "A1":"asdkfj"
                            }                                        
                        }
                    ]
                }

                f = open("test_file.txt", "w")
                json.dump(test_json, f, indent=4)
                f.close()

                fp = open("test_file.txt", "r")

                with self.assertRaises(TypeError):
                        sheets.Workbook.load_workbook(fp)
                fp.close()

        def test_sheet_name_not_str(self):
                test_json = {
                    "sheets":[
                        {
                            "name":["not", "a", "string"],
                            "":{
                                "A1":"asdkfj"
                            }                                        
                        }
                    ]
                }

                f = open("test_file.txt", "w")
                json.dump(test_json, f, indent=4)
                f.close()

                fp = open("test_file.txt", "r")

                with self.assertRaises(TypeError):
                        sheets.Workbook.load_workbook(fp)
                fp.close()
        
        def test_cell_contents_not_str(self):
                test_json = {
                    "sheets":[
                        {
                            "name":"Sheet1",
                            "cell-contents":{
                                "A1":[12345]
                            }                                        
                        }
                    ]
                }

                f = open("test_file.txt", "w")
                json.dump(test_json, f, indent=4)
                f.close()

                fp = open("test_file.txt", "r")

                with self.assertRaises(TypeError):
                        sheets.Workbook.load_workbook(fp)
                fp.close()

        def test_escape_double_quotes_save(self):
                wb = sheets.Workbook()
                sheet_num, sheet_name = wb.new_sheet()
                wb.set_cell_contents(sheet_name, "A1", "double \" quotes")
                a1_contents = wb.get_cell_contents(sheet_name, "A1")
                a1_value = wb.get_cell_value(sheet_name, "A1")

                f = open("test_file.txt", "w")
                wb.save_workbook(f)
                f.close()

                file = open("test_file.txt", "r")
                wb_load = sheets.Workbook.load_workbook(file)
                loaded_contents = wb_load.get_cell_contents(sheet_name, "A1")
                loaded_value = wb_load.get_cell_value(sheet_name, "A1")

                self.assertEqual(a1_contents, loaded_contents)
                self.assertEqual(a1_value, loaded_value)

        def test_escape_double_quotes_load(self):
                test_json = {
                    "sheets":[
                        {
                            "name":"Sheet1",
                            "cell-contents":{
                                "A1":"'escape \" double \" quotes"
                            }                                        
                        }
                    ]
                }

                f = open("test_file.txt", "w")
                json.dump(test_json, f, indent=4)
                f.close()

                file = open("test_file.txt", "r")
                wb = sheets.Workbook.load_workbook(file)
                a1_contents = wb.get_cell_contents("Sheet1", "A1")
                a1_value = wb.get_cell_value("Sheet1", "A1")
                file.close()

                self.assertEqual(a1_contents, "'escape \" double \" quotes")
                self.assertEqual(a1_value, "escape \" double \" quotes")          

if __name__ == "__main__":
        unittest.main()