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

                f = open("myfile.txt", "w")
                wb.save_workbook(f)
                f.close()

                file = open("myfile.txt", "r")
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
                f = open("myfile2.txt", "w")
                wb.save_workbook(f)
                f.close()

                file = open("myfile2.txt", "r")
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

                f = open("myfile3.txt", "w")
                json.dump(test_json, f, indent=4)
                f.close()

                fp = open("myfile3.txt", "r")
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

                f = open("myfile3.txt", "w")
                json.dump(test_json, f, indent=4)
                f.close()

                fp = open("myfile3.txt", "r")

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

                f = open("myfile3.txt", "w")
                json.dump(test_json, f, indent=4)
                f.close()

                fp = open("myfile3.txt", "r")

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

                f = open("myfile3.txt", "w")
                json.dump(test_json, f, indent=4)
                f.close()

                fp = open("myfile3.txt", "r")

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

                f = open("myfile3.txt", "w")
                json.dump(test_json, f, indent=4)
                f.close()

                fp = open("myfile3.txt", "r")

                with self.assertRaises(TypeError):
                        sheets.Workbook.load_workbook(fp)
                fp.close()

        def test_escape_double_quotes(self):
                pass

if __name__ == "__main__":
        unittest.main()