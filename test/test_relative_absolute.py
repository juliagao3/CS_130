#! /usr/bin/env python3
import unittest
import unittest.mock

import sys
import sheets
import decimal
import traceback
import re
import io

class TestClass(unittest.TestCase):
        # test across sheets, test move, test copy
        def test_same_cell(self):
                # wb = sheets.Workbook("wb")
                # sheet_num, sheet_name = wb.new_sheet(None)

                # wb.set_cell_contents(sheet_name, "A2", "=A1")
                # wb.set_cell_contents(sheet_name, "A3", "=$A1")
                # wb.set_cell_contents(sheet_name, "A1", "1")
                # self.assertEqual(wb.get_cell_value(sheet_name, "A2"), 
                                #    decimal.Decimal(1))
                # self.assertEqual(wb.get_cell_value(sheet_name, "A3"), 
                                #    decimal.Decimal(1))
                pass

        def test_relative_copies(self):
                # create 2+ cells, check that their formulas are modified
                pass
            
        def test_absolute_copies(self):
                # create 2+ cells, check that their formulas are not modified
                pass
            
        def test_relative_moves(self):
                pass
        
        def test_absolute_moves(self):
                pass

if __name__ == "__main__":
        unittest.main()