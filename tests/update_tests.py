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

    def test_update(self):
        updated = set()

        def on_cells_changed(workbook, changed_cells):
            nonlocal updated
            for cell in changed_cells:
                updated.add(cell)
            raise Exception

        # Make a workbook, and register our notification function on it.
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.notify_cells_changed(on_cells_changed)

        # Generates one call to notify functions, with the argument [('Sheet1', 'A1')].
        wb.set_cell_contents(n, "A1", "'123")

        self.assertEqual(updated, set([(n, "a1")]))
        updated.clear()

        # Generates one call to notify functions, with the argument [('Sheet1', 'C1')].
        wb.set_cell_contents(n, "C1", "=A1+B1")

        self.assertEqual(updated, set([(n, "c1")]))
        updated.clear()

        # Generates one or more calls to notify functions, indicating that cells B1
        # and C1 have changed.  For example, there might be one call with the argument
        # [('Sheet1', 'B1'), ('Sheet1', 'C1')].
        wb.set_cell_contents(n, "B1", "5.3")

        self.assertEqual(updated, set([(n, "b1"), (n, "c1")]))
        updated.clear()

    def test_exception(self):
        def on_cells_changed(workbook, changed_cells):
            raise Exception

        # Make a workbook, and register our notification function on it.
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        wb.notify_cells_changed(on_cells_changed)

        # Generates one call to notify functions, with the argument [('Sheet1', 'A1')].
        wb.set_cell_contents(n, "A1", "'123")

    def test_value_updated_first(self):
        expected = [decimal.Decimal(1), decimal.Decimal(2)]
        actual = []

        # Make a workbook, and register our notification function on it.
        wb = sheets.Workbook()
        i, n = wb.new_sheet()

        def on_cells_changed(workbook, changed_cells):
            nonlocal wb
            sheet_name, location = next(changed_cells)
            value = wb.get_cell_value(sheet_name, location)
            actual.append(value)

        wb.notify_cells_changed(on_cells_changed)

        # Generates one call to notify functions, with the argument [('Sheet1', 'A1')].
        wb.set_cell_contents(n, "A1", "1")
        wb.set_cell_contents(n, "A1", "2")

        self.assertEqual(actual, expected)

    def test_order(self):

        order = []

        def make_on_cell_changed(i):
            nonlocal order
            def on_cell_changed(workbook, location):
                nonlocal i
                nonlocal order
                order.append(i)
            return on_cell_changed
        
        wb = sheets.Workbook()
        _i, n = wb.new_sheet()

        k = 10
        for i in range(k):
            wb.notify_cells_changed(make_on_cell_changed(i))

        wb.set_cell_contents(n, "A1", "1")

        for i in range(0, len(order), k):
            sub_order = order[i:i+k]
            self.assertEqual(sub_order, sorted(sub_order))
    
    def test_delete(self):
        updated = set()

        def on_cells_changed(workbook, changed_cells):
            nonlocal updated
            for cell in changed_cells:
                updated.add(cell)
            raise Exception

        # Make a workbook, and register our notification function on it.
        wb = sheets.Workbook()
        i, n = wb.new_sheet()
        j, m = wb.new_sheet()

        wb.notify_cells_changed(on_cells_changed)

        # Generates one call to notify functions, with the argument [('Sheet1', 'A1')].
        wb.set_cell_contents(n, "A1", f"={m}!A1")

        self.assertEqual(updated, set([(n, "a1")]))
        updated.clear()

        wb.del_sheet(m)

        self.assertEqual(updated, set([(n, "a1")]))

if __name__ == "__main__":
        unittest.main()
