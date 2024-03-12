import functools

from . import base_types

from .cell import Cell
from .reference import Reference

from typing import List

@functools.total_ordering
class SortRow:

    def __init__(self, sheet, sort_cols, order, row_index):
        self.sheet = sheet
        self.sort_cols = sort_cols
        self.order = order
        self.row_index = row_index
    
    def __lt__(self, other):
        for col_order, col_index in zip(self.order, self.sort_cols):
            my_ref = Reference(self.sheet.sheet_name, col_index, self.row_index)
            my_value = self.sheet.get_cell_value(my_ref)

            other_ref = Reference(self.sheet.sheet_name, col_index, other.row_index)
            other_value = self.sheet.get_cell_value(other_ref)

            if base_types.lt(my_value, other_value):
                return True ^ col_order
            elif base_types.lt(other_value, my_value):
                return False ^ col_order
        return False
    
    def __str__(self):
        return str(self.row_index)

class Sheet:
    def __init__(self, workbook, sheet_name):
        self.workbook = workbook
        self.extent = (0, 0)
        self.sheet_name = sheet_name
        self.cells = {}
        self.cols_hist = []
        self.rows_hist = []
        
    def to_json(self):
        json_obj = {
            "name": self.sheet_name
        }
        locations = self.cells.keys()
        cell_contents = dict()
        for location in locations:
            if self.cells[location].contents is None:
                continue
            cell = self.cells[location]
            cell_contents[cell.location.location_string()] = str(cell)
        json_obj["cell-contents"] = cell_contents
        return json_obj

    def get_quoted_name(self):
        if base_types.sheet_name_needs_quotes(self.sheet_name):
            return "'" + self.sheet_name + "'"
        return self.sheet_name

    def on_update(self, locations):
        self.workbook.on_update(locations)

    def update_sheet_name(self, new_name):
        self.sheet_name = new_name

    def get_extent(self):
        extent = (0, 0)
        for location, cell in self.cells.items():
            if cell.contents is None:
                continue
            extent = (max(extent[0], location[0]),
                        max(extent[1], location[1]))
        return extent

    def set_cell_contents(self, workbook, ref: Reference, content: str):
        location = ref.tuple()

        if location not in self.cells:
            self.cells[location] = Cell(self, ref)

        self.cells[location].set_contents(workbook, content)

        return self.cells[location]

    def get_cell_contents(self, ref: Reference):
        location = ref.tuple()
        if location not in self.cells:
            return None
        return self.cells[location].contents

    def get_cell_value(self, ref: Reference):
        location = ref.tuple()
        try:
            return self.cells[location].get_value()
        except KeyError:
            # cell is not in the dict;
            # its value has not been set and it is empty
            return None
    
    def get_cell(self, ref: Reference):
        location = ref.tuple()

        if location not in self.cells:
            self.cells[location] = Cell(self, ref)
        
        return self.cells[location]

    def sort_region(self, workbook, start_ref, end_ref, sort_cols: List[int]):
        order = [False if col > 0 else True for col in sort_cols]
        sort_cols = [start_ref.col + abs(col) - 1 for col in sort_cols]

        sort_rows = [SortRow(self, sort_cols, order, row) for row in range(start_ref.row, end_ref.row + 1)]
        sort_rows.sort()

        copy = []

        for col in range(start_ref.col, end_ref.col + 1):
            row_list = []
            for row in range(start_ref.row, end_ref.row + 1):
                ref = Reference(self.sheet_name, col, row)
                row_list.append(self.get_cell(ref))
            copy.append(row_list)

        for to_row_relative, sort_row in enumerate(sort_rows):
            to_row = to_row_relative + start_ref.row
            from_row = sort_row.row_index
            for col in range(start_ref.col, end_ref.col + 1):
                c = copy[col - start_ref.col][from_row - start_ref.row]
                c.location = Reference(self.sheet_name, col, to_row)
                c.move_formula(workbook, (0, to_row - from_row))
                self.cells[(col, to_row)] = copy[col - start_ref.col][from_row - start_ref.row]

