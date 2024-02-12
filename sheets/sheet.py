from typing import *
import re
from .workbook import *
from .cell import Cell
from . import location as location_utils

import bisect

def name_is_valid(name: str):
    if name is None or name == "" or name.isspace():
        return False

    disallowed = r'[^A-Za-z0-9.?!,:;@#$%^&*()-_\s+]'

    if re.search(disallowed, name) is not None:
        return False

    return True

def name_needs_quotes(name: str):
    assert len(name) > 0

    if name[0] != '_' and not name[0].isalpha():
        return True

    for c in name:
        if c != '_' and not c.isalpha() and not c.isdigit():
            return True

    return False

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
            cell_contents[location] = str(self.cells[location])
        json_obj["cell-contents"] = cell_contents
        return json_obj

    def get_quoted_name(self):
        if name_needs_quotes(self.sheet_name):
            return "'" + self.sheet_name + "'"
        return self.sheet_name

    def on_update(self, locations):
        self.workbook.on_update(locations)

    def update_sheet_name(self, new_name):
        self.sheet_name = new_name

    def set_cell_contents(self, workbook, location: str, content: str):
        """
        location - string like '[col][row]'
        """
        if location not in self.cells:
            self.cells[location] = Cell(self, location)

        old_contents = self.cells[location].contents
        self.cells[location].set_contents(workbook, content)

        empty = content is None or content == "" or content.isspace()
        was_empty = old_contents is None or old_contents == "" or old_contents.isspace()
        
        if empty and not was_empty:
            location_num = location_utils.location_string_to_tuple(location)
            index_col = bisect.bisect_left(self.cols_hist, location_num[0])
            index_row = bisect.bisect_left(self.rows_hist, location_num[1])
            self.cols_hist.pop(index_col)
            self.rows_hist.pop(index_row)

            assert len(self.cols_hist) == len(self.rows_hist)

            if len(self.cols_hist) == 0:
                self.extent = (0, 0)
            else:
                self.extent = (self.cols_hist[len(self.cols_hist) - 1], self.rows_hist[len(self.rows_hist) - 1])
        elif not empty:
            location_num = location_utils.location_string_to_tuple(location)
            self.extent = (max(self.extent[0], location_num[0]),
                            max(self.extent[1], location_num[1]))
            bisect.insort(self.cols_hist, location_num[0])
            bisect.insort(self.rows_hist, location_num[1])

        return self.cells[location]
    

    def get_cell_contents(self, location: str):
        """
        location - string like '[col][row]'
        """
        if location not in self.cells:
            return None
        return self.cells[location].contents

    def get_cell_value(self, location: str):
        try:
            return self.cells[location].get_value()
        except KeyError:
            # cell is not in the dict;
            # its value has not been set and it is empty
            return None
    
    def get_cell(self, location: str):
        if location not in self.cells:
            self.cells[location] = Cell(self, location)
        
        return self.cells[location]
