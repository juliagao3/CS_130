from typing import *
import re
from .workbook import *
from .cell import Cell
from . import location as location_utils

def name_is_valid(name: str):
    if name == None or name == "" or name.isspace():
        return False

    disallowed = r'[^A-Za-z0-9.?!,:;@#$%^&*()-_\s+]'

    if re.search(disallowed, name) != None:
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
    
    def to_json(self):
        json_obj = {
            "name": self.sheet_name
        }
        locations = self.cells.keys()
        cell_contents = dict()
        for location in locations:
            if self.cells[location].contents == None:
                continue
            cell_contents[location] = str(self.cells[location])
        json_obj["cell-contents"] = cell_contents
        return json_obj

    def on_update(self, locations):
        self.workbook.on_update(locations)

    def update_sheet_name(self, new_name):
        self.sheet_name = new_name

    def set_cell_contents(self, workbook, location: str, content: str):
        """
        location - string like '[col][row]'
        """
        if not location in self.cells:
            self.cells[location] = Cell(self, location)
        self.cells[location].set_contents(workbook, content)
        
        if content == None or content == "" or content.isspace():
            self.extent = (0, 0)
            for location in self.cells.keys():
                cell_content = self.cells[location].contents
                if cell_content == None or cell_content == "" or cell_content.isspace():
                    continue
                location_num = location_utils.location_string_to_tuple(location)
                self.extent = (max(self.extent[0], location_num[0]),
                            max(self.extent[1], location_num[1]))
        else:
            location_num = location_utils.location_string_to_tuple(location)
            self.extent = (max(self.extent[0], location_num[0]),
                            max(self.extent[1], location_num[1]))

        return self.cells[location]
    

    def get_cell_contents(self, location: str):
        """
        location - string like '[col][row]'
        """
        if not location in self.cells:
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
        if not location in self.cells:
            self.cells[location] = Cell(self, location)
        
        return self.cells[location]