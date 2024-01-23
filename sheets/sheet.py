from typing import *
from .workbook import *
from .cell import Cell
from . import location as location_utils

class Sheet:
    def __init__(self, sheet_name):
        self.extent = (0, 0)
        self.sheet_name = sheet_name
        self.cells = {}
    
    def to_json(self):
        return {
            "name": self.sheet_name,
            "cell-contents": self.cells
        }

    def update_sheet_name(self, new_name):
        self.sheet_name = new_name

    def set_cell_contents(self, workbook, location: str, content: str):
        """
        location - string like '[col][row]'
        """
        if not location in self.cells:
            self.cells[location] = Cell(self)
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
            self.cells[location] = Cell(self)
        
        return self.cells[location]
