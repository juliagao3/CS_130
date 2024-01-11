from typing import *
from .workbook import *
from .cell import Cell

class Sheet:
    def __init__(self, sheet_name):
        self.extent = (0, 0)
        self.sheet_name = sheet_name
        self.cells = {} # cell : location 
        # Workbook.list_sheets()
        # include workbook name ?
    
    def update_sheet_name(self, new_name):
        self.sheet_name = new_name
    
    def check_location(self, location: str):
        """
        check that the given location
        - has the form [row][col]
        - is to the left and above ZZZZ9999

        ValueError is raised if and only if the location is invalid
        """
        alpha_end = 0
        while alpha_end < len(location) and location[alpha_end].isalpha():
            alpha_end += 1

        if alpha_end == 0 or alpha_end == len(location):
            raise ValueError

        col = location[:alpha_end]
        row = location[alpha_end:]

        if col > "zzzz" or row > "9999":
            raise ValueError

    def set_cell_contents(self, location: str, content: str):
        """
        location - string like '[col][row]'
        """
        location = location.lower()
        self.check_location(location)
        if not location in self.cells:
            self.cells[location] = Cell(self.sheet_name)
        self.cells[location].set_contents(content)

    def get_cell_contents(self, location: str):
        """
        location - string like '[col][row]'
        """
        location = location.lower()
        self.check_location(location)
        if not location in self.cells:
            return None
        return self.cells[location].contents

    def get_cell_value(self, workbook, location: str):
        location = location.lower()
        self.check_location(location)
        try:
            return self.cells[location].get_value(workbook, self)
        except KeyError:
            # cell is not in the dict;
            # its value has not been set and it is empty
            return None
