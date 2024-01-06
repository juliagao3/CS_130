from workbook import *
from cell import *

class Sheet:
    def __init__(self, sheet_name):
        self.extent = (0, 0)
        self.sheet_name = sheet_name
        self.cells = {} # cell : location 
        # TODO if sheet_name is None, then use default sheetname
        # Workbook.list_sheets()
        # include workbook name ?
    
    def update_sheet_name(self, new_name):
        pass

    def add_cell_to_sheet(self, cell : Cell):

        self.update_extent()
        pass

    def update_extent(self):
        pass