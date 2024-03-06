import bisect

from . import base_types

from .cell import Cell
from .reference import Reference

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
            cell_contents[str(cell.location)] = str(cell)
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
            self.cells[location] = Cell(self, str(ref))
        
        return self.cells[location]
