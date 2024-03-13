import re

from .reference import Reference

class CellRange:

    range_regex = re.compile("((([A-Za-z_][A-Za-z0-9_]*|'[^']*')!)?\$?[A-Za-z]+\$?[0-9]+):((([A-Za-z_][A-Za-z0-9_]*|'[^']*')!)?\$?[A-Za-z]+\$?[0-9]+)")

    def __init__(self, default_sheet_name: str, start_location: str, end_location: str):
        # The semantics around providing sheet names for each of the endpoints
        # of a range are a bit special. We pass None here to compute the sheet
        # name later...
        start_location_initial = Reference.from_string(None, start_location)
        end_location_initial = Reference.from_string(None, end_location)

        # We take the sheet name from the start ref, falling back on the end
        # ref's sheet if the start ref's was None, and finally falling back on
        # the provided default if both were None.
        self.sheet_name = \
                start_location_initial.sheet_name \
                or end_location_initial.sheet_name \
                or default_sheet_name

        # Refs must have a sheet name...
        if start_location_initial.sheet_name is None:
            start_location_initial.sheet_name = self.sheet_name

        if end_location_initial.sheet_name is None:
            end_location_initial.sheet_name = self.sheet_name

        # If the sheet names don't match, then we have a problem.
        if start_location_initial.sheet_name != end_location_initial.sheet_name:
            raise ValueError

        # Reorder the refs so start_ref is the upper left and end_ref is the
        # lower right
        self.start_ref = Reference.min(start_location_initial, end_location_initial)
        self.end_ref = Reference.max(start_location_initial, end_location_initial)

    def check_bounds(self):
        self.start_ref.check_bounds()
        self.end_ref.check_bounds()
        return self

    def check_absolute(self):
        self.start_ref.check_absolute()
        self.end_ref.check_absolute()
        return self

    def from_string(default_sheet_name: str, range_string: str):
        m = CellRange.range_regex.fullmatch(range_string)

        if m is None:
            raise ValueError

        groups = m.groups()

        start = groups[0]
        end = groups[3]

        return CellRange(default_sheet_name, start, end)
    
    def generate_column(self, col: int):
        for row in range(self.start_ref.row, self.end_ref.row + 1):
            yield Reference(self.sheet_name, int(col) + self.start_ref.col, row)

    def generate_row(self, row: int):
        for col in range(self.start_ref.col, self.end_ref.col + 1):
            yield Reference(self.sheet_name, col, int(row) + self.start_ref.row)

    def generate(self):
        for row in range(self.start_ref.row, self.end_ref.row + 1):
            for col in range(self.start_ref.col, self.end_ref.col + 1):
                yield Reference(self.sheet_name, col, row)

    def generate_column_values(self, col, workbook):
        for ref in self.generate_column(col):
            yield workbook.get_cell(ref).value

    def generate_row_values(self, row, workbook):
        for ref in self.generate_row(row):
            yield workbook.get_cell(ref).value

    def generate_values(self, workbook):
        for ref in self.generate():
            yield workbook.get_cell(ref).value
