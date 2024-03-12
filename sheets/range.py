import re

from .reference import Reference

class CellRange:

    range_regex = re.compile("((([A-Za-z_][A-Za-z0-9_]*|'[^']*')!)?\$?[A-Za-z]+\$?[0-9]+):((([A-Za-z_][A-Za-z0-9_]*|'[^']*')!)?\$?[A-Za-z]+\$?[0-9]+)")

    def __init__(self, sheet_name: str, start_location: str, end_location: str):
        self.sheet_name = sheet_name

        start_location_initial = Reference.from_string(sheet_name, start_location)
        end_location_initial = Reference.from_string(sheet_name, end_location)

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

    def from_string(range_string: str, default_sheet_name: str):
        m = CellRange.range_regex.fullmatch(range_string)

        if m is None:
            raise ValueError

        groups = m.groups()

        first_sheet_name = groups[2]
        second_sheet_name = groups[5]

        if first_sheet_name is not None and second_sheet_name is not None:
            if first_sheet_name != second_sheet_name:
                raise ValueError

        sheet_name = first_sheet_name or second_sheet_name or default_sheet_name

        start = groups[0]
        end = groups[3]

        return CellRange(sheet_name, start, end)
    
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
