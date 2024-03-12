import re

from typing import Optional

from .reference import Reference

class CellRange:

    range_regex = re.compile("((([A-Za-z_][A-Za-z0-9_]*|'[^']*')!)?\$?[A-Za-z]+\$?[0-9]+):((([A-Za-z_][A-Za-z0-9_]*|'[^']*')!)?\$?[A-Za-z]+\$?[0-9]+)")

    def __init__(self, sheet_name: str, start_location: str, end_location: str, check_bounds: bool = True):
        self.sheet_name = sheet_name

        start_location_initial = Reference.from_string(start_location, allow_absolute=True, check_bounds=check_bounds)
        end_location_initial = Reference.from_string(end_location, allow_absolute=True, check_bounds=check_bounds)
        
        self.start_ref = Reference(min(start_location_initial.col, end_location_initial.col), min(start_location_initial.row, end_location_initial.row), check_bounds=check_bounds)
        self.end_ref = Reference(max(start_location_initial.col, end_location_initial.col), max(start_location_initial.row, end_location_initial.row), check_bounds=check_bounds)

    def from_string(range_string: str, default_sheet_name: Optional[str] = None, check_bounds: bool = True):
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

        return CellRange(sheet_name, start, end, check_bounds)
    
    def generate(self):
        for row in range(self.start_ref.row, self.end_ref.row + 1):
            for col in range(self.start_ref.col, self.end_ref.col + 1):
                yield Reference(col, row, sheet_name=self.sheet_name)
    
    def generate_values(self, workbook):
        for ref in self.generate():
            yield workbook.get_cell(ref.sheet_name, ref).value