from .reference import Reference

class CellRange:

    def __init__(self, sheet_name: str, start_location: str, end_location: str):
        self.sheet_name = sheet_name

        start_location_initial = Reference.from_string(start_location)
        end_location_initial = Reference.from_string(end_location)
        
        self.start_ref = Reference(min(start_location_initial.col, end_location_initial.col), min(start_location_initial.row, end_location_initial.row))
        self.end_ref = Reference(max(start_location_initial.col, end_location_initial.col), max(start_location_initial.row, end_location_initial.row))

    
    def generate(self):
        for row in range(self.start_ref.row, self.end_ref.row + 1):
            for col in range(self.start_ref.col, self.end_ref.col + 1):
                yield Reference(col, row, sheet_name = self.sheet_name)
    
    def generate_values(self, workbook):
        for ref in self.generate():
            yield workbook.get_cell(ref.sheet_name, ref).value