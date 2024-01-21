from typing import *
import decimal
import re
from .sheet import Sheet
from .graph import Graph
from .cell import Cell, CellError, CellErrorType
from . import location as location_utils

class Workbook:
    # A workbook containing zero or more named spreadsheets.
    #
    # Any and all operations on a workbook that may affect calculated cell
    # values should cause the workbook's contents to be updated properly.

    def __init__(self, workbook_name: str=None):
        # Initialize a new empty workbook.
        if workbook_name != None:
            self.workbook_name: str = workbook_name
        else:
            self.workbook_name: str = "wb"

        # List of sheets in order of creation
        self.sheets: List[Sheet] = []

        # map from lowercase sheet name to sheet
        self.sheet_map: Dict[str, Sheet] = {}
        
        # cells point to sheet names they reference
        self.sheet_references = Graph[Any]()

        # Default sheet number
        self.sheet_num: int = 1

        # Graph
        self.graph = Graph[Cell]()

    def num_sheets(self) -> int:
        # Return the number of spreadsheets in the workbook.
        return len(self.sheets)

    def list_sheets(self) -> List[str]:
        # Return a list of the spreadsheet names in the workbook, with the
        # capitalization specified at creation, and in the order that the sheets
        # appear within the workbook.
        #
        # In this project, the sheet names appear in the order that the user
        # created them; later, when the user is able to move and copy sheets,
        # the ordering of the sheets in this function's result will also reflect
        # such operations.
        #
        # A user should be able to mutate the return-value without affecting the
        # workbook's internal state.
        return [sheet.sheet_name for sheet in self.sheets]

    def new_sheet(self, sheet_name: Optional[str] = None) -> Tuple[int, str]:
        # Add a new sheet to the workbook.  If the sheet name is specified, it
        # must be unique.  If the sheet name is None, a unique sheet name is
        # generated.  "Uniqueness" is determined in a case-insensitive manner,
        # but the case specified for the sheet name is preserved.
        #
        # The function returns a tuple with two elements:
        # (0-based index of sheet in workbook, sheet name).  This allows the
        # function to report the sheet's name when it is auto-generated.
        #
        # If the spreadsheet name is an empty string (not None), or it is
        # otherwise invalid, a ValueError is raised.

        pattern = r'[^A-Za-z0-9.?!,:;@#$%^&*()-_\s+]'

        if (sheet_name != None) and re.search(pattern, sheet_name) == None and (sheet_name != ""):

            sheet_name = sheet_name.strip()

            pattern_2 = r'[^_A-Za-z0-9]'

            if not (sheet_name[0].isalpha() or sheet_name[0] == "") or ' ' in sheet_name or re.search(pattern_2, sheet_name) != None:
                sheet_name = "'" + sheet_name + "'"

            if sheet_name.lower() in self.sheet_map.keys():
                raise ValueError

        elif sheet_name == None:

            sheet_name = 'Sheet' + str(self.sheet_num)
            self.sheet_num += 1

        else:

            raise ValueError

        sheet = Sheet(sheet_name)
        self.sheet_map[sheet_name.lower()] = sheet
        self.sheets.append(sheet)

        working = set()
        if sheet_name.lower() in self.sheet_references.backward:
            for sheet, cell in self.sheet_references.backward[sheet_name.lower()]:
                try:
                    cell.check_references(self, sheet)
                    cell.evaluate_formula(self, sheet)
                except FormulaError as e:
                    cell.value = e.value

                # this step is probly slow... it would be better to update
                # every affected cell outside this loop
                working.add(cell)

        self.update_ancestors(working)

        return (len(self.sheets)-1, sheet_name)

    def del_sheet(self, sheet_name: str) -> None:
        # Delete the spreadsheet with the specified name.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        sheet = self.sheet_map.pop(sheet_name.lower())
        self.sheets.remove(sheet)
        
        # nodes that reference this cell will have their value recomputed
        # and find that they now have a bad reference since the sheet has
        # been removed from the workbook
        broken = set()
        for sheet, cell in self.sheet_references.backward[sheet_name.lower()]:
            cell.value = CellError(CellErrorType.BAD_REFERENCE, "")
            broken.add(cell)
        self.update_ancestors(broken)

    def get_sheet_extent(self, sheet_name: str) -> Tuple[int, int]:
        # Return a tuple (num-cols, num-rows) indicating the current extent of
        # the specified spreadsheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        
        return self.sheet_map[sheet_name.lower()].extent

    def set_cell_contents(self, sheet_name: str, location: str,
                          contents: Optional[str]) -> None:
        # Set the contents of the specified cell on the specified sheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.  Additionally, the cell location can be
        # specified in any case.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If the cell location is invalid, a ValueError is raised.
        #
        # A cell may be set to "empty" by specifying a contents of None.
        #
        # Leading and trailing whitespace are removed from the contents before
        # storing them in the cell.  Storing a zero-length string "" (or a
        # string composed entirely of whitespace) is equivalent to setting the
        # cell contents to None.
        #
        # If the cell contents appear to be a formula, and the formula is
        # invalid for some reason, this method does not raise an exception;
        # rather, the cell's value will be a CellError object indicating the
        # naure of the issue.
        
        location = location_utils.check_location(location)
        self.sheet_map[sheet_name.lower()].set_cell_contents(self, location, contents)

    def get_cell_contents(self, sheet_name: str, location: str) -> Optional[str]:
        # Return the contents of the specified cell on the specified sheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.  Additionally, the cell location can be
        # specified in any case.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If the cell location is invalid, a ValueError is raised.
        #
        # Any string returned by this function will not have leading or trailing
        # whitespace, as this whitespace will have been stripped off by the
        # set_cell_contents() function.
        #
        # This method will never return a zero-length string; instead, empty
        # cells are indicated by a value of None.

        location = location_utils.check_location(location)
        return self.sheet_map[sheet_name.lower()].get_cell_contents(location)

    def get_cell_value(self, sheet_name: str, location: str) -> Any:
        # Return the evaluated value of the specified cell on the specified
        # sheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.  Additionally, the cell location can be
        # specified in any case.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If the cell location is invalid, a ValueError is raised.
        #
        # The value of empty cells is None.  Non-empty cells may contain a
        # value of str, decimal.Decimal, or CellError.
        #
        # Decimal values will not have trailing zeros to the right of any
        # decimal place, and will not include a decimal place if the value is a
        # whole number.  For example, this function would not return
        # Decimal('1.000'); rather it would return Decimal('1').

        location = location_utils.check_location(location)
        solution = self.sheet_map[sheet_name.lower().strip()].get_cell_value(self, location)
        if isinstance(solution, decimal.Decimal):
            solution = str(solution)
            if "." in solution:
                solution = solution.rstrip("0")
            if solution[-1] == ".":
                solution = solution[:-1]
            solution = decimal.Decimal(solution)
        return solution
    
    def get_cell(self, sheet_name: str, location:str):
        location = location_utils.check_location(location)
        return self.sheet_map[sheet_name.lower()].get_cell(location)

    def update_ancestors(self, nodes):
        order = self.graph.get_topological_order()
        ancestors = self.graph.get_ancestors_of_set(nodes)
        for cell in order:
            if cell in ancestors:
                cell.evaluate_formula(self, cell.sheet)

    def check_cycles(self):
        cycles = self.graph.get_cycles()
        circular = set()
        for cycle in cycles:
            for cell in cycle:
                cell.value = CellError(CellErrorType.CIRCULAR_REFERENCE, "")
                circular.add(cell)
        self.update_ancestors(circular)
