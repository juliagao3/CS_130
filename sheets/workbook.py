from typing import *
import decimal
import re
from .sheet import Sheet
from .graph import Graph
from .cell import Cell, CellError, CellErrorType, FormulaError
from . import location as location_utils
import copy

from typing import TextIO
import json


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
        self.dependency_graph = Graph[Cell]()
        
        # map from names to copies
        self.num_copies: Dict[str, int] = {}

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

        if (sheet_name != None) and re.search(pattern, sheet_name) == None and (sheet_name != "") and (not sheet_name.isspace()):

            sheet_name = sheet_name.strip()

            pattern_2 = r'[^_A-Za-z0-9]'

            if not (sheet_name[0].isalpha() or sheet_name[0] == "") or ' ' in sheet_name or re.search(pattern_2, sheet_name) != None:
                sheet_name = "'" + sheet_name + "'"

            if sheet_name.lower() in self.sheet_map.keys():
                raise ValueError

        elif sheet_name == None:
            sheet_name = 'Sheet' + str(self.sheet_num)
            while sheet_name.lower() in self.sheet_map.keys():                
                self.sheet_num += 1
                sheet_name = 'Sheet' + str(self.sheet_num)  

        else:

            raise ValueError

        sheet = Sheet(sheet_name)
        self.sheet_map[sheet_name.lower()] = sheet
        self.sheets.append(sheet)

        working = set()
        if sheet_name.lower() in self.sheet_references.backward:
            for sheet, cell in self.sheet_references.backward[sheet_name.lower()]:
                try:
                    cell.check_references(self)
                    cell.evaluate_formula(self)
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
        
        if not sheet_name.lower() in self.sheet_references.backward:
            return

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
        # nature of the issue.
        
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
        solution = self.sheet_map[sheet_name.lower().strip()].get_cell_value(location)
        # if isinstance(solution, decimal.Decimal):
        return solution
    
    def get_cell(self, sheet_name: str, location:str):
        location = location_utils.check_location(location)
        return self.sheet_map[sheet_name.lower()].get_cell(location)

    def update_ancestors(self, nodes):
        order = self.dependency_graph.get_topological_order()
        ancestors = self.dependency_graph.get_ancestors_of_set(nodes)
        for cell in order:
            if cell in ancestors:
                cell.evaluate_formula(self)

    def check_cycles(self):
        cycles = self.dependency_graph.get_cycles()
        circular = set()
        for cycle in cycles:
            for cell in cycle:
                cell.value = CellError(CellErrorType.CIRCULAR_REFERENCE, "")
                circular.add(cell)
        self.update_ancestors(circular)

    @staticmethod
    def load_workbook(fp: TextIO) -> Any:
        # returns Workbook

        # This is a static method (not an instance method) to load a workbook
        # from a text file or file-like object in JSON format, and return the
        # new Workbook instance.  Note that the _caller_ of this function is
        # expected to have opened the file; this function merely reads the file.
        #
        # If the contents of the input cannot be parsed by the Python json
        # module then a json.JSONDecodeError should be raised by the method.
        # (Just let the json module's exceptions propagate through.)  Similarly,
        # if an IO read error occurs (unlikely but possible), let any raised
        # exception propagate through.
        #
        # If any expected value in the input JSON is missing (e.g. a sheet
        # object doesn't have the "cell-contents" key), raise a KeyError with
        # a suitably descriptive message.
        #
        # If any expected value in the input JSON is not of the proper type
        # (e.g. an object instead of a list, or a number instead of a string),
        # raise a TypeError with a suitably descriptive message.
        workbook_json = json.load(fp)

        wb = Workbook()

        try:
            sheets_list = workbook_json["sheets"]
            for sheet_dict in sheets_list:
                try:
                    num, name = wb.new_sheet(sheet_dict["name"])
                    cell_contents = sheet_dict["cell-contents"]
                    for location in cell_contents.keys():
                        if type(cell_contents[location]) == str:
                            wb.set_cell_contents(name, location, cell_contents[location])
                        else:
                            raise TypeError("Input JSON has an incorrect type: cell contents should be strings.")            
                except TypeError:
                    raise TypeError("Input JSON has an incorrect type: sheet name should be a string.")
        except KeyError:
            raise KeyError("Input JSON is missing an expected key: 'sheets', 'name', or 'cell-contents'.")

        return wb

    def save_workbook(self, fp: TextIO) -> None:
        # Instance method (not a static/class method) to save a workbook to a
        # text file or file-like object in JSON format.  Note that the _caller_
        # of this function is expected to have opened the file; this function
        # merely writes the file.
        #
        # If an IO write error occurs (unlikely but possible), let any raised
        # exception propagate through.
        # TODO: double quotes within cell formulas must be escaped??
        workbook_dict = dict()
        workbook_dict["sheets"] = list()

        for sheet in self.sheets:
            workbook_dict["sheets"].append(sheet.to_json())
              
        json.dump(workbook_dict, fp, indent=4)

    def notify_cells_changed(self,
            notify_function: Callable[[Any, Iterable[Tuple[str, str]]], None]) -> None:
        # notify_function (Workbook, Iterable...) -> None

        # Request that all changes to cell values in the workbook are reported
        # to the specified notify_function.  The values passed to the notify
        # function are the workbook, and an iterable of 2-tuples of strings,
        # of the form ([sheet name], [cell location]).  The notify_function is
        # expected not to return any value; any return-value will be ignored.
        #
        # Multiple notification functions may be registered on the workbook;
        # functions will be called in the order that they are registered.
        #
        # A given notification function may be registered more than once; it
        # will receive each notification as many times as it was registered.
        #
        # If the notify_function raises an exception while handling a
        # notification, this will not affect workbook calculation updates or
        # calls to other notification functions.
        #
        # A notification function is expected to not mutate the workbook or
        # iterable that it is passed to it.  If a notification function violates
        # this requirement, the behavior is undefined.
        pass

    def rename_sheet(self, sheet_name: str, new_sheet_name: str) -> None:
        # Rename the specified sheet to the new sheet name.  Additionally, all
        # cell formulas that referenced the original sheet name are updated to
        # reference the new sheet name (using the same case as the new sheet
        # name, and single-quotes iff [if and only if] necessary).
        #
        # The sheet_name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # As with new_sheet(), the case of the new_sheet_name is preserved by
        # the workbook.
        #
        # If the sheet_name is not found, a KeyError is raised.
        #
        # If the new_sheet_name is an empty string or is otherwise invalid, a
        # ValueError is raised.
        pass

    def move_sheet(self, sheet_name: str, index: int) -> None:
        # Move the specified sheet to the specified index in the workbook's
        # ordered sequence of sheets.  The index can range from 0 to
        # workbook.num_sheets() - 1.  The index is interpreted as if the
        # specified sheet were removed from the list of sheets, and then
        # re-inserted at the specified index.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        #
        # If the index is outside the valid range, an IndexError is raised.
        if index >= len(self.sheets) or index < 0:
            raise IndexError
        
        sheet_object = self.sheet_map[sheet_name.lower()]
        index_sheet = self.sheets.index(sheet_object)
        del self.sheets[index_sheet]
        
        self.sheets.insert(index, sheet_object)

    def copy_sheet(self, sheet_name: str) -> Tuple[int, str]:
        # Make a copy of the specified sheet, storing the copy at the end of the
        # workbook's sequence of sheets.  The copy's name is generated by
        # appending "_1", "_2", ... to the original sheet's name (preserving the
        # original sheet name's case), incrementing the number until a unique
        # name is found.  As usual, "uniqueness" is determined in a
        # case-insensitive manner.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # The copy should be added to the end of the sequence of sheets in the
        # workbook.  Like new_sheet(), this function returns a tuple with two
        # elements:  (0-based index of copy in workbook, copy sheet name).  This
        # allows the function to report the new sheet's name and index in the
        # sequence of sheets.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        
        new_name = ""   
        if sheet_name.lower() in self.sheet_map.keys():
            sheet_object = self.sheet_map[sheet_name.lower()]
            new_sheet = copy.deepcopy(sheet_object)
            
            if sheet_name.lower() in self.num_copies.keys():
                # make deep copy
                # add to num_copies and sheet_map and sheets      
                num = self.num_copies[sheet_name.lower()]
                num += 1
                new_name = sheet_name + '_' + str(num)
                new_sheet.update_sheet_name(new_name)
                self.num_copies[sheet_name.lower()] = num
            else:
                new_name = sheet_name + '_1'
                new_sheet.update_sheet_name(new_name)
                self.num_copies[sheet_name.lower()] = 1
        
            self.sheets.append(new_sheet)
            self.sheet_map[new_name.lower()] = new_sheet
            
        return (len(self.sheets) - 1, new_name)