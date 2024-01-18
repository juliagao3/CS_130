from .workbook import *
import enum
import decimal
from typing import *

from . import interp

def is_empty_content_string(contents):
    return contents == None or contents == "" or contents.isspace()

class Cell: 
    def __init__(self):
        self.value = None
        self.contents = None
    
    def __str__(self):
        if self.contents == None:
            return "{None}"
        else:
            return self.contents
        
    def get_value(self, workbook, sheet):
        # if not None, we've already evaluated the contents
        if self.value == None and self.contents != None:
            # if the first char is =, evaluate as a formula
            if self.contents[0] == "=":
                # call evaluate formula function? also set self.value?
                # evaluate formula will return a value or raise an error
                self.value = interp.evaluate_formula(workbook, sheet, interp.parse_formula(self.contents))
            elif CellErrorType.from_string(self.contents) != None:
                self.value = CellError(CellErrorType.from_string(self.contents), "")
            elif self.contents[0] == "'":
                self.value = self.contents[1:]
            else:
                # evaluate as a literal (number, date, etc.)
                # for now, just numbers and strings
                # TODO: other values (NaN, Infinity, -Infinity) should not be converted to nums
                # can us decimal.is_finite() method to check
                # remove trailing zeros from number (str)
                if "." in self.contents:
                    num = self.contents.rstrip("0")
                else:
                    num = self.contents
                if num[-1] == ".":
                    num = num[:-1]
                try:
                    # store value as a Decimal
                    self.value = decimal.Decimal(num)
                except decimal.InvalidOperation:
                    # if it failed to parse store the contents as a string
                    self.value = self.contents
                if num == "NaN" or num == "Infinity" or num == "-Infinity" or num == "Inf" or num == "-Inf":
                    self.value = self.contents

        # evaluate formulas in workbook
        return self.value

    def recompute_value(self, workbook, sheet):
        if type(self.value) != CellError or self.value.get_type() != CellErrorType.CIRCULAR_REFERENCE:
            self.value = None
            self.get_value(workbook, sheet)
    
    def update_referencing_nodes(self, workbook, sheet):
        ancestors = workbook.graph.get_ancestors(self)
        ordered = ancestors.topological_sort()
        for c in ordered:
            c.recompute_value(workbook, sheet)

    def set_contents(self, workbook, sheet, location, contents: str):
        # if contents == "" or only whitespace, contents + value are still None
        if is_empty_content_string(contents):
            self.contents = None
            workbook.graph.clear_forward_links(self)
            self.update_referencing_nodes(workbook, sheet)
            return

        self.contents = contents.strip()
        self.value = None

        workbook.sheet_references.clear_forward_links((sheet, self))
        workbook.graph.clear_forward_links(self)
        # if it's a formula, scan for references
        if contents[0] == "=":
            tree = interp.parse_formula(self.contents)
            if tree == None:
                self.value = CellError(CellErrorType.PARSE_ERROR, "")
            else:
                for ref in interp.find_refs(workbook, sheet, tree):
                    sheet_name = sheet.sheet_name.lower()
                    location = ref
                    if "!" in ref:
                        index = ref.index("!")
                        sheet_name = ref[:index].lower()
                        location = ref[index+1:]

                    workbook.sheet_references.link((sheet, self), sheet_name)

                    try:
                        cell = workbook.get_cell(sheet_name, location)
                        workbook.graph.link(self, cell)
                    except KeyError:
                        self.value = CellError(CellErrorType.BAD_REFERENCE, "")
                        return
                    except ValueError:
                        self.value = CellError(CellErrorType.BAD_REFERENCE, "")
                        return

                cycle = workbook.graph.find_cycle(self)
                if cycle != None:
                    for cell in cycle:
                        cell.value = CellError(CellErrorType.CIRCULAR_REFERENCE, "")

        self.update_referencing_nodes(workbook, sheet)
    
class CellErrorType(enum.Enum):
    '''
    This enum specifies the kinds of errors that spreadsheet cells can hold.
    '''

    def from_string(s):
        errors = {
                "#ERROR!": CellErrorType.PARSE_ERROR,
                "#CIRCREF!": CellErrorType.CIRCULAR_REFERENCE,
                "#REF!": CellErrorType.BAD_REFERENCE,
                "#NAME?": CellErrorType.BAD_NAME,
                "#VALUE!": CellErrorType.TYPE_ERROR,
                "#DIV/0!": CellErrorType.DIVIDE_BY_ZERO
        }
        s = s.upper()
        if s in errors:
            return errors[s]
        return None
    
    # A formula doesn't parse successfully ("#ERROR!")
    PARSE_ERROR = 1

    # A cell is part of a circular reference ("#CIRCREF!")
    CIRCULAR_REFERENCE = 2

    # A cell-reference is invalid in some way ("#REF!")
    BAD_REFERENCE = 3

    # Unrecognized function name ("#NAME?")
    BAD_NAME = 4

    # A value of the wrong type was encountered during evaluation ("#VALUE!")
    TYPE_ERROR = 5

    # A divide-by-zero was encountered during evaluation ("#DIV/0!")
    DIVIDE_BY_ZERO = 6

class CellError:
    '''
    This class represents an error value from user input, cell parsing, or
    evaluation.
    '''

    def __init__(self, error_type: CellErrorType, detail: str,
                 exception: Optional[Exception] = None):
        self._error_type = error_type
        self._detail = detail
        self._exception = exception

    def get_type(self) -> CellErrorType:
        ''' The category of the cell error. '''
        return self._error_type

    def get_detail(self) -> str:
        ''' More detail about the cell error. '''
        return self._detail

    def get_exception(self) -> Optional[Exception]:
        '''
        If the cell error was generated from a raised exception, this is the
        exception that was raised.  Otherwise this will be None.
        '''
        return self._exception

    def __str__(self) -> str:
        return f'ERROR[{self._error_type}, "{self._detail}"]'

    def __repr__(self) -> str:
        return self.__str__()