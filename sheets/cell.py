from .workbook import *
import enum
import decimal
from typing import *

from . import interp

def is_empty_content_string(contents):
    return contents == None or contents == "" or contents.isspace()

class FormulaError(Exception):
    def __init__(self, value):
        self.value = value

class Cell: 
    def __init__(self, sheet):
        self.sheet = sheet
        self.value = None
        self.contents = None
        self.formula_tree = None
    
    def __str__(self):
        return str(self.contents)
        
    def get_value(self, workbook, sheet):
        return self.value

    def parse_formula(self, workbook, sheet):
        self.formula_tree = interp.parse_formula(self.contents)

        if self.formula_tree == None:
            raise FormulaError(CellError(CellErrorType.PARSE_ERROR, ""))

    def check_references(self, workbook, sheet):
        for ref in interp.find_refs(workbook, sheet, self.formula_tree):
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
                raise FormulaError(CellError(CellErrorType.BAD_REFERENCE, ""))
            except ValueError:
                raise FormulaError(CellError(CellErrorType.BAD_REFERENCE, ""))

    def check_cycles(self, workbook, sheet):
        for cycle in workbook.graph.get_cycles():
            if self in cycle:
                raise FormulaError(CellError(CellErrorType.CIRCULAR_REFERENCE, ""))

    def evaluate_formula(self, workbook, sheet):
        self.value = interp.evaluate_formula(workbook, sheet, self.formula_tree)
        
    def check_contents(self, workbook, sheet):
        if self.value == None:
            self.value = decimal.Decimal(0)

    def update_referencing_nodes(self, workbook, sheet):
        ancestors = workbook.graph.get_ancestors_of_set(set([self]))
        ordered = workbook.graph.get_topological_order()
        for c in ordered:
            if c == self:
                continue
            if not c in ancestors:
                continue
            c.evaluate_formula(workbook, sheet)

    def set_contents(self, workbook, sheet, location: str, contents: str):
        workbook.sheet_references.clear_forward_links((sheet, self))
        workbook.graph.clear_forward_links(self)

        if is_empty_content_string(contents):
            self.contents = None
            self.value = None
            return

        self.contents = contents.strip()

        if contents[0] == "=":
            try:
                self.parse_formula(workbook, sheet)
                self.check_references(workbook, sheet)
                workbook.check_cycles()
                self.check_cycles(workbook, sheet)
                self.evaluate_formula(workbook, sheet)
                self.check_contents(workbook, sheet)
            except FormulaError as e:
                self.value = e.value
        elif contents[0] == "'":
            self.value = contents[1:]
        elif CellErrorType.from_string(contents) != None:
            self.value = CellError(CellErrorType.from_string(contents), "")
        else:
            try:
                num = contents
                if "." in num:
                    num = num.rstrip("0")
                self.value = decimal.Decimal(num)

                if not self.value.is_finite():
                    self.value = contents
            except decimal.InvalidOperation:
                self.value = contents

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
