from .workbook import *
import enum
import decimal
from typing import *

from . import interp

def is_empty_content_string(contents):
    return contents == None or contents == "" or contents.isspace()

def remove_trailing_zeros(d: decimal.Decimal):
    num = str(d)
    if "." in num:
        num = num.rstrip("0")
    if num[-1] == ".":
        num = num[:-1]
    return decimal.Decimal(num)

def notify(workbook, cells):
    for func in workbook.notify_functions:
        try:
            func(workbook, map(lambda c: (c.sheet.sheet_name, c.location), cells))
        except:
            pass

class FormulaError(Exception):
    def __init__(self, value):
        self.value = value

class Cell: 
    def __init__(self, sheet, location):
        self.sheet = sheet
        self.location = location
        self.value = None
        self.contents = None
        self.formula_tree = None
    
    def __str__(self):
        return str(self.contents)
        
    def set_value(self, value):
        old_value = self.value
        self.value = value
        if isinstance(old_value, CellError) and isinstance(value, CellError):
            old_value = str(old_value)
            value = str(value)            
        if value != old_value:
            notify(self.sheet.workbook, {self})     

    def get_value(self):
        return self.value

    def parse_formula(self):
        self.formula_tree = interp.parse_formula(self.contents)

        if self.formula_tree == None:
            raise FormulaError(CellError(CellErrorType.PARSE_ERROR, ""))

    def check_references(self, workbook):
        is_error = False

        for ref in interp.find_refs(workbook, self.sheet, self.formula_tree):
            sheet_name = self.sheet.sheet_name.lower()
            location = ref
            if "!" in ref:
                index = ref.rfind("!")
                sheet_name = ref[:index].lower()
                location = ref[index+1:]
            workbook.sheet_references.link(self, sheet_name)

            try:
                cell = workbook.get_cell(sheet_name, location)
                workbook.dependency_graph.link(self, cell)
            except (KeyError, ValueError):
                is_error = True
                pass
                
        if is_error:
            raise FormulaError(CellError(CellErrorType.BAD_REFERENCE, ""))

    def check_cycles(self, workbook):
        for cycle in workbook.dependency_graph.get_cycles():
            if self in cycle:
                raise FormulaError(CellError(CellErrorType.CIRCULAR_REFERENCE, ""))

    
    def evaluate_formula(self, workbook):
        value = interp.evaluate_formula(workbook, self.sheet, self.formula_tree)

        if value == None:
            value = decimal.Decimal(0)

        if type(value) == decimal.Decimal:
            value = remove_trailing_zeros(value)

        self.set_value(value)

    def rename_sheet(self, old_name, new_name):
        self.contents = interp.rename_sheet(old_name, new_name, self.formula_tree)

    def recompute_value(self, workbook):
        if self.contents == None or self.contents[0] != "=":
            return
        try:
            self.check_references(workbook)
            self.check_cycles(workbook)
            self.evaluate_formula(workbook)
        except FormulaError as e:
            self.set_value(e.value)

    def set_contents(self, workbook, contents: str):
        workbook.sheet_references.clear_forward_links((self.sheet, self))
        workbook.dependency_graph.clear_forward_links(self)

        if is_empty_content_string(contents):
            self.contents = None
            self.set_value(None)
            return

        contents = contents.strip()
        self.contents = contents

        if contents[0] == "=":
            try:
                self.parse_formula()
                self.check_references(workbook)
                workbook.check_cycles()
                self.check_cycles(workbook)
                self.evaluate_formula(workbook)
            except FormulaError as e:
                self.set_value(e.value)
        elif contents[0] == "'":
            self.set_value(contents[1:])
        elif CellErrorType.from_string(contents) != None:
            self.set_value(CellError(CellErrorType.from_string(contents), ""))
        else:
            try:
                self.set_value(remove_trailing_zeros(decimal.Decimal(contents)))
                if not self.value.is_finite():
                    self.set_value(contents)
            except decimal.InvalidOperation:
                self.set_value(contents)
    
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
