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
        self.value = value
        self.sheet.on_update([(self.sheet.sheet_name, self.location)])

    def get_value(self):
        return self.value

    def parse_formula(self):
        self.formula_tree = interp.parse_formula(self.contents)

        if self.formula_tree == None:
            raise FormulaError(CellError(CellErrorType.PARSE_ERROR, ""))

    def check_references(self, workbook):
        for ref in interp.find_refs(workbook, self.sheet, self.formula_tree):
            sheet_name = self.sheet.sheet_name.lower()
            location = ref
            if "!" in ref:
                index = ref.index("!")
                sheet_name = ref[:index].lower()
                location = ref[index+1:]

            workbook.sheet_references.link((self.sheet, self), sheet_name)

            try:
                cell = workbook.get_cell(sheet_name, location)
                workbook.dependency_graph.link(self, cell)
            except KeyError:
                print("keyerror")
                raise FormulaError(CellError(CellErrorType.BAD_REFERENCE, ""))
            except ValueError:
                raise FormulaError(CellError(CellErrorType.BAD_REFERENCE, ""))

    def check_cycles(self, workbook):
        for cycle in workbook.dependency_graph.get_cycles():
            if self in cycle:
                raise FormulaError(CellError(CellErrorType.CIRCULAR_REFERENCE, ""))

    def evaluate_formula(self, workbook):
        self.set_value(interp.evaluate_formula(workbook, self.sheet, self.formula_tree))
        
    def check_value(self):
        if self.value == None:
            self.set_value(decimal.Decimal(0))

        if type(self.value) == decimal.Decimal:
            self.set_value(remove_trailing_zeros(self.value))

    def update_referencing_nodes(self, workbook):
        ancestors = workbook.dependency_graph.get_ancestors_of_set(set([self]))
        ordered = workbook.dependency_graph.get_topological_order()
        for c in ordered:
            if c == self:
                continue
            if not c in ancestors:
                continue
            try:
                c.check_references(workbook)
                c.check_cycles(workbook)
                c.evaluate_formula(workbook)
            except FormulaError as e:
                c.set_value(e.value)

    def rename_sheet(self, old_name, new_name):
        self.contents = interp.rename_sheet(old_name, new_name, self.formula_tree)

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
                self.check_value()
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

        self.update_referencing_nodes(workbook)
    
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
