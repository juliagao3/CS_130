from .workbook import *
import enum
import decimal
from typing import *

class Cell: 
    def __init__(self, sheet_name):
        self.sheet_name = sheet_name
        self.value = None
        self.contents = None

        # cells referenced by this cell's formula
        self.dependencies = [] 

        # cells that reference this cell
        self.referenced_by = []
    
    def get_value(self, workbook_name, sheet_name):
        # if not None, we've already evaluated the contents
        if self.value == None and self.contents != None:
            # if the first char is =, evaluate as a formula
            if self.contents[0] == "=":
                # call evaluate formula function? also set self.value?
                # evaluate formula will return a value or raise an error
                pass
            elif self.contents[0] == "'":
                self.value = self.contents[1:]
            else:
                # evaluate as a literal (number, date, etc.)
                # for now, just numbers and strings
                # TODO: other values (NaN, Infinity, -Infinity) should not be converted to nums
                # can use decimal.is_finite() method to check
                # remove trailing zeros from number (str)
                num = self.contents.rstrip("0")
                if num[-1] == ".":
                    num = num[:-1]
                # store value as a Decimal
                self.value = decimal.Decimal(num)  

        # evaluate formulas in workbook
        return self.value

    def set_contents(self, contents: str):
        # if contents == "" or only whitespace, contents + value are still None
        if contents != "" and contents != len(contents) * " ":
            self.contents = contents.strip()
            # if it's a formula, scan for references
    
class CellErrorType(enum.Enum):
    '''
    This enum specifies the kinds of errors that spreadsheet cells can hold.
    '''

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