import copy
import decimal

from typing import Tuple

from . import interp

from .error import CellError, CellErrorType, FormulaError
from .range import CellRange

def is_empty_content_string(contents):
    return contents is None or contents == "" or contents.isspace()

def remove_trailing_zeros(d: decimal.Decimal):
    num = str(d)
    e = num.rfind("E") if "E" in num else len(num)
    if "." in num:
        num = num[:e].rstrip("0") + num[e:]
    if num[-1] == ".":
        num = num[:-1]
    return decimal.Decimal(num)

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

    def get_value(self):
        return self.value

    def parse_formula(self):
        self.formula_tree = interp.parse_formula(self.contents)

        if self.formula_tree is None:
            raise FormulaError(CellError(CellErrorType.PARSE_ERROR, ""))

    def check_references(self, workbook):
        if self.formula_tree is None:
            return

        static_refs, all_refs = interp.find_refs(workbook, self.sheet, self.formula_tree)

        # link to all referenced sheet names - even if they're not used
        for ref in all_refs:
            workbook.sheet_references.link(self, ref.sheet_name.lower())

        # only link to the cells that are used in evaluation every time
        # (the static references)
        for ref in static_refs:
            try:
                ref.check_bounds()
                cell = workbook.get_cell(ref.sheet_name, ref)
                workbook.dependency_graph.link(self, cell)
            except (KeyError, ValueError):
                pass
                
    def check_cycles(self, workbook):
        for cycle in workbook.dependency_graph.get_cycles():
            if self in cycle:
                raise FormulaError(CellError(CellErrorType.CIRCULAR_REFERENCE, ""))

    def evaluate_formula(self, workbook):
        self.check_cycles(workbook)

        value = interp.evaluate_formula(workbook, self.sheet, self, self.formula_tree)

        if value is None:
            value = decimal.Decimal(0)

        if type(value) == decimal.Decimal:
            value = remove_trailing_zeros(value)

        if type(value) == CellRange:
            ref = next(value.generate())
            value = workbook.get_cell(ref.sheet_name, ref).value

        self.set_value(value)

    def rename_sheet(self, workbook, old_name, new_name):
        if self.formula_tree is None:
            return
        self.set_contents(workbook, interp.rename_sheet(old_name, new_name, self.formula_tree))

    def move_formula(self, workbook, offset):
        if self.formula_tree is None:
            return
        self.set_contents(workbook, interp.move_formula(offset, self.formula_tree))

    def recompute_value(self, workbook):
        if self.contents is None or self.formula_tree is None:
            return
        try:
            workbook.sheet_references.clear_forward_runtime_links((self.sheet, self))
            workbook.dependency_graph.clear_forward_runtime_links(self)

            self.check_references(workbook)
            self.evaluate_formula(workbook)
        except FormulaError as e:
            self.set_value(e.value)

    def copy_cell(self, other_cell, workbook, offset: Tuple[int, int]):
        workbook.sheet_references.clear_forward_links((self.sheet, self))
        workbook.dependency_graph.clear_forward_links(self)

        self.contents = other_cell.contents
        self.formula_tree = copy.deepcopy(other_cell.formula_tree)
        self.value = other_cell.value

        if self.formula_tree is not None:
            self.contents = interp.move_formula(offset, self.formula_tree)
            self.check_references(workbook)

    def set_contents(self, workbook, contents: str, evaluate_formulas = True):
        workbook.sheet_references.clear_forward_links((self.sheet, self))
        workbook.dependency_graph.clear_forward_links(self)
        self.formula_tree = None

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

                if evaluate_formulas:
                    self.evaluate_formula(workbook)
            except FormulaError as e:
                self.set_value(e.value)
        elif contents[0] == "'":
            self.set_value(contents[1:])
        elif contents.lower() == "true":
            self.set_value(True)
        elif contents.lower() == "false":
            self.set_value(False)
        elif CellError.from_string(contents) is not None:
            self.set_value(CellError(CellError.from_string(contents), ""))
        else:
            try:
                self.set_value(remove_trailing_zeros(decimal.Decimal(contents)))
                if not self.value.is_finite():
                    self.set_value(contents)
            except decimal.InvalidOperation:
                self.set_value(contents)
