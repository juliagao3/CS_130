import decimal

import lark
from lark.visitors import visit_children_decor
from lark.visitors import v_args

from typing import Tuple

from . import base_types
from . import error
from . import functions

from .error     import CellError, CellErrorType
from .reference import Reference

parser = lark.Lark.open('formulas.lark', rel_to=__file__, start='formula')

def remove_trailing_zeros(d: decimal.Decimal):
    num = str(d)
    e = num.rfind("E") if "E" in num else len(num)
    if "." in num:
        num = num[:e].rstrip("0") + num[e:]
    if num[-1] == ".":
        num = num[:-1]
    return decimal.Decimal(num)

def strip_quotes(s: str):
    if s[0] != "'":
        return s
    return s[1: -1]

class CellRefFinder(lark.visitors.Interpreter):
    def __init__(self, sheet_name):
        self.refs = []
        self.static_refs = []
        self.sheet_name = sheet_name
        self.static_context = True

    def func_expr(self, tree):
        name = str(tree.children[0]).lower()

        if name not in functions.functions:
            return CellError(CellErrorType.BAD_NAME, f"unrecognized function {name}")

        try:

            if functions.functions[name][0] == functions.ArgEvaluation.LAZY:
                # first child could be a static reference (meaning it must be
                # evaluated every time)
                self.visit(tree.children[1])

                # further arguments are not static, but we need to know about
                # them for sheet renaming...
                old = self.static_context
                self.static_context = False
                for child in tree.children[2:]:
                    self.visit(child)
                self.static_context = old
            else:
                self.visit_children(tree)

        except IndexError:
            return CellError(CellErrorType.TYPE_ERROR, "function requires at least one argument")

    def cell(self, tree):
        if len(tree.children) == 1:
            ref = (self.sheet_name, tree.children[0])
        else:
            assert len(tree.children) == 2
            ref = (strip_quotes(tree.children[0]).lower(), tree.children[1])

        self.refs.append(ref)

        if self.static_context:
            self.static_refs.append(ref)

class SheetRenamer(lark.visitors.Transformer_InPlace):

    def __init__(self, old_name, new_name):
        self.old_name = old_name
        self.new_name = new_name

    @v_args(tree=True)
    def cell(self, tree):
        values = tree.children
        if len(values) > 1:
            values[0] = strip_quotes(values[0])
            if values[0].lower() == self.old_name.lower():
                values[0] = self.new_name
            
            if base_types.sheet_name_needs_quotes(values[0]):
                values[0] = "'" + values[0] + "'"
        return tree
    
class FormulaPrinter(lark.visitors.Interpreter):
    
    @visit_children_decor
    def cmp_expr(self, values):
        return " ".join(values)

    @visit_children_decor
    def add_expr(self, values):
        return " ".join(values)
    
    @visit_children_decor
    def mul_expr(self, values):
        return " ".join(values)
    
    @visit_children_decor
    def concat_expr(self, values):
        return " & ".join(values)

    @visit_children_decor
    def unary_op(self, values):
        return "".join(values)
    
    @visit_children_decor
    def cell(self, values):
        return "!".join(values)
    
    @visit_children_decor
    def cell_range(self, values):
        return ":".join(values)
    
    @visit_children_decor
    def number(self, values):
        return str(values[0])
    
    @visit_children_decor
    def string(self, values):
        return values[0]
    
    @visit_children_decor
    def error(self, values):
        return values[0]
    
    @visit_children_decor
    def parens(self, values):
        return '(' + values[0] + ')'

    @visit_children_decor
    def boolean(self, values):
        return values[0]

    @visit_children_decor
    def func_expr(self, values):
        return values[0] + "(" + ", ".join(values[1:]) + ")"
        
class FormulaEvaluator(lark.visitors.Interpreter):

    def __init__(self, workbook, sheet, cell):
        self.workbook = workbook
        self.sheet = sheet
        self.c = cell

    @visit_children_decor
    def cmp_expr(self, values):
        ops = { "=":  lambda a, b: a == b,
                "==": lambda a, b: a == b,
                "<>": lambda a, b: a != b,
                "!=": lambda a, b: a != b,
                ">":  lambda a, b: a > b,
                "<":  lambda a, b: a < b,
                ">=": lambda a, b: a >= b,
                "<=": lambda a, b: a <= b }

        defaults = {
            bool: False,
            str: "",
            decimal.Decimal: decimal.Decimal("0"),
            type(None): None
        }

        e = error.propagate_errors([values[0], values[2]])

        if e is not None:
            return e

        if values[0] is None:
            values[0] = defaults[type(values[2])]

        if values[2] is None:
            values[2] = defaults[type(values[0])]

        if values[0] is None and values[2] is None:
            values[0] = "None"
            values[2] = "None"

        if isinstance(values[0], type(values[2])):
            if isinstance(values[0], str):
                values[0] = values[0].lower()

            if isinstance(values[2], str):
                values[2] = values[2].lower()

            if values[1] not in ops:
                assert f"Unexpected cmp_expr operator: {values[1]}"

            return ops[values[1]](values[0], values[2])
        else:
            types = {decimal.Decimal: 0, str: 1, bool: 2}
            return ops[values[1]](types[type(values[0])], types[type(values[2])])

    @visit_children_decor
    def add_expr(self, values):
        left = base_types.to_number(values[0])
        op = values[1]
        right = base_types.to_number(values[2])

        print(left, right, op)
        e = error.propagate_errors([left, right])

        if e is not None:
            return e

        if op == "+":
            return left + right
        elif op == "-":
            return left - right
        else:
            assert f"Unexpected add_expr operator: {values[1]}"
            
    @visit_children_decor
    def mul_expr(self, values):
        left = base_types.to_number(values[0])
        op = values[1]
        right = base_types.to_number(values[2])

        e = error.propagate_errors([left, right])

        if e is not None:
            return e

        if op == "*":
            return left * right
        elif op == '/':
            if right == decimal.Decimal(0):
                return CellError(CellErrorType.DIVIDE_BY_ZERO, "")
            else:
                return left / right
        else:
            assert f"Unexpected mul_expr operator: {values[1]}"

    @visit_children_decor
    def unary_op(self, values):
        op = values[0]
        value = base_types.to_number(values[1])

        if isinstance(value, CellError):
            return value

        if op == '+':
            return value
        elif op == '-':
            return -1 * value
        else:
            assert f"Unexpected unary operator: {op}"
        
    @visit_children_decor
    def cell(self, values):
        sheet_name = self.sheet.sheet_name
        cell_ref = values[0]

        if len(values) > 1:
            sheet_name = values[0]
            cell_ref = values[1]
        
        sheet_name = strip_quotes(sheet_name)

        ref = Reference.from_string(cell_ref, allow_absolute=True)
        ref.abs_col = False
        ref.abs_row = False
        cell_ref = str(ref)

        try:
            return self.workbook.get_cell_value(sheet_name, cell_ref) 
        except ValueError:
            assert "Uncaught bad reference!!!"
        except KeyError:
            assert "Uncaught bad reference!!!"

    @visit_children_decor
    def concat_expr(self, values):
        return "".join(map(base_types.to_string, values))

    @visit_children_decor
    def number(self, values):
        # if values[0] is a location
        # if values[0] is a CellError string representation
        return remove_trailing_zeros(decimal.Decimal(values[0]))
    
    @visit_children_decor
    def string(self, values):
        # if values[0] is a cell location
        return values[0].value[1:-1]
    
    @visit_children_decor
    def error(self, values):
        return CellError(CellErrorType.from_string(values[0]), "")
    
    @visit_children_decor
    def parens(self, values):
        return values[0]

    @visit_children_decor
    def boolean(self, values):
        return values[0].lower() == "true"
    
    def func_expr(self, tree):
        name = str(tree.children[0])

        if name.lower() not in functions.functions:
            return CellError(CellErrorType.BAD_NAME, f"unrecognized function {name}")
        
        arg_evaluation, f = functions.functions[name.lower()]

        if arg_evaluation == functions.ArgEvaluation.LAZY:
            args = tree.children[1:]
            return f(self, args)
        elif arg_evaluation == functions.ArgEvaluation.EAGER:
            args = self.visit_children(tree)[1:]
            return f(self, args)
        else:
            assert f"Invalid ArgEvaluation: {arg_evaluation}!"

class FormulaMover(lark.visitors.Transformer_InPlace):
    
    def __init__(self, offset: Tuple[int, int]):
        self.offset = offset

    @v_args(tree=True)
    def cell(self, tree):
        values = tree.children
        # checks and changes the referenced cells in the formula
        location = values[0]

        if len(values) > 1:
            location = values[1]

        try:
            ref = Reference.from_string(location, allow_absolute=True)
        except ValueError:
            # https://piazza.com/class/lqvau3tih6k26o/post/33
            # :>
            return tree

        # check if new loc is valid
        try:
            new_value = str(ref.moved(self.offset))
        except ValueError:
            new_value = "#REF!"

        if len(values) > 1:
            values[1] = new_value
        else:
            values[0] = new_value
        
        return tree
        
def parse_formula(formula):
    try:
        return parser.parse(formula)
    except lark.exceptions.ParseError:
        return None
    except lark.exceptions.UnexpectedCharacters:
        return None

def evaluate_formula(workbook, sheet, cell, tree):
    evaluator = FormulaEvaluator(workbook, sheet, cell)
    return evaluator.visit(tree)

def find_refs(workbook, sheet, tree):
    finder = CellRefFinder(sheet.sheet_name.lower())
    finder.visit(tree)
    return (finder.static_refs, finder.refs)

def rename_sheet(old, new, tree):
    renamer = SheetRenamer(old, new)
    renamer.transform(tree)
    printer = FormulaPrinter()
    return "=" + printer.visit(tree) 

def move_formula(offset: Tuple[int, int], tree):
    mover = FormulaMover(offset)
    mover.transform(tree)
    printer = FormulaPrinter()
    return "=" + printer.visit(tree)
    
