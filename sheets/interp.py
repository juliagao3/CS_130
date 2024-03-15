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
from .range     import CellRange

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
        try:
            ref = Reference.from_string(self.sheet_name, str(tree.children[0]))
        except ValueError:
            # don't include bad references
            return

        self.refs.append(ref)

        if self.static_context:
            self.static_refs.append(ref)

    def cell_range(self, tree):
        try:
            start = str(tree.children[0].children[0])
            end = str(tree.children[1].children[0])
            r = CellRange(self.sheet_name, start, end)
        except ValueError:
            # don't include bad ranges
            return

        for ref in r.generate():
            self.refs.append(ref)

            if self.static_context:
                self.static_refs.append(ref)


class SheetRenamer(lark.visitors.Transformer_InPlace):

    def __init__(self, old_name, new_name):
        self.old_name = old_name
        self.new_name = new_name

    @v_args(tree=True)
    def cell(self, tree):
        try:
            # don't need to pass sheet name here cause we're just doing a
            # cosmetic change
            ref = Reference.from_string(None, str(tree.children[0]))

            if ref.sheet_name is not None and ref.sheet_name.lower() == self.old_name.lower():
                ref.sheet_name = self.new_name
                tree.children[0] = str(ref)
        except ValueError:
            # don't include bad references
            pass

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
        return values[0]
    
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

    CMP_OPS = {
                "=":  lambda a, b: a == b,
                "==": lambda a, b: a == b,
                "<>": lambda a, b: a != b,
                "!=": lambda a, b: a != b,
                ">":  lambda a, b: a > b,
                "<":  lambda a, b: a < b,
                ">=": lambda a, b: a >= b,
                "<=": lambda a, b: a <= b
            }

    CMP_DEFAULTS = {
                bool: False,
                str: "",
                decimal.Decimal: decimal.Decimal("0"),
                type(None): None
            }

    def __init__(self, workbook, sheet, cell):
        self.workbook = workbook
        self.sheet = sheet
        self.c = cell

    @visit_children_decor
    def cmp_expr(self, values):
        e = error.propagate_errors([values[0], values[2]])

        if e is not None:
            return e

        if values[0] is None:
            values[0] = FormulaEvaluator.CMP_DEFAULTS[type(values[2])]

        if values[2] is None:
            values[2] = FormulaEvaluator.CMP_DEFAULTS[type(values[0])]

        if values[0] is None and values[2] is None:
            values[0] = "None"
            values[2] = "None"

        if isinstance(values[0], type(values[2])):
            if isinstance(values[0], str):
                values[0] = values[0].lower()

            if isinstance(values[2], str):
                values[2] = values[2].lower()

            if values[1] not in FormulaEvaluator.CMP_OPS:
                assert f"Unexpected cmp_expr operator: {values[1]}"

            return FormulaEvaluator.CMP_OPS[values[1]](values[0], values[2])
        else:
            types = {decimal.Decimal: 0, str: 1, bool: 2}
            return FormulaEvaluator.CMP_OPS[values[1]](types[type(values[0])], types[type(values[2])])

    @visit_children_decor
    def add_expr(self, values):
        left = base_types.to_number(values[0])
        op = values[1]
        right = base_types.to_number(values[2])

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
        try:
            ref = Reference.from_string(self.sheet.sheet_name, values[0])
            ref.check_bounds()
            ref.abs_col = False
            ref.abs_row = False
            return self.workbook.get_cell(ref).value
        except (ValueError, KeyError):
            return CellError(CellErrorType.BAD_REFERENCE, values[0])

    def cell_range(self, tree):
        try:
            start = str(tree.children[0].children[0])
            end = str(tree.children[1].children[0])
            return CellRange(self.sheet.sheet_name, start, end)
        except (ValueError, KeyError):
            return CellError(CellErrorType.BAD_REFERENCE, f"{tree.children[0]}:{tree.children[1]}")

    @visit_children_decor
    def concat_expr(self, values):
        e = error.propagate_errors(values)

        if e is not None:
            return e

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
        return CellError(CellError.from_string(values[0]), "")
    
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
        try:
            # We can safely pass None here because we don't use this reference
            # to get a cell value.
            ref = Reference.from_string(None, str(tree.children[0]))
        except ValueError:
            # https://piazza.com/class/lqvau3tih6k26o/post/33
            # :>
            return tree

        # check if new loc is valid
        try:
            new_value = str(ref.moved(self.offset).check_bounds())
        except ValueError:
            new_value = "#REF!"
        
        tree.children[0] = new_value

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
    
