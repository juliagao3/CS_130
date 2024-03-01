import sheets
import decimal
import lark
from lark.visitors import visit_children_decor
from lark.visitors import v_args
from . import sheet as sheet_util
from . import functions
from .reference import Reference

from typing import Tuple, List, Any

parser = lark.Lark.open('formulas.lark', rel_to=__file__, start='formula')

def propagate_errors(values: List[Any]):
    def get_error_priority(value):
        priorities = {
            sheets.CellErrorType.PARSE_ERROR:           6,
            sheets.CellErrorType.CIRCULAR_REFERENCE:    5,
            sheets.CellErrorType.BAD_REFERENCE:         4,
            sheets.CellErrorType.BAD_NAME:              4,
            sheets.CellErrorType.TYPE_ERROR:            4,
            sheets.CellErrorType.DIVIDE_BY_ZERO:        4,
        }

        if isinstance(value, sheets.CellError):
            return priorities[value.get_type()]
        else:
            return 0

    error = max(values, default=None, key=get_error_priority)

    if not isinstance(error, sheets.CellError):
        return None
    else:
        return error

def remove_trailing_zeros(d: decimal.Decimal):
    num = str(d)
    e = num.rfind("E") if "E" in num else len(num)
    if "." in num:
        num = num[:e].rstrip("0") + num[e:]
    if num[-1] == ".":
        num = num[:-1]
    return decimal.Decimal(num)

def number_arg(index):
    def check(f):
        def new_f(self, values):
            if values[index] is None:
                values[index] = decimal.Decimal(0)
            elif isinstance(values[index], str):
                try:
                    values[index] = decimal.Decimal(values[index])
                except decimal.InvalidOperation:
                    pass

            if isinstance(values[index], sheets.CellError):
                    return values[index]

            if not isinstance(values[index], decimal.Decimal) and not isinstance(values[index], bool):
                    return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, f"{values[index]} failed on parsing")

            value = f(self, values)

            if isinstance(value, decimal.Decimal):
                value = remove_trailing_zeros(value)

            return value
        return new_f
    return check

def string_arg(v):
    return "" if v is None else (str(v).upper() if isinstance(v, bool) else str(v))

def strip_quotes(s: str):
    if s[0] != "'":
        return s
    return s[1: -1]

class CellRefFinder(lark.visitors.Interpreter):
    def __init__(self, sheet_name):
        self.refs = []
        self.sheet_name = sheet_name

    def func_expr(self, tree):
        name = str(tree.children[0]).lower()

        if name not in functions.functions:
            return sheets.CellError(sheets.CellErrorType.BAD_NAME, f"unrecognized function {name}")

        if functions.functions[name][0] == functions.ArgEvaluation.LAZY:
            self.visit(tree.children[1])
        else:
            self.visit_children(tree)

    def cell(self, tree):
        if len(tree.children) == 1:
            self.refs.append((self.sheet_name, tree.children[0]))
        else:
            assert len(tree.children) == 2
            self.refs.append((strip_quotes(tree.children[0]).lower(), tree.children[1]))

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
            
            if sheet_util.name_needs_quotes(values[0]):
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

        error = propagate_errors([values[0], values[2]])

        if error is not None:
            return error

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
    @number_arg(0)
    @number_arg(2)
    def add_expr(self, values):
        if values[1] == "+":
            return values[0] + values[2]
        elif values[1] == "-":
            return values[0] - values[2]
        else:
            assert f"Unexpected add_expr operator: {values[1]}"
            
    @visit_children_decor
    @number_arg(0)
    @number_arg(2)
    def mul_expr(self, values):
        if values[1] == "*":
            return values[0] * values[2]
        elif values[1] == '/':
            if values[2] == decimal.Decimal(0):
                return sheets.CellError(sheets.CellErrorType.DIVIDE_BY_ZERO, "")
            else:
                return values[0] / values[2]
        else:
            assert f"Unexpected mul_expr operator: {values[1]}"

    @visit_children_decor
    @number_arg(1)
    def unary_op(self, values):
        if values[0] == '+':
            return values[1]
        elif values[0] == '-':
            return -1 * values[1]
        else:
            assert f"Unexpected unary operator: {values[0]}"
        
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
        return "".join(map(string_arg, values))

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
        return sheets.CellError(sheets.CellErrorType.from_string(values[0]), "")
    
    @visit_children_decor
    def parens(self, values):
        return values[0]

    @visit_children_decor
    def boolean(self, values):
        return values[0].lower() == "true"
    
    def func_expr(self, tree):
        name = str(tree.children[0])

        if name.lower() not in functions.functions:
            return sheets.CellError(sheets.CellErrorType.BAD_NAME, f"unrecognized function {name}")
        
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
    return finder.refs

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
    
