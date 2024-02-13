import sheets
import decimal
import lark
from lark.visitors import visit_children_decor
from lark.visitors import v_args
from . import sheet as sheet_util
from . import location as location_utils

from typing import Tuple

parser = lark.Lark.open('formulas.lark', rel_to=__file__, start='formula')

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

            if type(values[index]) == sheets.CellError:
                    return values[index]
            
            if type(values[index]) != decimal.Decimal:
                    return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, f"{values[index]} failed on parsing")

            return f(self, values)
        return new_f
    return check

def strip_quotes(s: str):
    if s[0] != "'":
        return s
    return s[1: -1]

class CellRefFinder(lark.Visitor):
    def __init__(self, sheet_name):
        self.refs = []
        self.sheet_name = sheet_name
        
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
    
        
class FormulaEvaluator(lark.visitors.Interpreter):

    def __init__(self, workbook, sheet):
        self.workbook = workbook
        self.sheet = sheet

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

        try:
            return self.workbook.get_cell_value(sheet_name, cell_ref) 
        except ValueError:
            assert "Uncaught bad reference!!!"
        except KeyError:
            assert "Uncaught bad reference!!!"

    @visit_children_decor
    def concat_expr(self, values):
        return "".join(["" if v is None else str(v) for v in values])

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
    
class FormulaMover(lark.visitors.Transformer_InPlace):
    
    def __init__(self, offset: Tuple[int, int]):
        self.offset = offset

    @v_args(tree=True)
    def cell(self, tree):
        values = tree.children
        # checks and changes the referenced cells in the formula
        location = values[0].lower()

        if len(values) > 1:
            location = values[1].lower()

        # sheet_name = values[0]
        to_loc = location_utils.location_string_to_tuple(location)

        # check if col/row is relative
        col, row = to_loc[0], to_loc[1]

        if not to_loc[2]:
            col = to_loc[0] + self.offset[0]
        if not to_loc[3]:
            row = to_loc[1] + self.offset[1]
        
        to_loc = (col, row)

        # check if new loc is valid
        try:    
            location_utils.check_location_tuple((to_loc[0], to_loc[1]))
        except:
            return "#REF!"

        if len(values) > 1:
            values[1] = location_utils.tuple_to_location_string(to_loc)
        else:
            values[0] = location_utils.tuple_to_location_string(to_loc)
        
        return tree
        
def parse_formula(formula):
    try:
        return parser.parse(formula)
    except lark.exceptions.ParseError:
        return None
    except lark.exceptions.UnexpectedCharacters:
        return None

def evaluate_formula(workbook, sheet, tree):
    evaluator = FormulaEvaluator(workbook, sheet)
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
    
