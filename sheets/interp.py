import sheets
import decimal
import lark
from lark.visitors import visit_children_decor
from lark.visitors import v_args
from . import sheet as sheet_util

parser = lark.Lark.open('formulas.lark', rel_to=__file__, start='formula')

def number_arg(index):
    def check(f):
        def new_f(self, values):
            if values[index] == None:
                values[index] = decimal.Decimal(0)
            elif type(values[index]) == str:
                try:
                    values[index] = decimal.Decimal(values[index])
                except decimal.InvalidOperation as e:
                    pass

            if type(values[index]) == sheets.CellError:
                    return values[index]
            
            if type(values[index]) != decimal.Decimal:
                    return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, f"{values[index]} failed on parsing")

            return f(self, values)
        return new_f
    return check

class CellRefFinder(lark.Visitor):
    def __init__(self):
        self.refs = []
        
    def cell(self, tree):
        if len(tree.children) == 1:
            self.refs.append(str(tree.children[0]))
        else:
            assert len(tree.children) == 2
            self.refs.append('!'.join(tree.children))

class SheetRenamer(lark.visitors.Transformer_InPlace):

    def __init__(self, old_name, new_name):
        self.old_name = old_name
        self.new_name = new_name

    @v_args(tree=True)
    def cell(self, tree):
        values = tree.children
        if len(values) > 1:
            if values[0].lower() == self.old_name.lower():
                values[0] = self.new_name
            need_quotes = sheet_util.name_needs_quotes(values[0])
            if need_quotes and values[0][0] != "'":
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
        
        if sheet_name[0] == "'":
            sheet_name = sheet_name[1:-1]

        try:
            return self.workbook.get_cell_value(sheet_name, cell_ref) 
        except ValueError as e:
            assert "Uncaught bad reference!!!"
        except KeyError as e:
            assert "Uncaught bad reference!!!"

    @visit_children_decor
    def concat_expr(self, values):
        return "".join(["" if v == None else str(v) for v in values])

    @visit_children_decor
    def number(self, values):
        # if values[0] is a location
        # if values[0] is a CellError string representation
        return decimal.Decimal(values[0])
    
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
    finder = CellRefFinder()
    finder.visit(tree)
    return finder.refs

def rename_sheet(old, new, tree):
    renamer = SheetRenamer(old, new)
    renamer.transform(tree)
    printer = FormulaPrinter()
    return "=" + printer.visit(tree) 
