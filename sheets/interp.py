import sheets
import decimal
import lark
from lark.visitors import visit_children_decor

parser = lark.Lark.open('formulas.lark', rel_to=__file__, start='formula')

def number_arg(index):
    def check(f):
        def new_f(self, values):
            if type(values[index]) == str:
                try:
                    values[index] = decimal.Decimal(values[index])
                except decimal.InvalidOperation as e:
                    pass

            if type(values[index]) != decimal.Decimal:
                    return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, f"{values[index]} failed on parsing")

            return f(self, values)
        return new_f
    return check

class CellRefFinder(lark.Visitor):
    def __init__(self):
        self.refs = []
        
    def cell(self, tree):
        #print(tree.children)
        if len(tree.children) == 1:
            self.refs.append(str(tree.children[0]))
        else:
            assert len(tree.children) == 2
            self.refs.append('!'.join(tree.children))

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
        if values[2] == decimal.Decimal(0):
            return sheets.CellError(sheets.CellErrorType.DIVIDE_BY_ZERO, "");
        if values[1] == "*":
            return values[0] * values[2]
        elif values[1] == '/':
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
        
        try:
            return self.workbook.get_cell_value(sheet_name, cell_ref) 
        except ValueError as e:
            return sheets.CellError(sheets.CellErrorType.BAD_REFERENCE, "!".join([sheet_name, cell_ref]), e)
        except KeyError as e:
            return sheets.CellError(sheets.CellErrorType.BAD_REFERENCE, "!".join([sheet_name, cell_ref]), e)

    @visit_children_decor
    def concat_expr(self, values):
        return "".join([str(v) for v in values])

    @visit_children_decor
    def number(self, values):
        # if values[0] is a location
        # if values[0] is a CellError string representation
        return decimal.Decimal(values[0])
    
    @visit_children_decor
    def string(self, values):
        # if values[0] is a cell location
        return values[0].value[1:-1]

def evaluate_formula(workbook, sheet, formula):
    evaluator = FormulaEvaluator(workbook, sheet)
    tree = parser.parse(formula)
    value = evaluator.visit(tree)
    return value

def find_refs(formula):
    finder = CellRefFinder()
    tree = parser.parse(formula)
    finder.visit(tree)
    return finder.refs