import decimal
import lark
from lark.visitors import visit_children_decor
from cell import CellError

class FormulaEvaluator(lark.visitors.Interpreter):
    @visit_children_decor
    def add_expr(self, values):
        if type(values[0]) == CellError:
            return CellError
        
        if values[1] == "+":
        elif values[1] == "-":
        else:

    def number(self, tree):
        return decimal.Decimal(tree.children[0])
    
    def string(self, tree):
        return tree.children[0].value[1:-1]

parser = lark.Lark.open('formulas.lark', start='formula')
evaluator = FormulaEvaluator()

tree = parser.parse('=123.456')
value = evaluator.visit(tree)
print(f'value = {value} (type is {type(value)})')

tree = parser.parse('="Hello!"')
value = evaluator.visit(tree)
print(f'value = {value} (type is {type(value)})')