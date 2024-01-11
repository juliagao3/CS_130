import sheets
import decimal

from sheets import CellError, CellErrorType

import lark
from lark.visitors import visit_children_decor

def number_arg(index):
    def check(f):
        def new_f(self, values):
            if type(values[index]) == str:
                try:
                    values[index] = decimal.Decimal(value)
                except decimal.InvalidOperation as e:
                    pass

            if type(values[index]) != decimal.Decimal:
                    return CellError(CellErrorType.TYPE_ERROR, f"{value} failed on parsing")

            return f(self, values)
        return new_f
    return check

class FormulaEvaluator(lark.visitors.Interpreter):
    @visit_children_decor
    @number_arg(0)
    @number_arg(2)
    def add_expr(self, values):
        # value_1 = get_number(values[0])
        # value_2 = get_number(values[2])
        # if type(value_1) == CellError:
        #     return value_1
        # elif type(value_2) == CellError:
        #     return value_2
        if values[1] == "+":
            return values[0] + values[2]
        elif values[1] == "-":
            return values[0] - values[2]
        else:
            assert f"Unexpected add_expr operator: {values[1]}"
            
    @visit_children_decor
    def mul_expr(self, values):
        if type(values[0]) == CellError:
            return CellError
        if values[1] == "*":
            return values[0] * values[2]
        elif values[1] == '/' and values[2] != 0:
            return values[0] / values[2]
        else:
            assert f"Unexpected mul_expr operator: {values[1]}"

    @visit_children_decor
    def unary_op(self, values):
        if type(values[1]) != int:
            assert f"Base number {values[1]} is not an integer"
        if values[0] == '+':
            return values[1]
        elif values[0] == '-':
            return -1 * values[1]
        else:
            assert f"Unexpected unary operator: {values[0]}"
        
    @visit_children_decor
    def concat_expr(self, values):
        if type(values[0]) == CellError:
            return CellError
        if values[1] == "&":
            # convert any non-string input to a string
            return str(values[0]) + str(values[2])
        else:
            assert f"Failed string concatenation on {values[0]}, {values[2]}" 

    @visit_children_decor
    def number(self, values):
        # if tree.children[0] is a location
        # if tree.children[0] is a CellError string representation
        return decimal.Decimal(values[0])
    
    @visit_children_decor
    def string(self, values):
        return values[0].value[1:-1]

parser = lark.Lark.open('formulas.lark', start='formula')
evaluator = FormulaEvaluator()

tree = parser.parse('=123.456')
value = evaluator.visit(tree)
print(f'value = {value} (type is {type(value)})')

tree = parser.parse('="Hello!"')
value = evaluator.visit(tree)
print(f'value = {value} (type is {type(value)})')

def test_default_sheet_name():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)
        assert sheet_name == "Sheet 1"

def test_one_plus_one():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", "1")
        wb.set_cell_contents(sheet_name, "A2", "1")
        wb.set_cell_contents(sheet_name, "A3", "=A1+A2")

        assert wb.get_cell_value(sheet_name, "A1") == decimal.Decimal(1)
        assert wb.get_cell_value(sheet_name, "A2") == decimal.Decimal(1)
        assert wb.get_cell_value(sheet_name, "A3") == decimal.Decimal(2)

def test_string_arithmetic():
        wb = sheets.Workbook("wb")
        sheet_num, sheet_name = wb.new_sheet(None)

        wb.set_cell_contents(sheet_name, "A1", "1")
        wb.set_cell_contents(sheet_name, "A2", "hello")
        wb.set_cell_contents(sheet_name, "A3", "=A1+A2")

        assert wb.get_cell_value(sheet_name, "A1") == decimal.Decimal(1)
        assert wb.get_cell_value(sheet_name, "A2") == decimal.Decimal(1)

        a3 = wb.get_cell_value(sheet_name, "A3")

        assert type(a3) == sheets.CellError
        assert a3.type == sheets.CellErrorType.TYPE_ERROR;

def test_all():
        tests = [
                test_default_sheet_name,
                test_one_plus_one,
                test_string_arithmetic
        ]

        for t in tests:
                t()

if __name__ == "__main__":
        test_all()

