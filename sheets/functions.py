import sheets

import decimal
import enum

from . import error
from . import interp
from . import reference
from . import base_types

from .range import CellRange

def link_subtree(evaluator, subtree):
    finder = interp.CellRefFinder(evaluator.sheet.sheet_name)
    finder.visit(subtree)

    for ref in finder.refs:
        cell = evaluator.workbook.get_cell(ref)

        evaluator.workbook.dependency_graph.link_runtime(evaluator.c, cell)
        evaluator.workbook.sheet_references.link_runtime(evaluator.c, ref.sheet_name or evaluator.sheet.sheet_name)

    evaluator.c.check_cycles(evaluator.workbook)

class ArgEvaluation(enum.Enum):
    EAGER = 0
    LAZY = 1

def func_version(_evaluator, args):
    if len(args) > 0:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "VERSION takes no arguments")

    return sheets.version

def func_and(_evaluator, args):
    if len(args) < 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "AND requires at least 1 argument")

    result = True

    errors = []
    for a in args:
        if type(a) == CellRange:
            return error.CellError(error.CellErrorType.TYPE_ERROR, "Unhandled argument cell range")

        b = base_types.to_bool(a)

        if isinstance(b, sheets.CellError) and b not in errors:
            errors.append(b)

        if not b:
            result = False
            
    if len(errors) > 0:
        return error.propagate_errors(errors)
    else:
        return result

def func_or(_evaluator, args):
    if len(args) < 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "OR requires at least 1 argument")

    result = False

    errors = []
    for a in args:
        if type(a) == CellRange:
            return error.CellError(error.CellErrorType.TYPE_ERROR, "Unhandled argument cell range")

        b = base_types.to_bool(a)

        if isinstance(b, sheets.CellError) and b not in errors:
            errors.append(b)

        if b:
            result = True
            
    if len(errors) > 0:
        return error.propagate_errors(errors)
    else:
        return result
    
def func_not(_evaluator, args):
    if len(args) != 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "NOT requires exactly 1 argument")
    
    if type(args[0]) == CellRange:
            return error.CellError(error.CellErrorType.TYPE_ERROR, "Unhandled argument cell range")
    
    b = base_types.to_bool(args[0])

    if isinstance(b, sheets.CellError):
        return b
    else:
        return not b

def func_xor(_evaluator, args):
    if len(args) < 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "XOR requires at least 1 argument")
     
    count_true = 0
    errors = []
    for a in args:
        if type(a) == CellRange:
            return error.CellError(error.CellErrorType.TYPE_ERROR, "Unhandled argument cell range")

        b = base_types.to_bool(a)
        
        if isinstance(b, sheets.CellError) and b not in errors:
            errors.append(b)
        
        if b:
            count_true += 1
    
    if len(errors) > 0:
        return error.propagate_errors(errors) 
    else:
        return (not count_true % 2 == 0)

def func_exact(_evaluator, args):
    if len(args) != 2:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "EXACT requires exactly 2 argument")
    
    if isinstance(args[0], sheets.CellError) or isinstance(args[1], sheets.CellError):
        errors = [args[0], args[1]]
        return error.propagate_errors(errors) 
    
    if type(args[0]) == CellRange or type(args[1]) == CellRange:
            return error.CellError(error.CellErrorType.TYPE_ERROR, "Unhandled argument cell range")
    
    return (base_types.to_string(args[0]) == base_types.to_string(args[1]))

def func_if(evaluator, args):
    # lazy!!!
    if len(args) < 2 or len(args) > 3:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "IF requires exactly 2 or 3 arguments")
    
    b = base_types.to_bool(evaluator.visit(args[0]))

    if isinstance(b, sheets.CellError):
        return b
        
    if b:
        branch = args[1]
    elif len(args) == 3:
        branch = args[2]
    elif len(args) == 2:
        return False

    link_subtree(evaluator, branch)

    return evaluator.visit(branch)

def func_iferror(evaluator, args):
    # lazy!!!
    if len(args) < 1 or len(args) > 2:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "IFERROR requires exactly 1 or 2 arguments")

    first = evaluator.visit(args[0])

    if not isinstance(first, sheets.CellError):
        return first
    elif len(args) == 2:
        link_subtree(evaluator, args[1])
        return evaluator.visit(args[1])
    else:
        return ""

def func_choose(evaluator, args):
    # lazy!!!
    if len(args) < 2:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "CHOOSE requires at least 2 arguments")
    
    try:

        index_org = evaluator.visit(args[0])

        # Check if error
        if isinstance(index_org, sheets.CellError):
            return index_org

        # Check if bool in string format
        if isinstance(index_org, str):
            if index_org.lower() == "true":
                index_org = 1
            elif index_org.lower() == "false":
                index_org = 0

        # Check if number is integer
        index = int(index_org)          
        index_aux = float(index_org)             
        if abs(index_aux - index) > 0:
            raise TypeError
        
        if index <= 0 or index >= len(args):
            return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "index is out of bounds")

        link_subtree(evaluator, args[index])

        return evaluator.visit(args[index])
    
    except (TypeError, ValueError):
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "index is not an integer")
        
def func_isblank(_evaluator, args):
    if len(args) != 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "ISBLANK requires exactly 1 argument")
    
    # Follow Piazza post regarding how to deal with errors
    # https://piazza.com/class/lqvau3tih6k26o/post/43
    if isinstance(args[0], sheets.CellError):
        if args[0].get_type() in [sheets.CellErrorType.PARSE_ERROR, sheets.CellErrorType.CIRCULAR_REFERENCE]:
            return args[0]
        else:
            return False
        
    if type(args[0]) == CellRange:
            return error.CellError(error.CellErrorType.TYPE_ERROR, "Unhandled argument cell range")
    
    return (args[0] is None)

def func_iserror(_evaluator, args):
    if len(args) != 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "ISERROR requires exactly 1 argument")
    
    if type(args[0]) == CellRange:
            return error.CellError(error.CellErrorType.TYPE_ERROR, "Unhandled argument cell range")
    
    return (isinstance(args[0], sheets.CellError))

def func_indirect(evaluator, args):
    if len(args) != 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "INDIRECT requires exactly 1 argument")

    try:
        r = CellRange.from_string(evaluator.sheet.sheet_name, str(args[0]).lower())

        for ref in r.generate():
            evaluator.workbook.sheet_references.link_runtime(evaluator.c, ref.sheet_name or evaluator.sheet.sheet_name)
            cell = evaluator.workbook.get_cell(ref)
            evaluator.workbook.dependency_graph.link_runtime(evaluator.c, cell)

        evaluator.c.check_cycles(evaluator.workbook)

        return r
    except (KeyError, ValueError):
        pass

    try:
        ref = reference.Reference.from_string(evaluator.sheet.sheet_name, str(args[0]).lower())

        evaluator.workbook.sheet_references.link_runtime(evaluator.c, ref.sheet_name or evaluator.sheet.sheet_name)

        cell = evaluator.workbook.get_cell(ref)

        evaluator.workbook.dependency_graph.link_runtime(evaluator.c, cell)
        evaluator.c.check_cycles(evaluator.workbook)

        # If the argument can be parsed as a cell reference, but is invalid due to an error
        # returns BAD_REFERENCE
        if isinstance(cell.value, sheets.CellError):
            return sheets.CellError(sheets.CellErrorType.BAD_REFERENCE, args[0])
        else:
            return cell.value

    except (KeyError, ValueError):
        return sheets.CellError(sheets.CellErrorType.BAD_REFERENCE, args[0])

def numeric_function(evaluator, args, f):
    if len(args) < 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, f"{f.__name__} requires at least 1 argument")

    numbers = []

    for arg in args:
        if type(arg) == CellRange:
            numbers += list(arg.generate_values(evaluator.workbook))
        else:
            numbers.append(base_types.to_number(arg))
    
    numbers = list(map(lambda i: base_types.to_number(i), filter(lambda i: i is not None, numbers)))

    e = error.propagate_errors(numbers)

    if e is not None:
        return e

    return f(numbers)

def func_min(evaluator, args):
    def custom_min(nums):
        if len(nums) == 0:
            return decimal.Decimal(0)
        return min(nums)
    return numeric_function(evaluator, args, custom_min)

def func_max(evaluator, args):
    def custom_max(nums):
        if len(nums) == 0:
            return decimal.Decimal(0)
        return max(nums)
    return numeric_function(evaluator, args, custom_max)

def func_sum(evaluator, args):
    return numeric_function(evaluator, args, sum)

def func_average(evaluator, args):
    def average(nums):
        if len(nums) == 0:
            return error.CellError(error.CellErrorType.DIVIDE_BY_ZERO, "")
        return sum(nums)/len(nums)
    return numeric_function(evaluator, args, average)

def func_vlookup(evaluator, args):
    if len(args) != 3:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "VLOOKUP requires at exactly 3 arguments")
    
    key = evaluator.visit(args[0])
    region = evaluator.visit(args[1])
    index = base_types.to_number(evaluator.visit(args[2]))

    if type(region) != CellRange:
        return error.CellError(error.CellErrorType.TYPE_ERROR, "")

    e = error.propagate_errors([key, region, index])
    
    if e is not None:
        return e
    
    if index < 1 or index > region.end_ref.col - region.start_ref.col + 1:
        return error.CellError(error.CellErrorType.TYPE_ERROR, "")

    search_values = []
    target_values = []

    for ref in region.generate_column(0):
        cell = evaluator.workbook.get_cell(ref)
        evaluator.workbook.dependency_graph.link_runtime(evaluator.c, cell)
        evaluator.workbook.sheet_references.link_runtime(evaluator.c, ref.sheet_name or evaluator.sheet.sheet_name)
        search_values.append(cell.value)

    for ref in region.generate_column(index-1):
        cell = evaluator.workbook.get_cell(ref)
        evaluator.workbook.dependency_graph.link_runtime(evaluator.c, cell)
        evaluator.workbook.sheet_references.link_runtime(evaluator.c, ref.sheet_name or evaluator.sheet.sheet_name)
        target_values.append(cell.value)

    evaluator.c.check_cycles(evaluator.workbook)

    for search_value, target_value in zip(search_values, target_values):
        if search_value == key:
            return target_value

    return error.CellError(error.CellErrorType.TYPE_ERROR, "")

def func_hlookup(evaluator, args):
    if len(args) != 3:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "HLOOKUP requires at exactly 3 arguments")
    
    key = evaluator.visit(args[0])
    region = evaluator.visit(args[1])
    index = base_types.to_number(evaluator.visit(args[2]))

    if type(region) != CellRange:
        return error.CellError(error.CellErrorType.TYPE_ERROR, "")

    e = error.propagate_errors([key, region, index])
    
    if e is not None:
        return e
    
    if index < 1 or index > region.end_ref.row - region.start_ref.row + 1:
        return error.CellError(error.CellErrorType.TYPE_ERROR, "")

    search_values = []
    target_values = []

    for ref in region.generate_row(0):
        cell = evaluator.workbook.get_cell(ref)
        evaluator.workbook.dependency_graph.link_runtime(evaluator.c, cell)
        evaluator.workbook.sheet_references.link_runtime(evaluator.c, ref.sheet_name or evaluator.sheet.sheet_name)
        search_values.append(cell.value)

    for ref in region.generate_row(index-1):
        cell = evaluator.workbook.get_cell(ref)
        evaluator.workbook.dependency_graph.link_runtime(evaluator.c, cell)
        evaluator.workbook.sheet_references.link_runtime(evaluator.c, ref.sheet_name or evaluator.sheet.sheet_name)
        target_values.append(cell.value)

    evaluator.c.check_cycles(evaluator.workbook)

    for search_value, target_value in zip(search_values, target_values):
        if search_value == key:
            return target_value

    return error.CellError(error.CellErrorType.TYPE_ERROR, "")

functions = {
    "version":  (ArgEvaluation.EAGER, func_version  ),
    "and":      (ArgEvaluation.EAGER, func_and      ),
    "or":       (ArgEvaluation.EAGER, func_or       ),
    "not":      (ArgEvaluation.EAGER, func_not      ),
    "xor":      (ArgEvaluation.EAGER, func_xor      ),
    "exact":    (ArgEvaluation.EAGER, func_exact    ),
    "if":       (ArgEvaluation.LAZY,  func_if       ),
    "iferror":  (ArgEvaluation.LAZY,  func_iferror  ),
    "choose":   (ArgEvaluation.LAZY,  func_choose   ),
    "isblank":  (ArgEvaluation.EAGER, func_isblank  ),
    "iserror":  (ArgEvaluation.EAGER, func_iserror  ),
    "indirect": (ArgEvaluation.EAGER, func_indirect ),
    "min":      (ArgEvaluation.EAGER, func_min),
    "max":      (ArgEvaluation.EAGER, func_max),
    "sum":      (ArgEvaluation.EAGER, func_sum),
    "average":  (ArgEvaluation.EAGER, func_average),
    "vlookup":  (ArgEvaluation.LAZY,  func_vlookup),
    "hlookup":  (ArgEvaluation.LAZY,  func_hlookup)
}
