import sheets

from . import reference
from . import interp

import decimal
import enum

def bool_arg(value):
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, decimal.Decimal):
        return value != 0
    if isinstance(value, str):
        if value.lower() == "true":
            return True
        elif value.lower() == "false":
            return False
        else:
            return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, f"invalid bool string {value}")
    if isinstance(value, sheets.CellError):
        return value
    return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, f"cant make bool from {type(value)}")

def string_arg(v):
    return "" if v is None else (str(v).upper() if isinstance(v, bool) else str(v))

def link_subtree(evaluator, subtree):
    finder = interp.CellRefFinder(evaluator.sheet.sheet_name)
    finder.visit(subtree)

    for sheet_name, location in finder.refs:
        ref = reference.Reference.from_string(location, allow_absolute=True)

        cell = evaluator.workbook.get_cell(sheet_name, ref)

        evaluator.workbook.dependency_graph.link_runtime(evaluator.c, cell)
        evaluator.workbook.sheet_references.link_runtime(evaluator.c, ref.sheet_name or evaluator.sheet.sheet_name)
        evaluator.c.check_references(evaluator.workbook)

class ArgEvaluation(enum.Enum):
    EAGER = 0
    LAZY = 1

def func_version(_evaluator, _args):
    return sheets.version

def func_and(_evaluator, args):
    if len(args) < 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "AND requires at least 1 argument")

    result = True

    errors = []
    for a in args:
        b = bool_arg(a)

        if isinstance(b, sheets.CellError) and b not in errors:
            errors.append(b)

        if not b:
            result = False
            
    if len(errors) > 0:
        return interp.propagate_errors(errors)
    else:
        return result

def func_or(_evaluator, args):
    if len(args) < 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "OR requires at least 1 argument")

    result = False

    errors = []
    for a in args:
        b = bool_arg(a)

        if isinstance(b, sheets.CellError) and b not in errors:
            errors.append(b)

        if b:
            result = True
            
    if len(errors) > 0:
        return interp.propagate_errors(errors)
    else:
        return result
    
def func_not(_evaluator, args):
    if len(args) != 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "NOT requires exactly 1 argument")
    
    return not bool_arg(args[0])

def func_xor(_evaluator, args):
    if len(args) < 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "OR requires at least 1 argument")
     
    count_true = 0
    errors = []
    for a in args:
        b = bool_arg(a)
        
        if isinstance(b, sheets.CellError) and b not in errors:
            errors.append(b)
        
        if b:
            count_true += 1
    
    if len(errors) > 0:
        return interp.propagate_errors(errors) 
    else:
        return (not count_true % 2 == 0)

def func_exact(_evaluator, args):
    if len(args) != 2:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "EXACT requires exactly 2 argument")
    
    if isinstance(args[0], sheets.CellError) or isinstance(args[1], sheets.CellError):
        errors = [args[0], args[1]]
        return interp.propagate_errors(errors) 
    
    return (string_arg(args[0]) == string_arg(args[1]))

def func_if(evaluator, args):
    # lazy!!!
    if len(args) < 2 or len(args) > 3:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "IF requires exactly 2 or 3 arguments")
    
    b = bool_arg(evaluator.visit(args[0]))

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
        index = int(evaluator.visit(args[0]))
        if index < 0 or index >= len(args):
            return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "index is out of bounds")

        link_subtree(evaluator, args[index])

        return evaluator.visit(args[index])
    
    except TypeError:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "index is not an integer")
        
def func_isblank(_evaluator, args):
    if len(args) != 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "ISBLANK requires exactly 1 argument")
    
    if isinstance(args[0], sheets.CellError):
        return args[0]
    
    return (args[0] is None)

def func_iserror(_evaluator, args):
    if len(args) != 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "ISERROR requires exactly 1 argument")
    
    return (isinstance(args[0], sheets.CellError))

def func_indirect(evaluator, args):
    if len(args) != 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "INDIRECT requires exactly 1 argument")

    # todo probably convert to string...
    if not isinstance(args[0], str):
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "argument to INDIRECT should be a cell reference")

    try:
        ref = reference.Reference.from_string(args[0], allow_absolute=True)

        cell = evaluator.workbook.get_cell(ref.sheet_name or evaluator.sheet.sheet_name, ref)

        evaluator.workbook.dependency_graph.link_runtime(evaluator.c, cell)
        evaluator.workbook.sheet_references.link_runtime(evaluator.c, ref.sheet_name or evaluator.sheet.sheet_name)
        evaluator.c.check_references(evaluator.workbook)

        return cell.value
    except KeyError:
        return sheets.CellError(sheets.CellErrorType.BAD_REFERENCE, args[0])
    except ValueError:
        return sheets.CellError(sheets.CellErrorType.BAD_REFERENCE, args[0])

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
}
