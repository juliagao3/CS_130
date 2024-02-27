import sheets

from . import reference

import decimal

def bool_arg(value):
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
    return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, f"cant make bool from {type(value)}")

def func_version(_wb, _sheet, _args):
    return sheets.version

def func_and(_wb, _sheet, args):
    if len(args) < 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "AND requires at least 1 argument")

    result = True

    for a in args:
        b = bool_arg(a)

        if isinstance(b, sheets.CellError):
            return b

        if not b:
            result = False

    return result

def func_or(_wb, _sheet, args):
    if len(args) < 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "OR requires at least 1 argument")

    result = False

    for a in args:
        b = bool_arg(a)

        if isinstance(b, sheets.CellError):
            return b

        if b:
            result = True

    return result

def func_not(wb, _sheet, args):
    if len(args) != 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "NOT requires exactly 1 argument")
    
    return not bool_arg(args[0])

def func_xor(_wb, _sheet, args):
    if len(args) < 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "OR requires at least 1 argument")
     
    count_true = 0
    for a in args:
        b = bool_arg(a)
        
        if isinstance(b, sheets.CellError):
            return b
        
        if b:
            count_true += 1
        
    if (count_true % 2 == 0):
        return False
    
    return True

def func_exact(_wb, _sheet, args):
    if len(args) != 2:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "EXACT requires exactly 1 argument")
    
    if isinstance(args[0], sheets.CellError):
        return args[0]
    
    if isinstance(args[1], sheets.CellError):
        return args[1]
    
    return (str(args[0]) == str(args[1]))

def func_if(_wb, _sheet, args):
    if len(args) < 2 or len(args) > 3:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "IF requires exactly 2 or 3 arguments")
    
    if isinstance(args[1], sheets.CellError):
        return args[1]


    if len(args) > 2 and isinstance(args[2], sheets.CellError):
        return args[2]
    
    b = bool_arg(args[0])
    if b:
        return str(args[1])
    elif not b and len(args) == 2:
        return False
    elif not b and len(args) == 3:
        return str(args[2])

def func_iferror(wb, sheet, args):
    if len(args) < 1 or len(args) > 2:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "IFERROR requires exactly 1 or 2 arguments")
    
    if not isinstance(args[0], sheets.CellError):
        return args[0]
    elif len(args) == 2:
        return args[1]
    else:
        return ""

def func_choose(_wb, _sheet, args):
    if len(args) < 2:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "CHOOSE requires at least 2 arguments")
    
    try:
        index = int(args[0])
        if index < 0 or index >= len(args):
            return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "index is out of bounds")
        return args[index]
    except TypeError:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "index is not an integer")
        
def func_isblank(_wb, _sheet, args):
    if len(args) != 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "ISBLANK requires exactly 1 argument")
    
    if isinstance(args[0], sheets.CellError):
        return args[0]
    
    return (args[0] is None)

def func_iserror(_wb, _sheet, args):
    if len(args) != 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "ISERROR requires exactly 1 argument")
    
    return (isinstance(args[0], sheets.CellError))

def func_indirect(wb, sheet, args):
    if len(args) != 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "INDIRECT requires exactly 1 argument")

    # todo probably convert to string...
    if not isinstance(args[0], str):
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "argument to INDIRECT should be a cell reference")

    try:
        ref = reference.Reference.from_string(args[0], allow_absolute=True)

        return wb.get_cell(ref.sheet_name or sheet.sheet_name, ref).value
    except KeyError:
        return sheets.CellError(sheets.CellErrorType.BAD_REFERENCE, args[0])
    except ValueError:
        return sheets.CellError(sheets.CellErrorType.BAD_REFERENCE, args[0])

functions = {
    "version":  func_version,
    "and":      func_and,
    "or":       func_or,
    "not":      func_not,
    "xor":      func_xor,
    "exact":    func_exact,
    "if":       func_if,
    "iferror":  func_iferror,
    "choose":   func_choose,
    "isblank":  func_isblank,
    "iserror":  func_iserror,
    "indirect": func_indirect,
}
