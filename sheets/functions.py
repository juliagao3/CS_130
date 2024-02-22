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

def func_indirect(wb, sheet, args):
    if len(args) != 1:
        return sheets.CellError(sheets.CellErrorType.TYPE_ERROR, "INDIRECT requires exactly 1 argument")

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
    "indirect": func_indirect,
}
