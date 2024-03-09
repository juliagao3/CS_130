import decimal
import re

from typing import Any

from .error import CellError, CellErrorType

def sheet_name_is_valid(name: str):
    if name is None or name == "" or name.isspace():
        return False

    disallowed = r'[^A-Za-z0-9.?!,:;@#$%^&*()-_\s+]'

    if re.search(disallowed, name) is not None:
        return False

    return True

def sheet_name_needs_quotes(name: str):
    assert len(name) > 0

    if name[0] != '_' and not name[0].isalpha():
        return True

    for c in name:
        if c != '_' and not c.isalpha() and not c.isdigit():
            return True

    return False

def to_number(value: Any) -> decimal.Decimal:
    if value is None:
        return decimal.Decimal("0")
    elif isinstance(value, CellError):
        return value
    elif isinstance(value, decimal.Decimal):
        return value
    elif isinstance(value, str):
        try:
            return decimal.Decimal(value)
        except decimal.InvalidOperation:
            return CellError(CellErrorType.TYPE_ERROR, f"{value} can't be parsed as a number")
    elif isinstance(value, bool):
        return decimal.Decimal("1") if value else decimal.Decimal("0")

    return CellError(CellErrorType.TYPE_ERROR, f"failed to convert type {type(value)} to a number")

def to_bool(value: Any) -> bool:
    if value is None:
        return False
    elif isinstance(value, CellError):
        return value
    elif isinstance(value, decimal.Decimal):
        return value != decimal.Decimal("0")
    elif isinstance(value, str):
        if value.lower() == "true":
            return True
        elif value.lower() == "false":
            return False
        else:
            return CellError(CellErrorType.TYPE_ERROR, f"invalid bool string {value}")
    elif isinstance(value, bool):
        return value

    return CellError(CellErrorType.TYPE_ERROR, f"failed to convert type {type(value)} to a bool")

def to_string(value: Any) -> str:
    if value is None:
        return ""
    elif isinstance(value, CellError):
        return str(value)
    elif isinstance(value, decimal.Decimal):
        return str(value)
    elif isinstance(value, str):
        return value
    elif isinstance(value, bool):
        return str(value).upper()

    return CellError(CellErrorType.TYPE_ERROR, f"failed to convert type {type(value)} to a string")

def lt(a, b):
    if isinstance(a, type(b)):
        if isinstance(a, str):
            a = a.lower()

        if isinstance(b, str):
            b = b.lower()

        if isinstance(a, CellError):
            a = a.get_type()

        if isinstance(b, CellError):
            b = b.get_type()

        return a < b
    else:
        types = {type(None): 0, CellError: 1, decimal.Decimal: 2, str: 3, bool: 4}
        return types[type(a)] < types[type(b)]
