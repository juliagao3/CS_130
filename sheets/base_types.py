import re

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

