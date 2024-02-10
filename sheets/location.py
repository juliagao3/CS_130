def from_base_26(s: str):
    result = 0
    for c in s:
        result *= 26
        result += ord(c) - ord('a') + 1
    return result

def location_split(location: str):
    alpha_end = 0
    is_absolute_col = (location[0] == "$")

    if is_absolute_col:
         location = location[1:]

    is_absolute_row = ("$" in location)
    
    while alpha_end < len(location) and location[alpha_end].isalpha():
        alpha_end += 1
                         
    col = location[:alpha_end]

    if is_absolute_row:
        row = location[alpha_end+1:]
    else:
        row = location[alpha_end:]
              
    if alpha_end == 0 or alpha_end == len(location) or ' ' in location or not row.isnumeric():
        raise KeyError    

    return (col, row, is_absolute_col, is_absolute_row)

def location_string_to_tuple(location: str):
    col, row, is_absolute_col, is_absolute_row = location_split(location)
    return (from_base_26(col), int(row), is_absolute_col, is_absolute_row)

def check_location(location: str):
        """
        check that the given location
        - has the form [row][col]
        - is to the left and above ZZZZ9999

        ValueError is raised if and only if the location is invalid
        """
        location = location.lower()
        col, row, is_absolute_col, is_absolute_row = location_string_to_tuple(location)
        if col > from_base_26("zzzz") or row > 9999:
            raise ValueError
        
        return location
