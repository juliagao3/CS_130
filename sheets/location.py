def from_base_26(s: str):
    result = 0
    for c in s:
        result *= 26
        result += ord(c) - ord('a') + 1
    return result

def location_split(location: str):
    alpha_end = 0
    while alpha_end < len(location) and location[alpha_end].isalpha():
        alpha_end += 1

    if alpha_end == 0 or alpha_end == len(location):
        raise ValueError

    col = location[:alpha_end]
    row = location[alpha_end:]

    return (col, row)

def location_string_to_tuple(location: str):
    col, row = location_split(location)
    return (from_base_26(col), int(row))

def check_location(location: str):
        """
        check that the given location
        - has the form [row][col]
        - is to the left and above ZZZZ9999

        ValueError is raised if and only if the location is invalid
        """
        location = location.lower()
        col, row = location_string_to_tuple(location)
        if col > from_base_26("zzzz") or row > 9999:
            raise ValueError
        return location