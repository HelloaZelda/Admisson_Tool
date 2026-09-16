def sanitize_excel_input(value):
    """
    Sanitizes user input to prevent CSV/Formula Injection in Excel exports.
    If the value starts with '=', '+', '-', '@', '\t', '\r', or '\n',
    it prepends a single quote to force Excel to treat it as text.
    """
    if value is None:
        return value

    str_value = str(value)
    if str_value and str_value[0] in ('=', '+', '-', '@', '\t', '\r', '\n'):
        return f"'{str_value}"

    return value
