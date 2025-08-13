def to_locale_str(a):
    """
    Convert a number to a string with thousands separators.

    Parameters:
    -----------
    a : int or float
        The number to format.

    Returns:
    --------
    str
        The number formatted with commas as thousands separators.

    Example:
    --------
    >>> to_locale_str(1234567)
    '1,234,567'
    """
    return "{:,}".format(int(a))
