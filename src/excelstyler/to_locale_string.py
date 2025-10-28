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

    Raises:
    -------
    ValueError
        If the input cannot be converted to a number.

    Example:
    --------
    >>> to_locale_str(1234567)
    '1,234,567'
    """
    if a is None:
        raise ValueError("Input cannot be None")
    
    try:
        return "{:,}".format(int(a))
    except (ValueError, TypeError) as e:
        raise ValueError(f"Cannot convert '{a}' to a number: {e}")
