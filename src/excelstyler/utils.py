from datetime import datetime

import jdatetime


def shamsi_date(date, in_value=None):
    """
    Convert a Gregorian date to a Shamsi (Persian) date.

    Parameters:
    -----------
    date : datetime.date or datetime.datetime
        The Gregorian date to convert.

    in_value : bool, optional
        - If True:
            Returns a `jdatetime.date` object.
            Suitable for storing as a value inside Excel cells.
        - If False or None:
            Returns a string in 'DD-MM-YYYY' format (reversed from 'YYYY-MM-DD').
            Suitable for displaying in reports or Excel sheets.

    Returns:
    --------
    jdatetime.date or str
        The Shamsi date, either as an object or a string depending on `in_value`.

    Raises:
    -------
    ValueError
        If the input date is invalid or None.

    Note:
    -----
    Setting `in_value=True` is useful when you want to write the date directly
    into Excel cell values, while `in_value=False` is for formatted text display.
    """
    if date is None:
        raise ValueError("Date cannot be None")
    
    try:
        if in_value:
            sh_date = jdatetime.date.fromgregorian(
                year=date.year,
                month=date.month,
                day=date.day
            )
        else:
            miladi_date = jdatetime.date.fromgregorian(
                year=date.year,
                month=date.month,
                day=date.day
            ).strftime('%Y-%m-%d')
            reversed_date = reversed(miladi_date.split("-"))
            separate = "-"
            sh_date = separate.join(reversed_date)
        return sh_date
    except (ValueError, AttributeError) as e:
        raise ValueError(f"Invalid date format: {e}")


def convert_str_to_date(string):
    """
    Convert a string to a datetime.date object.

    This function tries multiple common date formats, including ISO 8601 with or
    without milliseconds, and plain 'YYYY-MM-DD'. If the string cannot be parsed,
    it returns None.

    Parameters:
    -----------
    string : str
        The date string to convert.

    Returns:
    --------
    datetime.date or None
        A datetime.date object if conversion succeeds, otherwise None.

    Supported formats:
    ------------------
    - 'YYYY-MM-DDTHH:MM:SS.sssZ'  (ISO 8601 with milliseconds)
    - 'YYYY-MM-DDTHH:MM:SSZ'      (ISO 8601 without milliseconds)
    - 'YYYY-MM-DD'                 (Simple date)
    """
    string = str(string).strip()
    try:
        return datetime.strptime(string, '%Y-%m-%dT%H:%M:%S.%fZ').date()
    except ValueError:
        try:
            return datetime.strptime(string, '%Y-%m-%dT%H:%M:%SZ').date()
        except ValueError:
            try:
                return datetime.strptime(string, '%Y-%m-%d').date()
            except ValueError:
                return None
