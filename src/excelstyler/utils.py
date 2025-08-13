from datetime import datetime

import jdatetime


def shamsi_date(date, in_value=None):
    if in_value:
        sh_date = jdatetime.date.fromgregorian(
            year=date.year,
            month=date.month,
            day=date.day
        )
    else:
        gh_date = jdatetime.date.fromgregorian(
            year=date.year,
            month=date.month,
            day=date.day
        ).strftime('%Y-%m-%d')
        reversed_date = reversed(gh_date.split("-"))
        separate = "-"
        sh_date = separate.join(reversed_date)
    return sh_date


def convert_str_to_date(string):
    string = str(string).strip()
    try:
        return datetime.strptime(string, '%Y-%m-%dT%H:%M:%S.%fZ').date()
    except ValueError:
        try:
            return datetime.strptime(string, '%Y-%m-%dT%H:%M:%SZ').date()  # Added format without milliseconds
        except ValueError:
            try:
                return datetime.strptime(string, '%Y-%m-%d').date()
            except ValueError:
                return None
