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
