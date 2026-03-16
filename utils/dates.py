import datetime as dt
from bizdays import Calendar


cal = Calendar.load("ANBIMA")


def get_reference_date(days_ago: int) -> dt.date:
    """
    Get the reference date from today.
    """
    today = dt.date.today()
    reference_date = cal.offset(today, -days_ago)
    return reference_date

def get_dates_range(start_date: dt.date, end_date: dt.date) -> list:
    """
    Get a list of business days between two dates.
    """
    return cal.seq(start_date, end_date)