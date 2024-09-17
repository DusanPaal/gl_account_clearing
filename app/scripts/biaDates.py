"""
The 'biaDates.py' module provides procedures that perform
all date-related calculations of the application.

Version histry:
    1.0.20220429 - Initial version.
    1.0.20220615 - Refactored and simplified code, added/updated docstrings.
"""

from datetime import date, datetime, timedelta
import numpy as np

def _is_ultimo_plus_one(off_work_days: list) -> bool:
    """
    Checks wheher a current date is Ultimo + 1 date.

    Params:
        off_work_days: List of out of office dates (holidays, exceptional situations, etc ...).

    Returns: True if the current date is Ultimo + 1 date, otherwise False.
    """

    curr_date = get_date(0)
    first_day_month = start_of_month(curr_date)
    first_workday = first_day_month

    while not np.is_busday(first_workday, holidays = off_work_days):
        first_workday += timedelta(1)

    if first_workday == curr_date:
        return True

    return False

def calculate_fiscal_times(off_work_days: list) -> date:
    """
    Calculates fiscal year, period and clearing date of the previous calendar month.

    Params:
        off_work_days: List of out of office dates (holidays, exceptional situations, etc ...).

    Returns: Calculated clearing date.
    """

    curr_date = get_date(0)

    if not _is_ultimo_plus_one(off_work_days):
        clearing_date = curr_date
    else:

        last_day_prev_mon = start_of_month(curr_date) - timedelta(1)

        while not np.is_busday(last_day_prev_mon, holidays = off_work_days):
            last_day_prev_mon -= timedelta(1)

        clearing_date = last_day_prev_mon

    return clearing_date

def start_of_month(date: datetime) -> datetime:
    """
    Calculates first day of a month.

    Params:
        date: Date of the month for which the first day is calculated.

    Returns: First date of the month in the datetime format.
    """

    first_day = date.replace(day=1)

    return first_day

def get_date(day_offset: int = 0, weeks_offset: int = 0) -> datetime:
    """
    Calculates a date by adding an offset to the current date.

    Params:
        day_offset: Offset in days compared to the current date.
        weeks_offset: Offset in weeks compared to the current date.

    Returns: Offsetted date in the datetime format.
    """

    assert not(day_offset != 0 and weeks_offset != 0)

    if day_offset != 0:
        offset = datetime.date(datetime.now()) + timedelta(days=day_offset)
    elif weeks_offset != 0:
        offset = datetime.date(datetime.now()) + timedelta(weeks=weeks_offset)
    else:
        offset = datetime.date(datetime.now())

    return offset
