# -*- coding: utf-8 -*-

from datetime import date
from itertools import combinations
from typing import Iterable, Tuple

import pandas as pd

# Pandas: Version 2.2
# Deprecate aliases M, Q, Y, etc. in favour of ME, QE, YE, etc. for offsets
# Constants for frequency aliases
OFFSET_ALIASES = {
    'YE': 'YearEnd_year end frequency',
    'YS': 'YearBegin_year start frequency',
    'ME': 'MonthEnd_month end frequency',
    'MS': 'MonthBegin_month start frequency',
    'QE': 'QuarterEnd_quarter end frequency',
    'QS': 'QuarterBegin_quarter start frequency',
}

NO_FREQ = "N"  # Indicator for no frequency


def generate_range(start: date | None = None, end: date | None = None,
                   freq: str = "M") -> Iterable[Tuple[date, date]]:
    """
    Generate a range of dates based on the start and end dates, and the frequency provided.

    Parameters:
    - start: Optional start date for the range.
    - end: Optional end date for the range.
    - freq: Frequency for generating dates, defaults to monthly ('M').

    Returns:
    An iterable of tuples (start_date, end_date) for each interval within the range.
    """
    # Validate input arguments
    if freq is None and any(x is None for x in [start, end]):
        raise ValueError("Must provide freq argument if no data is supplied")
    if start is None and end is None:
        raise ValueError("Must provide either start or end argument")

    # Convert start and end dates to period start and end times
    start_time = pd.Timestamp(start).to_period(freq=freq).start_time
    end_time = pd.Timestamp(end).to_period(freq=freq).end_time

    # Define start and end aliases based on frequency
    alias_start, alias_end = f'{freq}S', f'{freq}E'

    # Generate date ranges for start and end times
    date_range_start = pd.date_range(start=start_time, end=end_time, freq=alias_start)
    date_range_end = pd.date_range(start=start_time, end=end_time, freq=alias_end)

    # Return a zip object combining start and end dates
    return zip(date_range_start, date_range_end)


def generate_period_range(timeseries):
    """
    Generate a range of periods within a given time range.

    Parameters:
    - timeseries: A pandas Timestamp object representing the time range.

    Returns:
    An iterable of tuples representing the periods within the time range.
    """
    # Handle no frequency or single point in time case
    if timeseries.Freq == NO_FREQ or timeseries.Start == timeseries.End:
        return combinations((timeseries.Start, timeseries.End), r=2)
    else:
        # Generate range using start, end, and frequency for other cases
        return generate_range(start=timeseries.Start, end=timeseries.End, freq=timeseries.Freq)
