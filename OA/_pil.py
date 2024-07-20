# -*- coding: utf-8 -*-

from collections import namedtuple
from pathlib import Path
from typing import NamedTuple

import xlwings as xw

from ._autozip import auto_zip
from ._search import TEMPLATE_DIR

# Constants for individual income tax script configuration
TPL_2018: Path = TEMPLATE_DIR.joinpath('Excel/2018.xls')
TPL_2019: Path = TEMPLATE_DIR.joinpath('Excel/2019.xls')


class EXCEL(NamedTuple):
    """Container for Excel template paths."""
    Template_2018 = TPL_2018
    Template_2019 = TPL_2019


# Define a namedtuple for periods
Period = namedtuple('Period', ['Start', 'End'])


@auto_zip
def generate_personal_income_tax(register, periods, output_folder):
    """
    Generate personal income tax for a specified period of time.

    Args:
        register (dict): Registration information.
        periods (list): List of period tuples with start and end dates.
        output_folder (Path): Folder to save output Excel files.

    Returns:
        Path: The output folder path.
    """
    # Launch Excel in the background
    with xw.App(visible=False, add_book=False) as app:
        app.display_alerts = False
        app.screen_updating = False

        for period in map(lambda x: Period._make(x), periods):
            # Merge period information into register
            register |= period._asdict()

            # Determine the template year based on the start year of the period
            anchor = 2018 if (year := period.Start.year) < 2019 else 2019

            # Open the corresponding Excel workbook
            wb = app.books.open(getattr(EXCEL, f'Template_{anchor}'))
            sht = wb.sheets[0]

            # Fill in the Excel sheet with the registration information
            for k, v in register.items():
                sht.range(k).value = v

            # Construct the output file path based on the period
            month = period.Start.month
            output_path = output_folder / f'{year}_{month:02}.xls'

            # Save and close the workbook
            wb.save(output_path)
            wb.close()

    return output_folder
