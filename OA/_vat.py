# -*- coding: utf-8 -*-

from dataclasses import asdict
from pathlib import Path
from typing import Union

import xlwings as xw

from ._classify_files import categorize_files
from ._convert2pdf import convert_to_pdf
from ._sentence import SmallScale, General


# Decorators for file categorization and PDF conversion
@categorize_files(merge=True)
@convert_to_pdf
def fill_sheet(path: Path, fullname: Union[str, Path], data):
    """Fill the sheet with data.

    Args:
        path: The directory path where the output files will be saved.
        fullname: The full path name of the Excel workbook to be processed.
        data: The data to be filled into the Excel workbook.
    """
    with xw.App(visible=False, add_book=False) as app:  # Launch Excel in the background
        wb = app.books.open(fullname=fullname)  # Open the workbook
        sheet = wb.sheets['主表']  # Access the main worksheet

        # Determine the starting cell based on workbook name
        cell = 'A6' if wb.name == '小规模.xlsx' else 'A5'

        for mapping in data:
            # Choose the taxpayer type based on workbook name
            if wb.name == '小规模.xlsx':  # For small scale taxpayers
                record = asdict(SmallScale(database=mapping))
            elif wb.name == '一般纳税人.xlsx':  # For general taxpayers
                record = asdict(General(database=mapping))

            # Fill in the worksheet with the data
            for key, value in filter(lambda kv: not kv[0].startswith('_Taxpayer'), record.items()):
                sheet.range(key).value = value

            # Update the period of tax payment
            sheet.range(cell).value = f'税款所属期：{mapping["Start"]:%Y年%m月%d日}至{mapping["End"]:%Y年%m月%d日}'

            # Convert the worksheet to PDF and save
            sheet.to_pdf(path=path / f'{mapping["End"]:%y_%m%d}.pdf')

            # Save the workbook
            wb.save(path=path / f'{mapping["End"]:%y_%m%d}.xlsx')

        # 关闭工作簿
        wb.close()  # Close the workbook

    return path
