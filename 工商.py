# -*- coding: utf-8 -*-

import xlwings as xw

from OA import Worksheet, TemplateEngine


def main():
    wb = xw.Book.caller()
    ws = Worksheet(wb.sheets.active)
    data = ws.data
    engine = TemplateEngine(data, only=False)
    engine.run()


if __name__ == '__main__':
    pass
