# -*- coding: utf-8 -*-

import xlwings as xw

from OA import Worksheet, TemplateEngine


def main():
    book = xw.Book.caller()
    sheet = book.sheets.active  # 当前工作表
    # 数据字典
    target_data = Worksheet(sheet).data
    # 引擎
    engine = TemplateEngine(target_data, only=True)
    # 自动化
    engine.run()


if __name__ == '__main__':
    pass
