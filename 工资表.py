# -*- coding: utf-8 -*-

from datetime import datetime
from typing import NamedTuple

import pandas as pd
import xlwings as xw


class SheetName(NamedTuple):
    task = '任务'  # Task
    perf = '绩效'  # Performance
    stats = '统计'  # Statistic

    def __str__(self):
        return "工作表的名字"


# Parameters --> 参数
class OptionsParams(NamedTuple):
    convert: pd.DataFrame = pd.DataFrame
    index: bool = False
    head: int = 1

    def __str__(self):
        return f'Pandas 序列转换器参数'


class DataColumns(NamedTuple):
    task = ['产品名称', '客户名称', '负责人', '协作人', 'Date', '服务金额']  # 任务工作表
    label = ['产品名称', '客户名称', '负责人', '协作人', 'Date', '服务金额', '提成基数', '预计提成']


class GroupbyParams(NamedTuple):
    by: list = ['产品名称', '客户名称']
    axis: int = 0
    as_index: bool = False  # 分组字典不设为索引
    sort: bool = False


class MergeParams(NamedTuple):
    how: str = 'outer'  # 并集
    on: list = ['产品名称', '客户名称']

    def __str__(self):
        return 'pandas.merge 参数'


class SortValuesParams(NamedTuple):
    by: list = ['Date']  # 任务工作表排序 by 参数
    axis: int = 0  #
    ascending: bool = True  # 升序
    ignore_index: bool = True  # 索引重新 0-(n-1)排序

    def __str__(self):
        return f'数值排序参数'


def str_to_date(df, label):
    """DataFrame 指定列 --> 字符串转换为 datetime"""
    return (result := df.convert_dtypes()).assign(
        Date=result.loc[:, label]
        .apply(lambda x: datetime.strptime(x, '%Y-%m-%d %H:%M:%S').date()))


def fetch_data(book, sheet_name):
    sheet = book.sheets[sheet_name]
    data_value = sheet.used_range.options(**OptionsParams._field_defaults).value
    return data_value.convert_dtypes()


def main():
    book = xw.Book.caller()
    # 任务工作表
    task_data_frame = fetch_data(book, SheetName.task).pipe(str_to_date,
                                                            '完成时间').loc[:, DataColumns.task]
    # 绩效工作表
    performance_data_frame = (fetch_data(book, SheetName.perf).pipe(pd.DataFrame.groupby,
                                                                    **GroupbyParams._field_defaults).agg
                              ({'预计提成': 'sum', '提成基数': 'mean'}))
    # 合并后的结果
    result = pd.merge(left=task_data_frame, right=performance_data_frame,
                      **MergeParams._field_defaults)
    # 排序并指定列值
    sorted_df = result.sort_values(**SortValuesParams._field_defaults).loc[:, DataColumns.label]
    # 输出
    stats_sht = book.sheets.add(SheetName.stats)
    table_name = 'Stats'  # 表的名称
    stats_sht.tables.add(source=stats_sht['A1'], name=table_name).update(sorted_df, index=False)
    stats_sht.autofit()


if __name__ == '__main__':
    pass
