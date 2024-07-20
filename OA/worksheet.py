# -*- coding: utf-8 -*-

import logging
from collections import namedtuple, UserDict
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Dict, Any, Set, List, Iterable, NamedTuple

from xlwings import Sheet

# Set up basic configuration for logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# NamedTuple to represent parts of a named range in Excel
Sandwich = namedtuple('Sandwich', ['prefix', 'sep', 'suffix'])


@dataclass
class DefinedName:
    """
    A class for handling named ranges in an Excel worksheet.
    """
    spe = '!'  # Separator character used in Excel named ranges --> '!' 类属性
    sheet: Sheet  # The Excel worksheet object

    def __post_init__(self):
        # Cache for named ranges to avoid repeated access
        self._names_cache: List[Any] | None = None  # 将 _names_cache 设置为实例属性

        if self._names_cache is None:
            # sheet.names -->返回所有名字与本工作表有关的命名区域的集合
            # (名字定义中包含”SheetName!” (工作表名!)前缀)。
            self._names_cache = self.sheet.names

    def __iter__(self) -> Iterable[str]:

        # len(names) Raise error if no named ranges are found
        if not self._names_cache:
            logging.error("Sheet does not have a named range.")
            raise TypeError("Sheet does not have a named range.")

        #  Yield the suffix part of each named range
        for item in self._names_cache:
            # ①:item.name --> 返回命名对象的名称。
            # ②:str.partition(sep) -->在 sep 首次出现的位置拆分字符串，
            # 返回一个3元组:其中包含分隔符之前的部分、分隔符本身，以及分隔符之后的部分。
            yield Sandwich._make(item.name.partition(self.spe)).suffix


# 定义一个名为DefaultValues的元组，用于存储默认值
class DefaultValues(NamedTuple):
    """
    NamedTuple to store default values for various fields.
    """
    Start: datetime = None
    End: datetime = None
    Date: datetime = None
    Freq: str = 'M'
    Phone: str = " " * 11  # 静态默认值
    IDN: str = " " * 18  # 静态默认值

    def __str__(self) -> str:
        return f'Default values dictionary'

    @staticmethod
    def current_date():
        """Returns the current date and time."""
        return datetime.now()


class NamedRangeDict(UserDict):
    """
    A UserDict subclass that has a NamedTuple for default values.
    """

    def __init__(self, data_dict=None, **kwargs):
        self.default_values_instance = DefaultValues()
        self.default_values = self.default_values_instance._asdict()
        super().__init__(data_dict, **kwargs)

    def check_and_set_values(self, key):
        """ Checks if the given key has a value.
        If not, sets it to its default """
        if self.data[key] is None:
            self.set_default(key)
            self.data[key] = self.default_values.setdefault(key, None)

    def set_default(self, _key):
        """ Sets the default value for the given key """
        date_fields = {'Start', 'End', 'Date'}
        if _key in date_fields:
            self.default_values[_key] = (self.default_values_instance.
                                         current_date())
        elif _key == 'phone':
            self.default_values[_key] = self.default_values_instance.Phone
        elif _key == 'id_num':
            self.default_values[_key] = self.default_values_instance.IDN

    def __setitem__(self, key, value):
        super().__setitem__(key, value)
        self.check_and_set_values(key)

    def get_clean_data(self):
        """Returns a dictionary copy without internal attributes."""
        return dict(self.data)

# ----------------------------------------------------------------------


@dataclass
class Worksheet:
    sheets: Sheet | list
    sheet_names: Set[str] = field(default_factory=set)

    def __post_init__(self):
        # Ensure sheets is a list for uniform processing
        self.sheets = [self.sheets] if isinstance(self.sheets, Sheet) else self.sheets
        if self.sheets is None:
            raise ValueError('No worksheets are provided.')

    @staticmethod
    def convert_to_dict(sheet: Sheet) -> Dict[str, Any]:
        """
        Converts named ranges in a sheet to a dictionary.
        """
        # 将命名区域的名称作为 key，命名区域的值作为 value，构建字典。
        # 通过 Range 的 options() 方法，可以指定读取数据的格式。
        # options() 方法的参数 dates 用于指定日期格式:将日期转换为 datetime.date 类型。
        # empty 用于指定空值的格式：将空值转换为 None。
        return {name: sheet.range(name).options(dates=date, empty=None).value
                for name in DefinedName(sheet)}

    def check_duplicate_sheets_and_convert(self, sheet: Sheet) -> Dict[str, Any]:
        """
        Check if there are any duplicate sheet names and convert the sheet data to a dictionary.

        :param sheet: The sheet object to check for duplicates and convert.
        :return: A dictionary representation of the sheet data.

        Raises:
            ValueError: If a duplicate sheet name is detected.
        """
        if sheet.name in self.sheet_names:
            message = f'Duplicate sheet names detected: {sheet.name}'
            logging.error(message)
            raise ValueError(message)
        else:
            # 将工作表名称添加到集合中
            self.sheet_names.add(sheet.name)
            # 将工作表数据转换为字典
            return self.convert_to_dict(sheet)

    @property
    def data(self) -> dict:
        """
        Retrieves data from sheets and returns it as a dictionary.

        :return: A NamedRangeDict object containing the data from the sheets.
        """
        data = dict()

        for sheet in self.sheets:
            try:
                data |= self.check_duplicate_sheets_and_convert(sheet=sheet)
            except Exception as e:
                logging.error(f"Handling worksheet error: {sheet.name}", e)
        # 使用 get_clean_data 方法来返回不包含特定属性的字典
        return NamedRangeDict(data).get_clean_data()

    def __iter__(self):
        yield from self.data

    def __len__(self):
        return len(self.sheets)


if __name__ == '__main__':
    pass
