# -*- coding: utf-8 -*-

import re
from collections import namedtuple
from pathlib import Path
from typing import Dict, Any, List, Tuple, Generator

import pandas as pd
from more_itertools import always_iterable, first

# 时间序列
TimeSeries = namedtuple('TimeSeries', ['Start', 'End'])


def iterdict(input_data: dict, only=False) -> List[Dict[str, Any]]:
    """
    Transforms input data into a list of dictionaries suitable for DataFrame conversion.

    Args:
        input_data (dict): The input dictionary to be transformed.
        only (bool, optional): A flag to determine the processing path. Defaults to False.

    Returns:
        List[Dict[str, Any]]: A list of dictionaries representing the processed input data.

    Raises:
        ValueError: If all scalar values are used without an index.
    """
    if only:
        target_data = list(always_iterable(input_data, base_type=dict))
    else:
        target_data = input_data | dict(Template=None)
    try:
        result = pd.DataFrame(target_data)
        df = (result
              # .replace(to_replace=None, value=pd.NA)  # 将 None 值替换为 NaN
              .convert_dtypes()
              .dropna(how='all', axis=0)  # 删除所有值均为 NA 的行
              .dropna(how='all', axis=1)  # 删除所有值均为 NA 的列
              )
    except ValueError as error:
        raise ValueError("If using all scalar values, you must pass an index") from error

    # Convert the DataFrame back to a list of dictionaries and return.
    return df.to_dict(orient='records')


def create_result_folder(top=None, *, target_folder_name='Result') -> Path:
    """
    Creates a folder with the specified name in the given top directory
    or on the Desktop if no directory is provided.

    Parameters:
    - top (str, optional): The path of the top directory where the folder will be created.
      Defaults to None, which means the Desktop will be used.
    - target_folder_name (str, optional): The name of the folder to be created. Defaults to 'Result'.

    Returns:
    - Path: The path of the created folder.

    Raises:
    - ValueError: If the target folder name contains invalid characters.
    """
    # Determine the base path: Desktop if 'top' is None, otherwise use 'top'
    desktop_path = Path.home() / 'Desktop'
    path = desktop_path if top is None else Path(top)

    # 验证目标文件夹名称是否包含非法字符
    # 正则表达式模式字符串r'^[a-zA-Z0-9-_\u4e00-\u9fa5]+$'的含义如下：
    # ^ ：匹配字符串的开头。
    # [a-zA-Z0-9-_\u4e00-\u9fa5]：字符类，表示允许的字符范围。这包括大写字母、小写字母、数字、下划线、（）、破折号和中文。
    # \u4e00 表示中文字符的起始编码， \u9fa5 表示中文字符的结束编码
    # +：匹配前面的字符类至少一次。
    # $：匹配字符串的结尾。
    if not re.match(r'^[a-zA-Z0-9-_（）\u4e00-\u9fa5]+$', target_folder_name):
        raise ValueError("Invalid folder name")

    # Construct the full path to the target folder
    target_path = path / target_folder_name

    # Attempt to create the folder, including any necessary parent directories
    try:
        target_path.mkdir(parents=True, exist_ok=True)
        print(f"Created new folder: {target_path}")
    except OSError as error:
        print(f"Failed to create folder: {error}")

    # Return the path to the newly created folder
    return target_path


def merge_range_and_data(time_stamps: List[Tuple[Any, Any]], data: dict) -> Generator[dict, None, None]:
    """
    Merges a list of time stamp tuples with a data dictionary and yields the merged dictionaries.

    Args:
        time_stamps (List[Tuple[Any, Any]]): A list of tuples representing time periods.
        data (dict): A dictionary containing data that needs to be merged with each time period.

    Yields:
        Generator[dict, None, None]: A generator that yields dictionaries containing the merged data and time periods.
    """
    # Convert each time stamp tuple into a TimeSeries object and then into a dictionary
    timestamp_periods = (TimeSeries(*time_stamp)._asdict() for time_stamp in time_stamps)

    # Iterate through each time period dictionary and merge it with the data dictionary
    for period in timestamp_periods:
        # Merging the data dictionary with the current time period dictionary
        yield {**data, **period}


def _clean_list_values(value: Any) -> List[Any]:
    """
    Cleans the input list or tuple by removing None values.

    Args:
        value (Any): The input value which can be a list, tuple, or any other data type.

    Returns:
        List[Any]: If the input is a list or tuple, it returns a new list with None values removed.
                   If the input is of any other data type, it returns the input as is.
    """
    # Check if the value is a list or tuple
    if isinstance(value, (list, tuple)):
        # return list(strip(value, pred=lambda x: x is None))

        # Return a new list with None values removed
        return [x for x in value if x is not None]
    # Return the input as is if it's not a list or tuple
    return value


def extract_first_value(dictionary):
    for k, v in dictionary.items():
        if isinstance(v, (list, tuple)):
            dictionary[k] = first(v, default=None)
    return dictionary
