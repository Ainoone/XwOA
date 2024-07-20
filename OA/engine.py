# -*- coding: utf-8 -*-

from collections import namedtuple
from dataclasses import dataclass
from typing import Dict, Any, Tuple, Iterator

from more_itertools import one

import OA.common as com
from ._pil import generate_personal_income_tax
from ._render import render_docx
from ._search import search_template_file
from ._vat import fill_sheet
from .timeperiod import generate_period_range

# 时期
TimeSeries = namedtuple('TimeSeries', ['Start', 'End', 'Freq'])


@dataclass
class DataPartition:
    """Class to partition dictionary into different data groups."""
    data: Dict[str, Any]

    def _split(self) -> Tuple[Dict[str, Any], TimeSeries, Dict[str, Any]]:
        """Splits data into timeseries and register_data"""
        template = {k: v for k, v in self.data.items() if k == 'Template'}
        timeseries = {k: v for k, v in self.data.items() if k in ('Start', 'End', 'Freq')}
        register = {k: v for k, v in self.data.items() if k not in ('Start', 'End', 'Freq', 'Template')}
        return template, TimeSeries(**timeseries), register

    def __iter__(self) -> Iterator[Dict]:
        return iter(self._split())


class TemplateEngine:

    def __init__(self, input_data, only=False):
        self.template = input_data.setdefault('Template', None)
        self.only = only
        self.data = input_data

    @property
    def template_path(self):
        if self.template is None:
            return None
        elif self.template in ('小规模', '一般纳税人'):
            return search_template_file(name=self.template, suffix='xlsx')
        else:
            return search_template_file(self.template)

    def run(self):

        if not self.only:
            out_path = com.create_result_folder(target_folder_name=self.template)
            match self.data:
                case {'Template': tpl} if tpl is not None:
                    return render_docx(initial_data=self,
                                       path=self.template_path, out_fd=out_path, label='Cloud')

                case _:
                    pass

        else:
            dictionary = one(self)
            template, timeseries, register = DataPartition(dictionary)
            periods = generate_period_range(timeseries=timeseries)

            target = register.get('CN', template)
            out_path = com.create_result_folder(target_folder_name=f'{target!s:.6}_{self.template}')

            match self.data:
                case {'Template': '个税压缩包'}:
                    return generate_personal_income_tax(register=register, periods=periods, output_folder=out_path)

                case {'Template': tpl} if tpl in ('小规模', '一般纳税人'):
                    context = com.merge_range_and_data(time_stamps=periods, data=register)
                    return fill_sheet(path=out_path, fullname=self.template_path, data=context)

                case {'Template': tpl} if tpl != '个税压缩包':
                    context = com.merge_range_and_data(time_stamps=periods, data=register)
                    return render_docx(initial_data=context,
                                       path=self.template_path, out_fd=out_path, label='Tax')

                case _:
                    pass

    def __iter__(self):
        return iter(com.iterdict(self.data, only=self.only))
