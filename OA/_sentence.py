# -*- coding: utf-8 -*-

from dataclasses import dataclass, InitVar, field
from string import Template
from typing import Any


@dataclass
class Taxpayer:
    """纳税人基类 Taxpayer base class."""
    database: InitVar[Any] = None
    __name: str = field(init=False)
    __code: str = field(init=False)
    __date: str = field(init=False)
    __agent: str = field(init=False)
    __agent_id: str = field(init=False)

    def __post_init__(self, database):
        if database:
            self.__name = Template('纳税人名称:$CN').substitute(CN=database['CN'])  # 纳税人名称
            self.__code = Template('纳税人识别号(统一社会信用代码):$CC').substitute(CC=database['CC'])  # 纳税人识别号
            self.__date = Template('$Date').substitute(Date=database['Date'])  # 填报日期
            self.__agent = Template('经办人:$Ag').substitute(Ag=database['Name'])  # 经办人
            self.__agent_id = Template('经办人身份证号码:$AID').substitute(AID=database['IDN'])  # 经办人身份证号码


@dataclass
class SmallScale(Taxpayer):
    """小规模纳税人 Small-scale taxpayer."""
    A4: str = field(init=False)
    A5: str = field(init=False)
    G6: str = field(init=False)
    G40: str = field(init=False)
    A41: str = field(init=False)
    A42: str = field(init=False)

    def __post_init__(self, database):
        super().__post_init__(database)
        self.A4 = self._Taxpayer__name  # 纳税人名称
        self.A5 = self._Taxpayer__code  # 纳税人识别号
        self.G6 = self._Taxpayer__date  # 填报日期
        self.G40 = self._Taxpayer__date  # 填报日期
        self.A41 = self._Taxpayer__agent  # 经办人
        self.A42 = self._Taxpayer__agent_id  # 经办人身份证号码


@dataclass
class General(Taxpayer):
    """一般纳税人 General taxpayer"""
    A6: str = field(init=False)
    A7: str = field(init=False)
    U5: str = field(init=False)
    AK54: str = field(init=False)
    A55: str = field(init=False)
    A56: str = field(init=False)

    def __post_init__(self, database):
        super().__post_init__(database)
        self.A7 = self._Taxpayer__name  # 纳税人名称
        self.A6 = self._Taxpayer__code  # 纳税人识别号
        self.U5 = self._Taxpayer__date  # 填报日期
        self.AK54 = self._Taxpayer__date  # 填报日期
        self.A55 = self._Taxpayer__agent  # 经办人
        self.A56 = self._Taxpayer__agent_id  # 经办人身份证号码
