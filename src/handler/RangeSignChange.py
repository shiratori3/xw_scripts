#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   RangeSignChange.py
@Author  :   Billy Zhou
@Time    :   2021/08/30
@Desc    :   None
'''


import sys
from pathlib import Path
cwdPath = Path(__file__).parents[2]  # the num depend on your filepath
sys.path.append(str(cwdPath))

from src.manager.LogManager import logmgr
log = logmgr.get_logger(__name__)

import numpy as np
from typing import Any
from xlwings import Range
from decimal import Decimal
from src.basic.cell_func import str2num
from src.handler.BaseCellHandler import break_where
from src.handler.BaseCellHandler import BaseCellHandler
from src.handler.BaseRangeHandler import BaseRangeHandler


class CellSignChange(BaseCellHandler):
    """change the sign of a numberic cell

    Methods:
    -------
        in_supported_type:
            check whether can change the sign of cell value directly
        not_only_backslash:
            return break while value equal to \\n
        value2deciaml:
            try to convert the value of cell to Decimal
    """
    def __init__(self, cell_value: Any, cell_format: str = ''):
        super().__init__(cell_value, cell_format=cell_format)

    @break_where
    def in_supported_type(self, supported_type: tuple = (np.str_, str)):
        """check whether can change the sign of cell value directly"""
        if self.breaked:
            return self
        if isinstance(self.cell_value, supported_type):
            log.debug('the type of cell value: {}'.format(
                type(self.cell_value)))
            return self
        else:
            # not str or np.str_, change the sign of cell value
            self.cell_value *= -1
            self.breaked = True
            return self

    @break_where
    def value2deciaml(self):
        """try to convert the value to deciaml"""
        if self.breaked:
            return self
        self.cell_value = str2num(self.cell_value)
        if isinstance(self.cell_value, Decimal):
            self.cell_value *= -1
            self.cell_value = self.cell_value.to_eng_string()
            return self
        else:
            self.breaked = True
            return self


class RangeSignChange(BaseRangeHandler):
    """change the sign of numberic cells in a selected range

    Args:
    ----
        rang: str or Range, default None
            the address of range to handler.
            If none, use the xw.apps.active.selection
        pre_formatter: str
            whether to format the cells in rang before handling
    """
    def __init__(self, rang: str or Range = '', pre_formatter: str = '') -> None:
        super().__init__(rang=rang, pre_formatter=pre_formatter)

    def cell_handler(self):
        """try to change the sign for each cell in a selected range"""
        for cell_value in self.rang_np_list:
            cell = CellSignChange(cell_value)
            cell.in_supported_type().not_only_backslash().value2deciaml()
            self.result_list.append(cell.cell_value)


if __name__ == '__main__':
    # test for CellSignChange
    if True:
        cell1 = CellSignChange(100)
        cell1.in_supported_type().not_only_backslash().value2deciaml()

        cell2 = CellSignChange('\n')
        cell2.in_supported_type().not_only_backslash().value2deciaml()

        cell3 = CellSignChange('100')
        cell3.in_supported_type().not_only_backslash().value2deciaml()

        cell4 = CellSignChange('100\n')
        cell4.in_supported_type().not_only_backslash().value2deciaml()
