#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   RangeHandler.py
@Author  :   Billy Zhou
@Time    :   2021/08/26
@Desc    :   None
'''


import sys
from pathlib import Path
cwdPath = Path(__file__).parents[2]
sys.path.append(str(cwdPath))

from src.manager.LogManager import logmgr
log = logmgr.get_logger(__name__)

import numpy as np
import xlwings as xw
from xlwings import Range
from decimal import Decimal
from src.utils.cell import num2alpha


class BaseRangeHandler(object):
    """The base class for RangeHandler

    Args:
    ----
        rang: str or Range, default None
            the address of range to handler.
            If none, use the xw.apps.active.selection
        pre_formatter: str
            whether to format the cells in rang before handling
    """
    def __init__(self, rang: str or Range = None, pre_formatter: str = '') -> None:
        """get range under wb and sht"""
        if not rang:
            self.rang = xw.apps.active.selection
        else:
            if isinstance(rang, Range):
                self.rang = rang
            elif isinstance(rang, str):
                self.rang = xw.Range(cell1='', cell2='')

        self.sht = self.rang.sheet
        self.wb = self.sht.book
        self.result_list = []

        if pre_formatter:
            self.rang.number_format = pre_formatter

        self.to_np_list()

    def to_np_list(self) -> None:
        """Convert the value in self.rang to a list of numpy.flatiter instance"""
        self.rang_np_list = list(
            self.rang.options(np.array, ndim=2, empty='').value.flat)
        log.debug('self.rang_np_list: {}'.format(self.rang_np_list))

    def cell_handler(self):
        """rewrite this function for your own macro"""
        pass

    def set_value(self, long_num_to_text: bool = True) -> None:
        """return the result to inputed range

        Args:
        ----
            long_num_to_text: bool, default True
                for nums which length more than 17, turn the format of their cells to text
        """
        if long_num_to_text:
            text_cells_range = []
            for no, value in enumerate(self.result_list, 1):
                # add cell address to text_cells_range if too long
                log.debug("type of value in self.result_list: {}".format(type(value)))
                try:
                    to_text = False
                    if isinstance(value, Decimal):
                        if len(value.to_eng_string()) > 17:
                            to_text = True
                    elif not isinstance(value, np.float64):
                        if len(value) > 15:
                            to_text = True
                    if to_text:
                        row, col = no // self.rang.shape[1] - 1, no % self.rang.shape[1]
                        cell_range = "".join([num2alpha(
                            self.rang.column + col), str(self.rang.row + row)])
                        text_cells_range.append(cell_range)
                except Exception as e:
                    log.error('An error occurred. {}'.format(e.args))
                    log.error('The value of cell: {}'.format(value))
                    log.error('The type of cell value: {}'.format(value))
            log.debug("cells to text: {}".format(text_cells_range))
            for text_cell in text_cells_range:
                self.sht.range(text_cell).number_format = "@"

        log.debug("self.result_list: {}".format(self.result_list))
        self.sht.range(self.rang.address).value = np.array(
            self.result_list).reshape(self.rang.shape[0], self.rang.shape[1])


if __name__ == '__main__':
    pass
