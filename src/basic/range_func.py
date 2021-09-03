#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   range_func.py
@Author  :   Billy Zhou
@Time    :   2020/08/26
@Desc    :   None
'''


import sys
from pathlib import Path
cwdPath = Path(__file__).parents[2]
sys.path.append(str(cwdPath))

from src.manager.LogManager import logmgr  # noqa: E402
log = logmgr.get_logger(__name__)

from xlwings import Range
from src.basic.app_func import try_expect_stop_refresh
from src.handler.RangeSignChange import RangeSignChange
from src.handler.RangeNumClean import RangeNumClean


@try_expect_stop_refresh
def change_sign(rang: str or Range = '', pre_formatter: str = ''):
    """change the sign of numberic cells in a selected range

    Args:
    ----
        rang: str or Range, default None
            the address of range to handler.
            If none, use the xw.apps.active.selection
        pre_formatter: str
            whether to format the cells in rang before handling
    """
    rang_to_change = RangeSignChange(rang=rang, pre_formatter=pre_formatter)
    rang_to_change.cell_handler()
    rang_to_change.set_value()


@try_expect_stop_refresh
def clean_num(rang: str or Range = '', pre_formatter: str = '', point_digit: int = 2, truncate: bool = False):
    """change the sign of numberic cells in a selected range

    Args:
    ----
        rang: str or Range, default None
            the address of range to handler.
            If none, use the xw.apps.active.selection
        pre_formatter: str
            whether to format the cells in rang before handling
    """
    rang_to_change = RangeNumClean(rang=rang, pre_formatter=pre_formatter, point_digit=point_digit, truncate=truncate)
    rang_to_change.cell_handler()
    rang_to_change.set_value()


if __name__ == '__main__':
    import xlwings as xw

    # test for change_sign
    if False:
        filepath = cwdPath.joinpath('res\\dev\\change_sign')
        testfile = str(filepath.joinpath('test_change_sign.xlsx'))
        xw.Book(testfile).set_mock_caller()
        change_sign()

    # test for clean_num
    if True:
        filepath = cwdPath.joinpath('res\\dev\\num_clean')
        testfile_clean = filepath.joinpath('test_num_clean.xlsx')
        xw.Book(testfile_clean).set_mock_caller()
        if False:
            clean_num(point_digit=2, truncate=False)
        if False:
            clean_num(point_digit=4, truncate=True)
        if True:
            clean_num(point_digit=0, truncate=True)
