#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   sheet_udf.py
@Author  :   Billy Zhou
@Time    :   2021/08/26
@Desc    :   None
'''


import sys
from pathlib import Path
cwdPath = Path(__file__).parents[2]  # the num depend on your filepath
sys.path.append(str(cwdPath))

from src.manager.LogManager import logmgr
log = logmgr.get_logger(__name__)

import xlwings as xw
from src.basic.range_func import change_sign
from src.basic.range_func import clean_num


@xw.sub
def change_selected_sign():
    """change the sign of cells in selected range"""
    change_sign()


@xw.sub
def change_selected_sign_preformat_text():
    """change the format of cells to text and change the sign of cells in selected range"""
    change_sign(pre_formatter="@")


@xw.sub
def clean_num_with2dec_round():
    """clean the cells in selected range to nums with 2 decimal point, round if more than 2"""
    clean_num(pre_formatter="", point_digit=2, truncate=False)


@xw.sub
def clean_num_with2dec_truncate():
    """clean the cells in selected range to nums with 2 decimal point, truncate if more than 2"""
    clean_num(pre_formatter="", point_digit=2, truncate=True)


@xw.sub
def clean_num_with4dec_round():
    """clean the cells in selected range to nums with 4 decimal point, round if more than 4"""
    clean_num(pre_formatter="", point_digit=4, truncate=False)


@xw.sub
def clean_num_with0dec_truncate():
    """clean the cells in selected range to nums with no decimal point, truncate if more than 0"""
    clean_num(pre_formatter="", point_digit=0, truncate=True)


if __name__ == '__main__':
    if False:
        filepath = cwdPath.joinpath('res\\dev\\change_sign')
        testfile_sum = filepath.joinpath('test_change_sign.xlsx')
        xw.Book(testfile_sum).set_mock_caller()
        change_selected_sign_preformat_text()
