#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   func_test.py
@Author  :   Billy Zhou
@Time    :   2020/02/25
@Version :   0.1.0
@Desc    :   None
'''


import logging
import numpy as np
import xlwings as xw


def test_func(pre_format=''):
    wb = xw.books.active
    # sht = wb.sheets.active
    selected_range = wb.app.selection
    print(selected_range.raw_value)
    # get value of selected range and turn to list
    selected_range_np = selected_range.options(
        np.array, ndim=2, empty='', numbers=np.str_).value
    data_list = list(selected_range_np.flat)

    print(data_list)


if __name__ == '__main__':
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(name)s %(levelname)s %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S')
    logging.debug('start DEBUG')
    logging.debug('==========================================================')

    from pathlib import Path

    filepath = Path.cwd().joinpath('test_data\\md5')
    testfile = str(filepath.joinpath('md5_test.xlsx'))
    xw.Book(testfile).set_mock_caller()
    test_func()

    logging.debug('==========================================================')
    logging.debug('end DEBUG')
