#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   common_udfs.py
@Author  :   Billy Zhou
@Time    :   2021/03/03
@Version :   1.2.0
@Desc    :   None
'''

import logging
import xlwings as xw
from main import stop_refresh
from main import start_refresh
from main import num_clean
from main import to_csv
from main import sum_sheets
from main import change_minus


@xw.sub
def merge_current_sheet():
    stop_refresh()
    sum_sheets(1, current_book=True)
    start_refresh()


@xw.sub
def merge_csv_to_same_sheet():
    stop_refresh()
    sum_sheets(0, f_type=".csv")
    start_refresh()


@xw.sub
def merge_xlsx_to_same_sheet():
    stop_refresh()
    sum_sheets(0, f_type=".xlsx")
    start_refresh()


@xw.sub
def merge_csv_to_diff_sheets():
    stop_refresh()
    sum_sheets(
        ignore_row_num=0, f_type=".csv", current_book=False,
        copy_first_sheet=True,
        target_sheet1=False, target_same_sheetname=False
    )
    start_refresh()


@xw.sub
def merge_xls_to_diff_sheets():
    stop_refresh()
    sum_sheets(
        ignore_row_num=0, f_type=".xls", current_book=False,
        copy_first_sheet=True,
        target_sheet1=False, target_same_sheetname=False
    )
    start_refresh()


@xw.sub
def merge_xlsx_to_diff_sheets():
    stop_refresh()
    sum_sheets(
        ignore_row_num=0, f_type=".xlsx", current_book=False,
        copy_first_sheet=True,
        target_sheet1=False, target_same_sheetname=False
    )
    start_refresh()


@xw.sub
def merge_xlsx_to_diff_sheets_with_same_name():
    stop_refresh()
    sum_sheets(
        ignore_row_num=0, f_type=".xlsx", current_book=False,
        copy_first_sheet=False,
        target_sheet1=False, target_same_sheetname=True
    )
    start_refresh()


@xw.sub
def from_xlsx_to_csv():
    to_csv(filetype=".xlsx")


@xw.sub
def from_xls_to_csv():
    to_csv(filetype=".xls")


@xw.sub
def num_with_2_dec():
    stop_refresh()
    num_clean(point_exist=True, point_digit=2, point_overflow_accept=True)
    start_refresh()


@xw.sub
def num_with_4_dec():
    stop_refresh()
    num_clean(point_exist=True, point_digit=4, point_overflow_accept=True)
    start_refresh()


@xw.sub
def num_with_0_dec():
    stop_refresh()
    num_clean(point_exist=False, point_digit=0, point_overflow_accept=True)
    start_refresh()


@xw.sub
def num_change_minus():
    stop_refresh()
    change_minus()
    start_refresh()


if __name__ == '__main__':
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(name)s %(levelname)s %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S')
    logging.debug('start DEBUG')
    logging.debug('==========================================================')

    # cwd = os.getcwd()

    # wb = xw.Book(cwd + '\\regex_test.xlsx')
    # sht = wb.sheets['Sheet1']
    # sht.range('A1').value = 'Foo 1'

    # xw.Book(cwd + '\\regex_test.xlsx').set_mock_caller()
    # regex_testing()
    # from_xlsx_to_csv()

    logging.debug('==========================================================')
    logging.debug('end DEBUG')
