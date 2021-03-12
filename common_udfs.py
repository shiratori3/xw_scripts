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

"""
      Maybe you will have problem to import udfs into xlsm file, just as below:

 Error 
 --------------------------- 
 Original error was: DLL load failed while importing _multiarray_umath: 鎵句笉鍒版寚瀹?an existing issue about this - open a new one instead. 
 Note: this error has many possible causes, so please don't comment on 
   (removes all files not under version control) and rebuild numpy. 
 - If you're working with a numpy git repository, try `git clean -xdf` 
      - if you built from source, your compiler versions and ideally a build log 
      - whether or not you have multiple versions of Python installed 
      - your operating system 
      - how you installed numpy 
      - how you installed Python 
      https://github.com/numpy/numpy/issues.  Please include details on: 
   2. If (1) looks fine, you can open a new issue at 
      interfere with the Python and numpy version "1.18.1" you're trying to use. 
      and that you have no directories in your PATH or PYTHONPATH that can 
   1. Check that you expected to use Python3.8 from "D:\Miniconda3\envs\pyexcel\pythonw.exe", 
 - If you have already done that, then: 
 - Try uninstalling and r 
 --------------------------- 
 确定    
 --------------------------- 

      This's a numpy import error with mkl_XXX.dll missing.
      When this happen, add ...\Miniconda3\envs\env_name\Library\bin to system path, and then try to reinstall numpy. 
""" # noqa


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
