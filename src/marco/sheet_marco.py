#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   sheet_marco.py
@Author  :   Billy Zhou
@Time    :   2021/09/03
@Desc    :   None
'''


import sys
from pathlib import Path
cwdPath = Path(__file__).parents[2]  # the num depend on your filepath
sys.path.append(str(cwdPath))

from src.manager.LogManager import logmgr
log = logmgr.get_logger(__name__)

import xlwings as xw
from src.basic.sheet_func import sum_sheets


@xw.sub
def merge_cur_wb_to_cur_sheet_NoRow1():
    """merge sheets to current sheet in current workbook without first row"""
    sum_sheets(1, sum_current=True, target_cur_sht=True)


@xw.sub
def merge_cur_wb_to_first_sheet_NoRow1():
    """merge sheets to Sheet1 or first sheet in current workbook without first row"""
    sum_sheets(1, sum_current=True, target_cur_sht=False)


@xw.sub
def merge_csv_to_cur_sheet():
    """merge files with csv suffix to current sheet in current workbook"""
    sum_sheets(0, suffix=".csv", target_cur_sht=True)


@xw.sub
def merge_csv_to_diff_sheets():
    """merge files with csv suffix to diff target sheets named their sheetname in current workbook"""
    sum_sheets(0, suffix=".csv", target_cur_sht=False, target_same_sheetname=True)


@xw.sub
def merge_xls_first_sheet_to_cur_sheet_NoRow1():
    """merge the first sheet in files with xls suffix to current sheet in current workbook without first row"""
    sum_sheets(1, suffix=".xls", only_copy_first_sheet=True, target_cur_sht=True)


@xw.sub
def merge_xlsx_first_sheet_to_cur_sheet_NoRow1():
    """merge the first sheet in files with xlsx suffix to current sheet in current workbook without first row"""
    sum_sheets(1, suffix=".xlsx", only_copy_first_sheet=True, target_cur_sht=True)


@xw.sub
def merge_xlsx_first_sheet_to_diff_sheets_with_sheetname():
    """merge the first sheet in files with xlsx suffix to diff target sheets named their sheetname in current workbook"""
    sum_sheets(
        0, suffix=".xlsx", only_copy_first_sheet=True,
        target_cur_sht=False, target_same_sheetname=True
    )


@xw.sub
def merge_xlsx_first_sheet_to_diff_sheets_with_bookname():
    """merge the first sheet in files with xlsx suffix to target sheets named their workbook name in current workbook"""
    sum_sheets(
        0, suffix=".xlsx", only_copy_first_sheet=True,
        target_cur_sht=False, target_same_sheetname=False
    )


if __name__ == '__main__':
    from src.basic.sheet_func import sheet_seacrh

    if False:
        # test for sum_sheets
        filepath = cwdPath.joinpath('res\\dev\\sum_sheets')
        testfile_sum = filepath.joinpath('test_sum_sheets.xlsx')
        xw.Book(testfile_sum).set_mock_caller()
        try:
            wb_sum = xw.Book(testfile_sum)

            # delete the test data beforeclear()
            csv_sht_list = sheet_seacrh(
                wb_sum, ['test.csv', 'test', 'test_100.csv', 'test_100', 'test_1000.csv', 'test_1000'])
            for sht in csv_sht_list:
                if sht.name in ('test.csv', 'test', 'test_100.csv', 'test_100', 'test_1000.csv', 'test_1000'):
                    sht.delete()
            sheet_seacrh(wb_sum, ['Sheet1'])[0].clear()
            sheet_seacrh(wb_sum, ['Sheet2'])[0].clear()

            # test sum all .csv files to diff sheets with filename
            sum_sheets(
                ignore_row_num=0, sum_current=False, suffix=".csv",
                only_copy_first_sheet=True,
                target_cur_sht=False, target_same_sheetname=False
            )

            # test sum all .csv files to diff sheets with sheet name
            # sum_sheets(
            #     ignore_row_num=0, sum_current=False, suffix=".csv",
            #     only_copy_first_sheet=True,
            #     target_cur_sht=False, target_same_sheetname=True
            # )

            # test sum all .csv files to current sheet
            # sum_sheets(
            #     ignore_row_num=0, sum_current=False, suffix=".csv",
            #     only_copy_first_sheet=True,
            #     target_cur_sht=True, target_same_sheetname=False
            # )

            # test sum all sheets to "Sheet1" or sheet[0] in current workbook
            # sum_sheets(
            #     ignore_row_num=0, sum_current=True,
            #     target_cur_sht=False
            # )

            # test sum all sheets to current sheet in current workbook
            # sum_sheets(
            #     ignore_row_num=0, sum_current=True,
            #     target_cur_sht=True
            # )

        except Exception as e:
            log.error('An error occurred. {}'.format(e.args[-1]))
            raise
