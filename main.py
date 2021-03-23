#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   main.py
@Author  :   Billy Zhou
@Time    :   2020/03/23
@Version :   1.2.0
@Desc    :   None
'''

import os
import logging
import xlwings as xw
import pandas as pd
from func_basic import search_file_by_type
from func_basic_excel_sht import sheet_seacrh
from func_basic_excel_sht import excel_range_func
from func_change_minus import array_change_minus
from func_cleansing import array_regex_cleansing
from func_copy import copy_and_paste


def stop_refresh():
    xw.apps.active.screen_updating = False


def start_refresh():
    if not xw.apps.active.screen_updating:
        xw.apps.active.screen_updating = True


def to_csv(filetype=".xlsx"):
    wb = xw.books.active
    wb_path, wb_name = os.path.split(wb.fullname)
    logging.debug(os.listdir(wb_path))
    file_list = search_file_by_type(wb_path, filetype)
    if file_list:
        for i in file_list:
            if '~$' not in i:
                file_path = wb_path + '\\' + i
                csv_file_path = wb_path + '\\' + i[:-len(filetype)] + ".csv"
                logging.debug("file name: " + file_path)
                logging.debug("csv name: " + csv_file_path)
                pd.read_excel(
                    file_path, index_col=0).to_csv(
                    csv_file_path, encoding='utf-8-sig')


def num_clean(
        point_exist=True, point_digit=2, point_overflow_accept=False):
    # change input value to text by format "@"
    # take care of the max value of the float64 type
    if point_exist and point_digit > 0:
        num_format = "#,##0." + "0" * point_digit
    else:
        num_format = "#,##0"
    excel_range_func(
        excel_func=array_regex_cleansing, pre_format=num_format,
        point_exist=point_exist, point_digit=point_digit,
        point_overflow_accept=point_overflow_accept
        )


def change_minus():
    excel_range_func(excel_func=array_change_minus)


def sum_sheets(
        ignore_row_num, current_book=False, f_type=".csv",
        copy_first_sheet=True,
        target_sheet1=True, target_same_sheetname=False):
    wb_active = xw.books.active
    wb_active_path, wb_active_name = os.path.split(wb_active.fullname)

    # select "Sheet1" or the first sheet as default target sheet
    sht_target = sheet_seacrh(wb_active, 'Sheet1')
    shts_target = {}
    shts_target[0] = sht_target

    if current_book:
        # sum all sheets in current workbook to target sheet
        logging.debug(wb_active.sheets)

        for sht in wb_active.sheets:
            if sht != sht_target:
                print("Copying from: {0} in {1}".format(
                    sht.name, wb_active.name))
                copy_and_paste(
                    wb_active.sheets[sht.name], sht_target,
                    ignore_row_num, close_after_copy=False)
    else:
        # sum sheets in other workbooks to target sheet
        # add workbook list for copying
        f_type = "." + f_type if f_type[0] != "." else f_type
        file_copy_list = search_file_by_type(wb_active_path, f_type)
        logging.debug(file_copy_list)

        for wb_copy_name in file_copy_list:
            if not (wb_copy_name == wb_active_name or "~$" in wb_copy_name):
                print("Handling workbook: {0} under path {1}".format(
                    wb_copy_name, wb_active_path))
                # open dst file and define the sheet for copying
                logging.info("\\".join([wb_active_path, wb_copy_name]))
                wb_copying = xw.Book("\\".join([wb_active_path, wb_copy_name]))
                if wb_copying:
                    # select which sheet to copy from wb_copy
                    shts_copying = []
                    for sht_copying in wb_copying.sheets:
                        if copy_first_sheet:
                            # the first sheet: sheets[0]
                            shts_copying.append(wb_copying.sheets[0])
                        else:
                            # All sheets
                            shts_copying.append(sht_copying)

                        # copy all data into same target sheet or not
                        # default target sheet is
                        #    wb_active.sheets["Sheet1"] or sheets[0]
                        # if not:
                        #    change target to sheet named copy_sheet_name
                        if not target_sheet1:
                            target_name = sht_copying.name if target_same_sheetname else wb_copy_name  # noqa: E501
                            logging.info('target_name: %s' % target_name)
                            sht_target = sheet_seacrh(
                                wb_active, target_name,
                                create=True, create_last=True)
                            logging.info('sht_target: %s' % sht_target)
                            shts_target[target_name] = sht_target

                        if copy_first_sheet:
                            break

                    logging.info('shts_copying: %s' % shts_copying)
                    logging.info('shts_target: %s' % shts_target)
                    for sht_copy in shts_copying:
                        print("Copying {0} from {1} to {2} in {3}".format(
                            sht_copy.name, wb_copying.name,
                            sht_target.name, wb_active.name))
                        logging.info('target_sheet1: %s' % target_sheet1)
                        if copy_first_sheet:
                            sht_paste = shts_target[0]
                        if target_sheet1:
                            sht_paste = shts_target[0]
                        else:
                            if target_same_sheetname:
                                sht_paste = shts_target[sht_copy.name]
                            else:
                                sht_paste = shts_target[wb_copy_name]
                        # copy value from sht_copying and paste in sht_target
                        copy_and_paste(
                            sht_copy, sht_paste,
                            ignore_row_num, close_after_copy=True)


if __name__ == '__main__':
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(name)s %(levelname)s %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S')
    logging.debug('start DEBUG')
    logging.debug('==========================================================')

    from pathlib import Path

    # test for to_csv

    # filepath = Path.cwd().joinpath('test_data\\to_csv')
    # testfile = str(filepath.joinpath('test.xlsx'))
    # xw.Book(testfile).set_mock_caller()
    # to_csv(filetype=".xlsx")
    # to_csv(filetype=".xls")

    # test for excel_cleaning

    # filepath = Path.cwd().joinpath('test_data\\cleansing')
    # testfile = str(filepath.joinpath('regex_test.xlsx'))
    # xw.Book(testfile).set_mock_caller()
    # stop_refresh()
    # num_clean(
    #     point_exist=True, point_digit=2, point_overflow_accept=False)
    # start_refresh()

    # test for change_minus

    # filepath = Path.cwd().joinpath('test_data\\change_minus')
    # testfile = str(filepath.joinpath('change_test.xlsx'))
    # xw.Book(testfile).set_mock_caller()
    # stop_refresh()
    # change_minus()
    # start_refresh()

    # test for sum_sheets

    filepath = Path.cwd().joinpath('test_data\\sum_sheets')
    testfile = str(filepath.joinpath('sum_test.xlsx'))
    xw.Book(testfile).set_mock_caller()
    stop_refresh()
    sum_sheets(
        ignore_row_num=0, current_book=False, f_type=".csv",
        copy_first_sheet=True,
        target_sheet1=False, target_same_sheetname=False
    )
    start_refresh()

    logging.debug('==========================================================')
    logging.debug('end DEBUG')
