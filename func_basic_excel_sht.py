#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   func_basic_excel.py
@Author  :   Billy Zhou
@Time    :   2021/03/23
@Version :   1.0.0
@Desc    :   None
'''


import logging
import xlwings as xw
import numpy as np
from func_basic_excel_cell import range_get_cell_start_to_end


def sheet_seacrh(
        wb_for_search, sht_name='Sheet1', default_first=True,
        create=False, create_last=False):
    sht_target = ''
    sht_name_list = []
    for sht in wb_for_search.sheets:
        sht_name_list.append(sht.name)
    if sht_name not in sht_name_list and create:
        if not create_last:
            sht_target = wb_for_search.sheets.add(
                sht_name, before=wb_for_search.sheets[0])
        else:
            sht_target = wb_for_search.sheets.add(
                sht_name, after=wb_for_search.sheets[-1])
    else:
        sht_target = wb_for_search.sheets[sht_name]
    if not sht_target and default_first:
        sht_target = wb_for_search.sheets[0]
    return sht_target


def excel_range_func(
        excel_func, pre_format='', after_format_num2text=False,
        *args, **kwargs):
    wb = xw.books.active
    sht = wb.sheets.active
    selected_range = wb.app.selection
    # set format selected range of before handling
    if pre_format:
        selected_range.number_format = pre_format
    # get value of selected range and turn to list
    selected_range_np = selected_range.options(
        np.array, ndim=2, empty='').value
    data_list = list(selected_range_np.flat)

    # apply the func to the list and get result
    success_list, failure_list = excel_func(data_list, *args, **kwargs)
    logging.warning("success: %s", success_list)
    logging.warning("failure: %s", failure_list)

    # turn the result to array and put it to selected range
    data_array = np.array(success_list).reshape(
        selected_range_np.shape[0], selected_range_np.shape[1])
    logging.info("data_array: %s" % data_array)

    # set cell format to text in selected range for handling
    if after_format_num2text:
        text_range = []
        x = 0
        for row in range(0, selected_range_np.shape[0]):
            for col in range(0, selected_range_np.shape[1]):
                x += 1
                cleaned_data = success_list[x-1]
                if cleaned_data:
                    logging.debug("value: %s" % (cleaned_data))
                    logging.debug("type: %s" % (type(cleaned_data)))
                    cell_row = selected_range.row + row
                    cell_col = selected_range.column + col
                    cell_range = (cell_row, cell_col)

                    if isinstance(cleaned_data, np.float64):
                        pass
                    elif isinstance(cleaned_data, np.str_):
                        text_range.append(cell_range)
                    elif len(cleaned_data) > 17:
                        text_range.append(cell_range)
        logging.debug(text_range)
        for text_cell in text_range:
            sht.range(text_cell).number_format = "@"

    cell_start, cell_end = range_get_cell_start_to_end(selected_range)
    cell_region = cell_start + ":" + cell_end
    logging.debug("cell_region: %s" % cell_region)

    sht.range(cell_region).value = data_array


if __name__ == '__main__':
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(name)s %(levelname)s %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S')
    logging.debug('start DEBUG')
    logging.debug('==========================================================')

    logging.debug('==========================================================')
    logging.debug('end DEBUG')
