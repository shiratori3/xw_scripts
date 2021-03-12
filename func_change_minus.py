#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   func_change_minus.py
@Author  :   Billy Zhou
@Time    :   2020/03/11
@Version :   1.0.0
@Desc    :   None
'''


import logging
import numpy as np
import xlwings as xw
from func_basic import dict_append
from func_basic import str2num
from func_basic import selected_range


def array_change_minus(data_list):
    dict_out = {
        'success': [],
        'failure': [],
    }

    for line in data_list:
        logging.debug("line: %s", line)
        logging.debug("type(line): %s", type(line))
        if not isinstance(line, (np.str_, str)):
            # cleaned num
            dict_append(dict_out, "Float", line * -1, "", 'DEBUG')
        else:
            if not line.strip():
                # skip for "\n" row
                dict_append(dict_out, "Blank cell", "", "", 'INFO')
            else:
                logging.debug("Orginal: %s", line.strip())
                num_line = str2num(line)
                logging.debug("num_line: %s", num_line)
                logging.debug("type(num_line): %s", type(num_line))
                if num_line:
                    if not isinstance(num_line, str):
                        dict_append(
                            dict_out, "Num", str(num_line*-1), "", 'DEBUG')
                    else:
                        dict_append(
                            dict_out, "Str", line, line, 'INFO')
                else:
                    dict_append(
                        dict_out, "Str", line, line, 'INFO')

    return dict_out['success'], dict_out['failure']


def excel_change_minus():
    wb = xw.books.active
    sht = wb.sheets.active
    cell_selected = wb.app.selection
    cell_selected_np = cell_selected.options(np.array, ndim=2, empty='').value
    data_list = list(cell_selected_np.flat)

    # change minus
    success_list, failure_list = array_change_minus(data_list)
    logging.warning("success: %s", success_list)
    logging.warning("failure: %s", failure_list)

    # put changed data to cell
    data_array = np.array(success_list).reshape(
        cell_selected_np.shape[0], cell_selected_np.shape[1])
    logging.info("data_array: %s" % data_array)
    cell_start, cell_end = selected_range(cell_selected)
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

    from pathlib import Path
    from main import stop_refresh
    from main import start_refresh
    filepath = Path.cwd().joinpath('test_data\\change_minus')
    testfile = str(filepath.joinpath('change_test.xlsx'))
    xw.Book(testfile).set_mock_caller()
    stop_refresh()
    excel_change_minus()
    start_refresh()

    logging.debug('==========================================================')
    logging.debug('end DEBUG')
