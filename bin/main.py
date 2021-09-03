#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   common_udfs.py
@Author  :   Billy Zhou
@Time    :   2021/05/21
@Version :   1.3.0
@Desc    :   None
'''

import sys
from pathlib import Path
cwdPath = Path(__file__).parents[1]  # the num depend on your filepath
sys.path.append(str(cwdPath))

from src.manager.LogManager import logmgr
log = logmgr.get_logger(__name__)

import hashlib
import numpy as np
import xlwings as xw
from src.basic.utils import to_csv
from src.marco.range_marco import change_selected_sign  # noqa: F401
from src.marco.range_marco import change_selected_sign_preformat_text  # noqa: F401
from src.marco.range_marco import clean_num_with2dec_round  # noqa: F401
from src.marco.range_marco import clean_num_with2dec_truncate  # noqa: F401
from src.marco.range_marco import clean_num_with4dec_round  # noqa: F401
from src.marco.range_marco import clean_num_with0dec_truncate  # noqa: F401
from src.marco.sheet_marco import merge_cur_wb_to_cur_sheet_NoRow1  # noqa: F401
from src.marco.sheet_marco import merge_cur_wb_to_first_sheet_NoRow1  # noqa: F401
from src.marco.sheet_marco import merge_csv_to_cur_sheet  # noqa: F401
from src.marco.sheet_marco import merge_csv_to_diff_sheets  # noqa: F401
from src.marco.sheet_marco import merge_xls_first_sheet_to_cur_sheet_NoRow1  # noqa: F401
from src.marco.sheet_marco import merge_xlsx_first_sheet_to_cur_sheet_NoRow1  # noqa: F401
from src.marco.sheet_marco import merge_xlsx_first_sheet_to_diff_sheets_with_bookname  # noqa: F401
from src.marco.sheet_marco import merge_xlsx_first_sheet_to_diff_sheets_with_sheetname  # noqa: F401


# @xw.arg('data', 'range')
@xw.func
@xw.arg('data', ndim=2, numbers=str, empty='')
def MDF_SINGLE(data):
    data_joined = ''.join(list(np.array(data).flatten()))
    md5_str = hashlib.md5(data_joined.encode(encoding='UTF-8')).hexdigest()

    return md5_str


@xw.func
@xw.arg('data', ndim=2, numbers=str, empty='')
@xw.ret(expand='table')
def MDF(data, debug_flag=False):
    log.info(data)
    data_array = np.array(data)
    data_list = list(data_array.flatten())
    data_list_md5 = []
    for i, datum in enumerate(data_list):
        data_list_md5.append(hashlib.md5(datum.encode(encoding='UTF-8')).hexdigest())

    data_array_md5 = np.array(data_list_md5).reshape(data_array.shape)
    log.info(data_array_md5)

    if debug_flag:
        return data_array
    else:
        return data_array_md5


@xw.sub
def from_xlsx_to_csv():
    to_csv(suffix=".xlsx")


@xw.sub
def from_xls_to_csv():
    to_csv(suffix=".xls")


if __name__ == '__main__':
    xw.serve()
