#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   main.py
@Author  :   Billy Zhou
@Time    :   2021/09/03
@Desc    :   None
'''

import sys
from pathlib import Path
cwdPath = Path(__file__).parents[1]  # the num depend on your filepath
sys.path.append(str(cwdPath))


import xlwings as xw
from src.basic.utils import to_csv  # noqa: F401
from src.func.md5 import udf_md5  # noqa: F401
from src.func.md5 import udf_md5_single  # noqa: F401
from src.func.sample import udf_sample  # noqa: F401
from src.func.arrange import udf_arrange  # noqa: F401
from src.func.arrange import udf_arrange_mutli  # noqa: F401
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


@xw.sub
def from_xlsx_to_csv():
    to_csv(suffix=".xlsx")


@xw.sub
def from_xls_to_csv():
    to_csv(suffix=".xls")


if __name__ == '__main__':
    xw.serve()
