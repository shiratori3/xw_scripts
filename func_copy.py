#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   func_copy.py
@Author  :   Billy Zhou
@Time    :   2021/03/03
@Version :   1.1.0
@Desc    :   None
'''


import logging
import xlwings as xw
from func_basic import sheet_seacrh
from func_basic import address_getnum
from func_basic import address_getheight


def copy_and_paste(
        sheet_copy, sheet_paste, ignore_row_num=0,
        close_after_copy=False,
        copy_region_used=False, paste_region_used=False):
    # define the copy range of copy region
    if copy_region_used:
        copy_region = sheet_copy.used_range
    else:
        copy_region = sheet_copy.range("A1").current_region
    range_copy = copy_region.address

    # get the height of copy region
    cells_list = copy_region.address.split(":")
    cell_s_rownum = address_getnum(cells_list[0])
    cell_e_rownum = address_getnum(cells_list[1])
    height_copy = cell_e_rownum - cell_s_rownum + 1
    if ignore_row_num:
        # change the start row
        cells_list[0] = cells_list[0].replace(
            cell_s_rownum, str(cell_s_rownum + ignore_row_num))
        range_copy = ":".join(cells_list)
    height_copy = height_copy - ignore_row_num

    logging.debug("range_copy: %s" % range_copy)
    logging.debug("height_copy: %s" % height_copy)

    # define the paste range of target sheet
    if paste_region_used:
        paste_region = sheet_paste.used_range
    else:
        paste_region = sheet_paste.range("A1").current_region
    logging.debug("paste_region: %s", paste_region)
    logging.debug("     address: %s", paste_region.address)

    if paste_region.address != '$A$1':
        cell_s_paste = paste_region.address.split(":")[0]
        height_paste = address_getheight(paste_region.address)
    else:
        cell_s_paste = 'A1'
        height_paste = 0
    logging.debug("cell_s_paste: %s", cell_s_paste)
    logging.debug("height_paste: %s", height_paste)

    range_paste = sheet_paste.range(
        cell_s_paste).offset(row_offset=height_paste)
    logging.debug("range_paste: %s", range_paste.address)

    # copy value from range_copy in sheet_copy to range_paste in sheet_paste
    sheet_copy.range(range_copy).copy(destination=range_paste)

    if close_after_copy:
        sheet_copy.book.close()


if __name__ == '__main__':
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(name)s %(levelname)s %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S')
    logging.debug('start DEBUG')
    logging.debug('==========================================================')

    from pathlib import Path
    filepath = Path.cwd().joinpath('test_data\\copy_and_paste')
    testfile_target = str(filepath.joinpath('target.xlsx'))
    testfile_copy = str(filepath.joinpath('copy.xlsx'))

    wb_target = xw.Book(testfile_target)
    wb_copy = xw.Book(testfile_copy)

    sht_target = sheet_seacrh(wb_target)
    sht_copy = sheet_seacrh(wb_copy)

    logging.info(sht_target)
    logging.info(sht_target)
    copy_and_paste(
        sht_copy, sht_target,
        ignore_row_num=0, close_after_copy=False,
        copy_region_used=False, paste_region_used=False)

    logging.debug('==========================================================')
    logging.debug('end DEBUG')
