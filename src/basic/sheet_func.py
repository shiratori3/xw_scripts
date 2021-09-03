#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   sheet_func.py
@Author  :   Billy Zhou
@Time    :   2020/08/26
@Desc    :   None
'''


import sys
from pathlib import Path
cwdPath = Path(__file__).parents[2]  # the num depend on your filepath
sys.path.append(str(cwdPath))

from src.manager.LogManager import logmgr
log = logmgr.get_logger(__name__)

import xlwings as xw
from xlwings import Book
from xlwings import Sheet
from typing import List
from src.basic.utils import search_file_by_type
from src.basic.app_func import try_expect_stop_refresh
from src.basic.cell_func import get_height


def sheet_seacrh(
        wb: Book, sht_name_list: List[str] = ['Sheet1'],
        create: bool = False, create_at_last: bool = False, return_first: bool = True) -> List[Sheet]:
    """find the sheet with sht_name in wb

    Args:
    ----
        wb: Book
            a Book instance of workbook
        sht_name_list: List[str], default ['Sheet1']
            a list of name of sheets to search
        create: bool, default False
            if unfound, create a worksheet named sht_name
        create_at_last: bool, default False
            create the worksheet at last, else behind the first sheet
        return_first: bool, default True
            if unfound, return the first sheet, else the last sheet

    Return:
    ------
        sht_target_list: List[Sheet]
            a list of Sheet instance
    """
    sht_target_list = []
    original_sht_name_list = [sht.name for sht in wb.sheets]
    log.debug("sht_name_list: {}".format(sht_name_list))
    for sht_name in sht_name_list:
        sht_target = None
        if sht_name not in original_sht_name_list:
            log.warning('sht_name[{}] unfound in wb.sheets{}'.format(sht_name, original_sht_name_list))
            if create:
                sht_target = wb.sheets.add(
                    sht_name, before=wb.sheets[0]) if not create_at_last else wb.sheets.add(
                        sht_name, after=wb.sheets[-1])
        else:
            sht_target = wb.sheets[sht_name]
        if not sht_target:
            sht_target = wb.sheets[0] if return_first else wb.sheets[-1]
        sht_target_list.append(sht_target)
    log.debug("sht_target_list: {}".format(sht_target_list))
    return sht_target_list


def copy_and_paste(
        sheet_copy: Sheet, sheet_paste: Sheet, ignore_row_num: int = 0,
        copy_region_used: bool = False, paste_region_used: bool = False,
        close_after_copy: bool = False) -> None:
    """copy a range from sheet_copy and paste it to sheet_paste

    Args:
    ----
        sheet_copy: Sheet
            the sheet instance for copying
        sheet_paste: Sheet
            the sheet instance for pasting
        ignore_row_num: int, default 0
            the row num to ignore for copying
        copy_region_used: bool, default False
        paste_region_used: bool, default False
            whether to use the used region, if false, use range("A1").current_region instead
        close_after_copy: bool, default False
            whether close workbook after pasted
    """
    # define the copy range of copy region
    copy_region = sheet_copy.used_range if copy_region_used else sheet_copy.range(
        "A1").current_region
    log.debug("copy_region: {}".format(copy_region))

    range_copy = copy_region.get_address(False, False)
    # change the start row if ignore_row_num
    if ignore_row_num:
        h_tuple = get_height(range_copy, return_details=True)
        copy_region_height, s_row = h_tuple[0], h_tuple[1]
        if copy_region_height > ignore_row_num:
            cells_list = range_copy.split(":")
            cells_list[0] = cells_list[0].replace(
                str(s_row), str(s_row + ignore_row_num))
            range_copy = ":".join(cells_list)
        else:
            log.error('ignore_row_num[{}] is equal or bigger than copy_region_height[{}]'.format(
                ignore_row_num, copy_region_height))
        log.debug("height_copy: {}".format(copy_region_height - ignore_row_num))
    log.debug("range_copy: {}".format(range_copy))

    # define the paste range of target sheet
    paste_region = sheet_paste.used_range if paste_region_used else sheet_paste.range("A1").current_region
    log.debug("paste_region.address: {}".format(paste_region.address))

    if paste_region.address != '$A$1':
        cell_s_paste = paste_region.address.split(":")[0]
        height_paste = get_height(paste_region.address)
    else:
        cell_s_paste = 'A1'
        height_paste = 0
    # paste after the last row in sheet_paste
    range_paste = sheet_paste.range(cell_s_paste).offset(row_offset=height_paste)
    log.debug("range_paste: %s", range_paste.address)

    # copy value from range_copy in sheet_copy to range_paste in sheet_paste
    sheet_copy.range(range_copy).copy(destination=range_paste)

    if close_after_copy:
        sheet_copy.book.close()


@try_expect_stop_refresh
def sum_sheets(
        ignore_row_num: int, sum_current: bool = False, suffix: str = ".csv",
        only_copy_first_sheet: bool = True,
        target_cur_sht: bool = True, target_same_sheetname: bool = False) -> None:
    """sum sheets from current workbook or other workbooks

    Args:
    ----
        ignore_row_num: int
            the row num to ignore for copying.
        sum_current: bool, default False
            sum from current workbook.
        suffix: str, default '.csv'
            sum from other files with suffix.
        only_copy_first_sheet: bool, default True
            while copy from other files, only copy the first sheet.
        target_cur_sht: bool, default True
            Paste into current active sheet in current workbook.
            Else, paste into "Sheet1" or sheet[0] if "Sheet1" not found.
        target_same_sheetname:  bool, default False
            If not target_cur_sht, paste into diff sheets with same sheetname in current workbook.
            Else, paste into diff sheets with their workbook.name

    Example:
    -------
        # test sum all .csv files to diff sheets with filename

        sum_sheets(
            ignore_row_num=0, sum_current=False, suffix=".csv",
            only_copy_first_sheet=True,
            target_cur_sht=False, target_same_sheetname=False
        )

        # test sum all .csv files to diff sheets with sheet name

        sum_sheets(
            ignore_row_num=0, sum_current=False, suffix=".csv",
            only_copy_first_sheet=True,
            target_cur_sht=False, target_same_sheetname=True
        )

        # test sum all .csv files to current sheet

        sum_sheets(
            ignore_row_num=0, sum_current=False, suffix=".csv",
            only_copy_first_sheet=True,
            target_cur_sht=True, target_same_sheetname=False
        )

        # test sum all sheets to "Sheet1" or sheet[0] in current workbook

        sum_sheets(ignore_row_num=0, sum_current=True, target_cur_sht=False)

        # test sum all sheets to current sheet in current workbook

        sum_sheets(ignore_row_num=0, sum_current=True, target_cur_sht=True)
    """
    wb_active = xw.books.active
    wb_active_dir, wb_active_name = Path(wb_active.fullname).parent, Path(wb_active.fullname).name

    # select current active sheet or "Sheet1" or sheet[0] as default target sheet
    sht_target = sheet_seacrh(
        wb_active, ['Sheet1'], return_first=True) if not target_cur_sht else wb_active.sheets.active
    target_shts_dict = {}
    target_shts_dict[0] = sht_target

    if sum_current:
        # sum all sheets in current workbook to target sheet
        log.debug("wb_active.sheets: {}".format(wb_active.sheets))

        for sht in wb_active.sheets:
            if sht != sht_target:
                print("Copying from {} to {} in {}".format(
                    sht.name, sht_target.name, wb_active.name))
                copy_and_paste(
                    wb_active.sheets[sht.name], sht_target,
                    ignore_row_num, close_after_copy=False)
    else:
        # sum sheets in other workbooks to target sheet
        # add workbook list for copying
        wb_copying_list = search_file_by_type(wb_active_dir, suffix)
        log.debug("wb_copying_list: {}".format(wb_copying_list))

        for wb_copying in wb_copying_list:
            log.debug("wb_copying: {}".format(wb_copying))
            if not (Path(wb_copying).name == wb_active_name or "~$" in wb_copying):
                print("Handling file: {}".format(wb_copying))
                # open dst file and define the sheet for copying
                wb_copying = xw.Book(wb_copying)
                # select which sheet to copy from wb_copy
                copy_shts_list = []
                for sht_copying in wb_copying.sheets:
                    if not target_cur_sht:
                        # if target_same_sheetname,
                        #     copy into sheet named same with sht_copying.name
                        # if not target_same_sheetname,
                        #     copy into sheet named wb_copying.name
                        target_name = sht_copying.name if target_same_sheetname else wb_copying.name
                        log.debug('target_name: %s' % target_name)
                        sht_target = sheet_seacrh(
                            wb_active, [target_name],
                            create=True, create_at_last=True)[0]
                        log.debug('sht_target: %s' % sht_target)
                        target_shts_dict[target_name] = sht_target

                    # copy first sheet or each sheet in wb_copying
                    copy_shts_list.append(
                        wb_copying.sheets[0]) if only_copy_first_sheet else copy_shts_list.append(
                            sht_copying)

                    if only_copy_first_sheet:
                        break

                log.debug('copy_shts_list: {}'.format(copy_shts_list))
                log.debug('target_shts_dict: {}'.format(target_shts_dict))
                for sht_copy in copy_shts_list:
                    # if target_cur_sht,
                    #     copy into wb_active.sheets.active
                    # else "Sheet1" or wb_active.sheet[0]
                    if target_cur_sht:
                        sht_paste = target_shts_dict[0]
                    else:
                        # if not target_cur_sht and target_same_sheetname,
                        #     copy into sheet named same with sht_copying.name
                        # if not target_cur_sht and not target_same_sheetname,
                        #     copy into sheet named wb_copying.name
                        if target_same_sheetname:
                            sht_paste = target_shts_dict[sht_copy.name]
                        else:
                            sht_paste = target_shts_dict[wb_copying.name]
                    # copy value from sht_copying and paste in sht_target
                    print("Copying {} from {} to {} in {}".format(
                        sht_copy.name, wb_copying.name,
                        sht_paste.name, wb_active.name))
                    copy_and_paste(
                        sht_copy, sht_paste,
                        ignore_row_num, close_after_copy=True)


if __name__ == '__main__':
    from pathlib import Path
    filepath = cwdPath.joinpath('res\\dev\\copy_and_paste')
    testfile_paste = filepath.joinpath('test_paste.xlsx')
    testfile_copy = filepath.joinpath('test_copy.xlsx')

    if False:
        # test for sheet_seacrh
        wb_paste = xw.Book(testfile_paste)
        wb_copy = xw.Book(testfile_copy)

        sht_copy1 = sheet_seacrh(wb_copy, ['Sheet0'], return_first=True)
        log.info(sht_copy1)
        sht_copy2 = sheet_seacrh(wb_copy, ['Sheet4'], return_first=False)
        log.info(sht_copy2)
        sht_copy2 = sheet_seacrh(wb_copy, ['Sheet4', ], return_first=False)
        log.info(sht_copy2)
        sht_paste1 = sheet_seacrh(wb_paste, ['Sheet000'], create=True)
        log.info(sht_paste1)
        sht_paste1.delete()
        sht_paste2 = sheet_seacrh(wb_paste, ['Sheet100'], create=True, create_at_last=True)
        log.info(sht_paste2)
        sht_paste2.delete()

    if False:
        # test for copy_and_paste
        copy_and_paste(
            sht_copy1, sht_paste1,
            ignore_row_num=0, close_after_copy=False,
            copy_region_used=False, paste_region_used=False)
        copy_and_paste(
            sht_copy2, sht_paste2,
            ignore_row_num=5, close_after_copy=False,
            copy_region_used=True, paste_region_used=True)
