#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   func_basic.py
@Author  :   Billy Zhou
@Time    :   2021/03/03
@Version :   1.1.0
@Desc    :   None
'''


import os
import logging
import re
import numpy as np
from decimal import Decimal
from string import ascii_uppercase
from collections import Counter


def count_sum(list_count, *args):
    result_sum = 0
    dict_count = Counter(list_count)
    for i in args:
        result_sum += dict_count[i]
    logging.debug('result_sum: %s', result_sum)
    return result_sum


def num_to_ascii(num, digits=0, b=26, string_ABC=""):
    logging.debug("num: %s" % str(num))
    logging.debug("bit of num: %s" % str(num // b))
    logging.debug("rmd of num: %s" % str(num % b))
    logging.debug("string: %s" % string_ABC)

    if num // b != 0 or num % b != 0:
        logging.debug("chr of num: %s" % ascii_uppercase[num % b - 1])
    if num // b == 0 and (
            (num % b == 0) or (
            digits > 0 and num % b == 1 and string_ABC[-1] == "Z")):
        return string_ABC
    else:
        digits = + 1
        return num_to_ascii(
            num//b, digits, b, ascii_uppercase[num % b - 1] + string_ABC)


def selected_range(cell_selected):
    # change col num to col alpha
    col_alpha = col_num2alpha(
        cell_selected.column, cell_selected.last_cell.column)

    # debug range of selection
    cell_start = col_alpha[
        cell_selected.column] + str(cell_selected.row)
    cell_end = col_alpha[
        cell_selected.last_cell.column] + str(cell_selected.last_cell.row)
    logging.info("Start: %s", cell_start)
    logging.info("End: %s", cell_end)
    return cell_start, cell_end


def col_num2alpha(col_min, col_max):
    col_alpha = {}
    for i in range(col_min, col_max+1):
        col_alpha[i] = num_to_ascii(i)
    logging.debug("col_alpha: %s" % col_alpha)
    return col_alpha


def address_getnum(cell):
    logging.debug("NUM: %s" % int(re.split("\\D+", cell)[-1]))
    return int(re.split("\\D+", cell)[-1])


def address_getheight(address):
    # get the height of address
    cells_list = address.split(":")
    logging.debug('cells_list: %s' % cells_list)
    if len(cells_list) > 1:
        cell_s_rownum = address_getnum(cells_list[0])
        cell_e_rownum = address_getnum(cells_list[1])
    else:
        cell_s_rownum = 1
        cell_e_rownum = address_getnum(cells_list[0])
    height = cell_e_rownum - cell_s_rownum + 1
    return height


def filename_split(name):
    return name.split(".")[0], name.split(".")[-1]


def search_file_by_type(f_path, f_type):
    result_list = []
    for i in os.listdir(f_path):
        if i[-len(f_type):] == f_type:
            result_list.append(i)
    return result_list


def sheet_seacrh(
        wb_for_search, sht_name='Sheet1', default_first=True,
        create=False, create_last=False):
    sht_target = ''
    sht_name_list = []
    for sht in wb_for_search.sheets:
        sht_name_list.append(sht.name)
    if sht_name in sht_name_list:
        sht_target = wb_for_search.sheets[sht_name]
    if create:
        if not create_last:
            sht_target = wb_for_search.sheets.add(
                sht_name, before=wb_for_search.sheets[0])
        else:
            sht_target = wb_for_search.sheets.add(
                sht_name, after=wb_for_search.sheets[-1])
    if not sht_target and default_first:
        sht_target = wb_for_search.sheets[0]
    return sht_target


def num_check(str_input):
    logging.debug("str_input: %s", str_input)
    logging.debug("type(str_input): %s", type(str_input))
    if isinstance(str_input, (np.str_, str)):
        try:
            float(str_input)
            return True
        except ValueError:
            logging.debug("Input for num_check is not numberic.")
            return False
    else:
        logging.debug("Input for num_check is not str.")
        return False


def str2num(str_input, nan_vaild=True, inf_vaild=True):
    logging.debug("str_input: %s", str_input)
    logging.debug("type(str_input): %s", type(str_input))
    if num_check(str_input):
        input_lower = str_input.lower()
        if input_lower == 'nan':
            logging.debug("float value: NaN")
            num_out = 'NaN' if nan_vaild else ''
        elif 'inf' in input_lower:
            logging.debug("float value: %s", input_lower)
            num_out = input_lower if inf_vaild else ''
        else:
            num_out = Decimal(str_input)
        logging.debug("num_out: %s", num_out)
        logging.debug("type(num_out): %s", type(num_out))
        return num_out
    else:
        return ''


def dict_append(
        dict_data, info_message, success_append, failure_append,
        info_lvl='INFO'):
    if info_lvl == 'CRITICAL':
        logging.critical(info_message)
    elif info_lvl == 'ERROR':
        logging.warning(info_message)
    elif info_lvl == 'WARNING':
        logging.error(info_message)
    elif info_lvl == 'DEBUG':
        logging.debug(info_message)
    else:
        logging.info(info_message)

    dict_data['success'].append(success_append)
    dict_data['failure'].append(failure_append)


if __name__ == '__main__':
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(name)s %(levelname)s %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S')
    logging.debug('start DEBUG')
    logging.debug('==========================================================')

    # test for count_sum
    # list_test = [1, 2, 2, 2, 3, 5]
    # logging.info("count: %s" % count_sum(list_test, 2, 5))

    # test for num_to_ascii
    # logging.info("output: " + num_to_ascii(1))
    # logging.info("output: " + num_to_ascii(26))
    # logging.info("output: " + num_to_ascii(27))
    # logging.info("output: " + num_to_ascii(52))
    # logging.info("output: " + num_to_ascii(53))
    # logging.info("output: " + num_to_ascii(78))
    # logging.info("output: " + num_to_ascii(676))
    # logging.info("output: " + num_to_ascii(677))

    # test for num_to_ascii
    # logging.info("col_num2alpha: %s" % col_num2alpha(2, 5))

    # test for address_getnum
    # logging.info(address_getnum('A10000'))
    # logging.info(type(address_getnum('A10000')))

    # test for address_getheight
    # logging.info(address_getheight('D100:F293'))
    # logging.info(type(address_getheight('D100:F293')))
    # logging.info(address_getheight('D100'))
    # logging.info(type(address_getheight('D100')))

    # test for str2num
    print(str2num('2222') * -1)
    print(str2num('2222.2222') * -1)
    print(str2num('2222.2222') * -1)
    print(str2num('2222.') * -1)
    print(str2num('.2222') * -1)
    print(str2num('-.2222') * -1)
    print(str2num('nan'))
    print(str2num('aaa'))
    print(str2num('inf'))
    print(str2num('-inf'))

    logging.debug('==========================================================')
    logging.debug('end DEBUG')
