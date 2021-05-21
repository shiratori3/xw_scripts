#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   func_basic_excel_cell.py
@Author  :   Billy Zhou
@Time    :   2021/03/23
@Version :   1.0.0
@Desc    :   None
'''


import logging
import re
from string import ascii_uppercase


def range_get_cell_start_to_end(selected_range):
    # change col num to col alpha
    col_alpha = col_num2alpha(
        selected_range.column, selected_range.last_cell.column)

    # debug range of selection
    cell_start = col_alpha[
        selected_range.column] + str(selected_range.row)
    cell_end = col_alpha[
        selected_range.last_cell.column] + str(selected_range.last_cell.row)
    logging.info("Start: %s", cell_start)
    logging.info("End: %s", cell_end)
    return cell_start, cell_end


def col_num2alpha(col_min, col_max):
    col_alpha = {}
    for i in range(col_min, col_max + 1):
        col_alpha[i] = num_to_ascii(i)
    logging.debug("col_alpha: %s" % col_alpha)
    return col_alpha


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
            num // b, digits, b, ascii_uppercase[num % b - 1] + string_ABC)


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


def address_getnum(cell):
    logging.debug("NUM: %s" % int(re.split("\\D+", cell)[-1]))
    return int(re.split("\\D+", cell)[-1])


if __name__ == '__main__':
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(name)s %(levelname)s %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S')
    logging.debug('start DEBUG')
    logging.debug('==========================================================')

    # test for num_to_ascii
    # logging.info("col_num2alpha: %s" % col_num2alpha(2, 5))

    # test for num_to_ascii
    # logging.info("output: " + num_to_ascii(1))
    # logging.info("output: " + num_to_ascii(26))
    # logging.info("output: " + num_to_ascii(27))
    # logging.info("output: " + num_to_ascii(52))
    # logging.info("output: " + num_to_ascii(53))
    # logging.info("output: " + num_to_ascii(78))
    # logging.info("output: " + num_to_ascii(676))
    # logging.info("output: " + num_to_ascii(677))

    # test for address_getheight
    # logging.info(address_getheight('D100:F293'))
    # logging.info(type(address_getheight('D100:F293')))
    # logging.info(address_getheight('D100'))
    # logging.info(type(address_getheight('D100')))

    # test for address_getnum
    # logging.info(address_getnum('A10000'))
    # logging.info(type(address_getnum('A10000')))

    logging.debug('==========================================================')
    logging.debug('end DEBUG')
