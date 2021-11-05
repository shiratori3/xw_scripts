#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   cell_func.py
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

import re
from decimal import Decimal
from typing import Tuple


def str2num(input_value: str) -> Decimal or str:
    """try to turn input_value to Decimal, if failed, return input_value"""
    try:
        num_out = Decimal(input_value)
        log.debug("num_out: {}".format(num_out))
        log.debug("type(num_out): {}".format(type(num_out)))
        return num_out
    except Exception as e:
        log.warning('An error occurred. {}'.format(e.args))
        log.debug("Input for str2num is not numberic.")
        log.debug("input_value: {}".format(input_value))
        log.debug("type(input_value): {}".format(type(input_value)))
        return input_value


def get_height(address: str, return_details: bool = False) -> int or Tuple[int, int, int]:
    """get the height of address

    Example:
    --------
    >>> get_height('A5')
    5
    >>> get_height('A1:B10')
    10
    >>> get_height('A1:B10', return_details=True)
    (10, (1, 10))
    """
    cells_list = address.split(":")
    if len(cells_list) > 1:
        s_num, e_num = get_num(cells_list[0]), get_num(cells_list[1])
    else:
        s_num, e_num = 1, get_num(cells_list[0])
    return e_num - s_num + 1 if not return_details else (e_num - s_num + 1, s_num, e_num)


def get_num(inputed: str):
    """get num part"""
    return int(re.split("\\D+", inputed)[-1])


def num2alpha(num: int) -> str:
    """turn num to alpha"""
    return "" if num == 0 else num2alpha((num - 1) // 26) + chr((num - 1) % 26 + ord('A'))


if __name__ == '__main__':
    if True:
        # test for str2num
        log.info(str2num('2222') * -1)
        log.info(str2num('2222.2222') * -1)
        log.info(str2num('2222.') * -1)
        log.info(str2num('.2222') * -1)
        log.info(str2num('-.2222') * -1)
        log.info(str2num('nan') * -1)
        log.info(type(str2num('nan')))
        log.info(str2num('inf') * -1)
        log.info(str2num('-inf') * -1)
        log.info(type(str2num('inf')))
        log.info(str2num('222222222222222222222222222222222222222.222222222222222222222').to_eng_string())
        log.info(type(str2num('222222222222222222222222222222222222222.222222222222222222222')))
        log.info(str2num(' （7 ,495,420,538) '))
        log.info(str2num('(933755422)'))
        log.info(str2num('aaa'))

    if False:
        # test for address_getnum
        log.info(get_num('A10000'))

    if False:
        # test for get_height
        log.info(get_height('D100:F293'))
        log.info(get_height('D100'))
        h_tuple = get_height('D100:F293', return_details=True)
        log.info("h_tuple: {}".format(h_tuple))
        log.info("height: {}, start: {}, end: {}".format(*get_height('D100:F293', return_details=True)))

    if False:
        # test for num2alpha
        log.info("output: " + num2alpha(1))
        log.info("output: " + num2alpha(26))
        log.info("output: " + num2alpha(27))
        log.info("output: " + num2alpha(52))
        log.info("output: " + num2alpha(53))
        log.info("output: " + num2alpha(78))
        log.info("output: " + num2alpha(676))
        log.info("output: " + num2alpha(677))
