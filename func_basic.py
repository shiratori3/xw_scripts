#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   func_basic.py
@Author  :   Billy Zhou
@Time    :   2021/03/03
@Version :   1.2.0
@Desc    :   None
'''


import os
import logging
import numpy as np
from decimal import Decimal

from collections import Counter


def count_sum(list_count, *args):
    result_sum = 0
    dict_count = Counter(list_count)
    for i in args:
        result_sum += dict_count[i]
    logging.debug('result_sum: %s', result_sum)
    return result_sum


def search_file_by_type(filepath, filetype='json'):
    file_list = [os.path.join(root, f) for root, dirs, files in os.walk(filepath) for f in files if f[-len(filetype):] == filetype]  # noqa: E501
    logging.debug(file_list[0])
    logging.debug(file_list[-1])
    logging.debug(file_list)
    return file_list


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
