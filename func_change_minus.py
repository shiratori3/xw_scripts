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
from func_basic import dict_append
from func_basic import str2num


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
                            dict_out, "Num", str(num_line * -1), "", 'DEBUG')
                    else:
                        dict_append(
                            dict_out, "Str", line, line, 'INFO')
                else:
                    dict_append(
                        dict_out, "Str", line, line, 'INFO')

    return dict_out['success'], dict_out['failure']


if __name__ == '__main__':
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(name)s %(levelname)s %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S')
    logging.debug('start DEBUG')
    logging.debug('==========================================================')

    logging.debug('==========================================================')
    logging.debug('end DEBUG')
