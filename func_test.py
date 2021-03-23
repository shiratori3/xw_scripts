#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   func_test.py
@Author  :   Billy Zhou
@Time    :   2020/02/25
@Version :   0.1.0
@Desc    :   None
'''


import logging


if __name__ == '__main__':
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(name)s %(levelname)s %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S')
    logging.debug('start DEBUG')
    logging.debug('==========================================================')

    logging.debug('==========================================================')
    logging.debug('end DEBUG')
