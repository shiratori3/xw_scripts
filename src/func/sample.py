#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   sample.py
@Author  :   Billy Zhou
@Time    :   2021/10/21
@Desc    :   None
'''


import sys
from pathlib import Path

cwdPath = Path(__file__).parents[2]  # the num depend on your filepath
sys.path.append(str(cwdPath))

from src.manager.LogManager import logmgr
log = logmgr.get_logger(__name__)

import xlwings as xw
import numpy as np
import random


@xw.func
@xw.arg('data', np.array, ndim=2, numbers=str, empty='', doc='Select the range you want to sample')
@xw.arg('sample_num', doc='The num or percent you want to catch from the data. Default is 10')
@xw.ret(expand='table', transpose=True)
def udf_sample(data, sample_num):
    """Return a boolean array according to the sample data and sample num."""
    row_num = data.shape[0]
    if isinstance(sample_num, float):
        if int(sample_num) == sample_num:
            if sample_num >= row_num:
                sample_row_num = row_num
            else:
                sample_row_num = int(sample_num)
        else:
            # float input
            if sample_num > 0 and sample_num < 1:
                sample_row_num = int(row_num * sample_num)
            else:
                return 'Error input value of sample_num, must be a integer or a float between 0 and 1'
    else:
        return 'Error input type of sample_num'

    rand_array = np.array([random.random() for _ in range(row_num)])
    return np.where(rand_array <= np.sort(rand_array)[sample_row_num - 1], True, False)


if __name__ == '__main__':
    xw.serve()
