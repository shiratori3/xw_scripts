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
import pandas as pd
from numpy.random import default_rng
from typing import Any


@xw.func
@xw.arg('data', np.array, ndim=2, numbers=str, empty='', doc='The range of values you want to sample')
@xw.arg('sample_num', doc='The num or percent you want to catch from the sample data. Default is 10')
@xw.ret(expand='table', transpose=True)
def udf_sample(data, sample_num: Any = 10):
    """Return a list of boolean values which the num of True is base on data and sample_num"""
    row_num = data.shape[0]
    if isinstance(sample_num, (int, float)):
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

    rng = default_rng()
    rand_array = rng.random(row_num)
    return np.where(rand_array <= np.sort(rand_array)[sample_row_num - 1], True, False)


@xw.func
@xw.arg('data', pd.DataFrame, index=False, header=False, numbers=str, empty='', doc='The range of values you want to sample')
@xw.arg('names', np.array, ndim=1, numbers=str, empty='', doc='The range of names you want to assign')
@xw.arg('nums_to_assign', np.array, ndim=1, numbers=str, empty='', doc='The range of nums for assignment')
@xw.ret(expand='table', transpose=True)
def udf_sample_assign(data, names, nums_to_assign, to_distinct: bool = False, col_index: int = 0):
    """Return a list of boolean values which the num of True is base on data and sample_num"""
    # check the shape of names and nums_to_assign
    if names.shape != nums_to_assign.shape:
        return 'Incorrect input. The num of name to assign is diff from the num of nums_to_assign'
    else:
        # generate a list of name_assign according to names and nums_to_assign
        try:
            # init
            sample_row_num = 0
            name_assign = []

            if names.shape[0] == 1:
                sample_row_num = int(float(nums_to_assign[0]))
                if not sample_row_num > 0:
                    return 'Invaild input value'
                name_assign.extend([names[0]] * sample_row_num)
            else:
                for name, n in zip(names, nums_to_assign):
                    num = int(float(n))
                    sample_row_num += num
                    name_assign.extend([name] * num)
            if sample_row_num > data.shape[0]:
                return 'Error. The sum of nums_to_assign is bigger than the row num of data'
        except ValueError as e:
            log.error('An error occurred. {}'.format(e.args[-1]))
            return 'Invaild input values of nums_to_assign. Failed to convert the value of nums_to_assign to float'

    # drop duplicate values in col labeled col_index
    data_org = pd.DataFrame()
    if bool(to_distinct):
        try:
            col_index = int(float(col_index))
            if col_index not in data.columns:
                return 'col_index out of range while bool(to_distinct) is True'
            else:
                data_org = data.copy()
                data = data.drop_duplicates(subset=[col_index])[col_index]
                print(data[:10])
        except ValueError as e:
            log.error('An error occurred. {}'.format(e.args[-1]))
            return 'Invaild input values of col_index. Failed to convert col_index to float'

    # generate a random array
    rng = default_rng()
    rand_array = rng.random(data.shape[0])

    # use the mask to replace str_array with name_assign
    str_array = np.empty(data.shape[0], dtype=object)
    str_array[rand_array <= np.sort(rand_array)[sample_row_num - 1]] = name_assign

    if not data_org.empty:
        return pd.merge(
            data_org,
            pd.concat([data.reset_index(drop=True), pd.Series(str_array).rename('sampled_names')], axis=1),
            how='outer', on=col_index
        )['sampled_names'].values.tolist()
    else:
        return str_array


if __name__ == '__main__':
    xw.serve()
