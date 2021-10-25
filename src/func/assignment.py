#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   arrange.py
@Author  :   Billy Zhou
@Time    :   2021/10/22
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


@xw.func
@xw.arg('names', np.array, ndim=1, numbers=str, empty='', doc='The range of names you want to assign')
@xw.arg('nums_to_assign', np.array, ndim=1, numbers=str, empty='', doc='The range of nums for assignment')
@xw.ret(expand='table', transpose=True)
def udf_assign(names, nums_to_assign):
    """Return a list of names"""
    if names.shape != nums_to_assign.shape:
        return 'Incorrect input. The shape of names is diff from the num of nums_to_assign'
    else:
        try:
            if names.shape[0] == 1:
                if int(float(nums_to_assign[0])) > 0:
                    return [[name] * int(float(n)) for name, n in zip(names, nums_to_assign)][0]
                else:
                    return 'Invaild input value of nums_to_assign'
            else:
                res = []
                for name, n in zip(names, nums_to_assign):
                    res.extend([name] * int(float(n)))
                return res
        except ValueError as e:
            log.error('An error occurred. {}'.format(e.args[-1]))
            return 'Invaild input values. Failed to convert the value of nums_to_assign to float'


@xw.func
@xw.arg('names', np.array, ndim=1, numbers=str, empty='', doc='The range of names you want to assign')
@xw.arg('tasks', np.array, ndim=1, numbers=str, empty='', doc='The range of tasks you want to assign')
@xw.arg('nums_to_assign', np.array, ndim=2, empty=0, transpose=True, doc='The range of nums of names and tasks for assignment')
@xw.ret(expand='table', transpose=True)
def udf_assign_multi(names, tasks, nums_to_assign):
    """Return an array of names and tasks"""
    if names.shape[0] * tasks.shape[0] != nums_to_assign.shape[0] * nums_to_assign.shape[1]:
        return 'Incorrect input. The product of the shape of names and tasks is diff from the total num of nums_to_assign'
    else:
        try:
            if np.any(nums_to_assign):
                res = [[], []]
                for no_task, task in enumerate(tasks):
                    for no_name, name in enumerate(names):
                        res[0].extend([task] * int(float(nums_to_assign[no_task][no_name])))
                        res[1].extend([name] * int(float(nums_to_assign[no_task][no_name])))
                return res
            else:
                return 'Invaild input value of nums_to_assign which is all zero'
        except ValueError as e:
            log.error('An error occurred. {}'.format(e.args[-1]))
            return 'Invaild input values. Failed to convert the value of nums_to_assign to float'


if __name__ == '__main__':
    xw.serve()
