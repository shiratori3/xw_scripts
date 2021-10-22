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
@xw.arg('names', np.array, ndim=1, numbers=str, empty='', doc='Select the range of names you want to arrange for')
@xw.arg('nums_to_arrange', np.array, ndim=1, numbers=str, empty='', doc='Select the range of arrange nums')
@xw.ret(expand='table', transpose=True)
def udf_arrange(names, nums_to_arrange):
    if names.shape != nums_to_arrange.shape:
        return 'Incorrect input. The num of name to arrange is diff from the num of nums_to_arrange'
    else:
        try:
            if names.shape[0] == 1:
                if int(float(nums_to_arrange[0])) > 0:
                    return [[name] * int(float(n)) for name, n in zip(names, nums_to_arrange)][0]
                else:
                    return 'Invaild input value'
            else:
                res = []
                for name, n in zip(names, nums_to_arrange):
                    res.extend([name] * int(float(n)))
                return res
        except ValueError as e:
            log.error('An error occurred. {}'.format(e.args[-1]))
            return 'Invaild input values. Failed to convert the nums_to_arrange to float'


@xw.func
@xw.arg('names', np.array, ndim=1, numbers=str, empty='', doc='Select the range of names you want to arrange for')
@xw.arg('tasks', np.array, ndim=1, numbers=str, empty='', doc='Select the range of tasks you want to arrange for')
@xw.arg('nums_to_arrange', np.array, ndim=2, empty=0, transpose=True, doc='Select the range of arrange nums')
@xw.ret(expand='table', transpose=True)
def udf_arrange_multi(names, tasks, nums_to_arrange):
    if names.shape[0] * tasks.shape[0] != nums_to_arrange.shape[0] * nums_to_arrange.shape[1]:
        return 'Incorrect input. The product of the num of names and tasks is diff from the num of nums_to_arrange'
    else:
        try:
            if np.any(nums_to_arrange):
                res = [[], []]
                for no_task, task in enumerate(tasks):
                    for no_name, name in enumerate(names):
                        log.info(nums_to_arrange[no_task][no_name])
                        log.info(type(nums_to_arrange[no_task][no_name]))
                        res[0].extend([task] * int(float(nums_to_arrange[no_task][no_name])))
                        res[1].extend([name] * int(float(nums_to_arrange[no_task][no_name])))
                return res
            else:
                return 'Invaild input values of nums_to_arrange which is all zero'
        except ValueError as e:
            log.error('An error occurred. {}'.format(e.args[-1]))
            return 'Invaild input values. Failed to convert the nums_to_arrange to float'


if __name__ == '__main__':
    xw.serve()
