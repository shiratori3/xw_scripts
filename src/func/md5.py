#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   md5.py
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

import hashlib
import numpy as np
import xlwings as xw


# @xw.arg('data', 'range')
@xw.func
@xw.arg('data', ndim=2, numbers=str, empty='')
def udf_md5_single(data):
    data_joined = ''.join(list(np.array(data).flatten()))
    return hashlib.md5(data_joined.encode(encoding='UTF-8')).hexdigest()


@xw.func
@xw.arg('data', ndim=2, numbers=str, empty='')
@xw.ret(expand='table')
def udf_md5(data):
    data_array = np.array(data)
    md5_data_list = list(map(lambda x: hashlib.md5(x.encode(encoding='UTF-8')).hexdigest(), data_array.flatten()))
    return np.array(md5_data_list).reshape(data_array.shape)


if __name__ == '__main__':
    xw.serve()
