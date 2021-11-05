#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   common_func.py
@Author  :   Billy Zhou
@Time    :   2020/08/26
@Desc    :   None
'''


import sys
from pathlib import Path
cwdPath = Path(__file__).parents[2]
sys.path.append(str(cwdPath))

from src.manager.LogManager import logmgr  # noqa: E402
log = logmgr.get_logger(__name__)

import pandas as pd
import xlwings as xw
from collections import Counter
from typing import Iterable, List


def count_sum(list_count: list, *args) -> int:
    """sum the count of *args in list_count"""
    result_sum = 0
    dict_count = Counter(list_count)
    for i in args:
        result_sum += dict_count[i]
    log.debug('result_sum: %s', result_sum)
    return result_sum


def walk_path(path_to_walk: str or Path) -> Iterable[Path]:
    """walk over path_to_walk"""
    try:
        for p in Path(path_to_walk).iterdir():
            if p.is_dir():
                yield from walk_path(p)
                continue
            yield p.resolve()
    except NotADirectoryError as e:
        log.error('An error occurred. {!r}'.format(e.args))
        log.error("The path_to_walk[{}] is not a directory".format(path_to_walk))


def search_file_by_type(dirpath: Path, suffix: str = '.json') -> List[str]:
    """search the file with dirpath under filepath"""
    file_list = [str(path) for path in walk_path(dirpath) if path.is_file() and str(path)[-len(suffix):] == suffix]
    log.debug("search_file_list: {}".format(file_list))
    return file_list


def to_csv(suffix: str = ".xlsx"):
    """turn files with suffix to csv files in same directory"""
    wb_path = Path(xw.books.active.fullname).parent
    suffix = "." + suffix if suffix[0] != "." else suffix
    file_list = search_file_by_type(wb_path, suffix)
    log.debug("file_list: {}".format(file_list))
    if file_list:
        for f in file_list:
            log.info("file: {}".format(f))
            if '~$' not in f:
                csv_file_path = wb_path.joinpath(Path(f).name.replace(suffix, ".csv"))
                log.info("to csv_file: {}".format(csv_file_path))
                pd.read_excel(
                    f, index_col=0).to_csv(
                    csv_file_path, encoding='utf-8-sig')
            else:
                log.warning('~$ in file[{}]. Not convert'.format(f))


if __name__ == '__main__':
    # test for count_sum
    log.info("count: %s" % count_sum([1, 2, 2, 2, 3, 5], 2, 5))

    # test for search_file_by_type
    testpath = cwdPath.joinpath('res\\dev\\to_csv')
    search_file_by_type(testpath, '.xlsx')

    # test for to_csv
    import time
    testfile = testpath.joinpath('test_to_csv.xlsx')
    xw.Book(testfile).set_mock_caller()
    to_csv(".xls")
    time.sleep(10)
    to_csv(".xlsx")
