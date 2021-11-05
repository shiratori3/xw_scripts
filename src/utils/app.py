#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   app_func.py
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


import xlwings as xw
from functools import wraps


def stop_refresh():
    """stop the screen update of excel"""
    xw.apps.active.screen_updating = False


def start_refresh():
    """start the screen update of excel if stoped"""
    if not xw.apps.active.screen_updating:
        xw.apps.active.screen_updating = True


def try_expect_stop_refresh(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            stop_refresh()
            log.info("refresh stoped")
            result = func(*args, **kwargs)
            log.info("Function handle over")
            return result
        except Exception as e:
            log.error('An error occurred. {}'.format(e.args))
        finally:
            start_refresh()
            log.info("refresh restart")
    return wrapper


if __name__ == '__main__':
    # test for search_file_by_type
    import time
    testpath = cwdPath.joinpath('res\\dev\\to_csv')
    testfile = testpath.joinpath('test_to_csv.xlsx')
    xw.Book(testfile).set_mock_caller()
    stop_refresh()
    time.sleep(10)
    start_refresh()
