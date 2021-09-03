#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   CellHandler.py
@Author  :   Billy Zhou
@Time    :   2021/08/26
@Desc    :   None
'''


import sys
from pathlib import Path
cwdPath = Path(__file__).parents[2]
sys.path.append(str(cwdPath))

from src.manager.LogManager import logmgr
log = logmgr.get_logger(__name__)

import numpy as np
from typing import Any
from functools import wraps, partial


def break_where(func, *, break_info: bool = False):
    """A decorator to check whether the CellHandler breaked at first time"""
    if func is None:
        return partial(break_where, break_info=break_info)

    @wraps(func)
    def wrapper(self, *args, **kwargs):
        if break_info:
            if not getattr(self, 'break_checked', False) and getattr(self, 'breaked'):
                log.info('cell_value[{!r}] breaked before func[{}()]'.format(
                    getattr(self, 'cell_value'), func.__name__))
                setattr(self, 'break_checked', True)
        return func(self, *args, **kwargs)
    return wrapper


class BaseCellHandler(object):
    """The base class for CellHandler

    Base class for Cellhandler.
    You can add other methods to handle with the cell value and format.
    Two base handlers are provided.
    These methods work in a fluent style, a function call like \
        cell.in_supported_type((np.str_, str)) \
        .not_only_backslash() is supported

    Attrs:
    -----
        value: Any
            the value of a cell instance
        format: str

    Methods:
    -------
        in_supported_type:
            check the type of value in supported_type
        not_only_backslash:
            return break while value equal to \\n

    Examples:
    --------
        cell = BaseCellHandler(100)
        cell.not_only_backslash()

        cell
    """
    def __init__(self, cell_value: Any, cell_format: str = ''):
        self.cell_value = cell_value
        self.cell_format = cell_format
        self.breaked = False

    @break_where
    def in_supported_type(self, supported_type: tuple):
        """check the type of cell_value in supported_type"""
        if self.breaked:
            return self
        if isinstance(self.cell_value, supported_type):
            return self
        else:
            log.debug('the type of cell_value: {}'.format(
                type(self.cell_value)))
            self.breaked = True
            return self

    @break_where
    def not_only_backslash(self):
        """return break while cell_value equal to \\n"""
        if self.breaked:
            return self
        try:
            if self.cell_value.strip():
                return self
            else:
                self.breaked = True
                return self
        except AttributeError as e:
            log.error('An error occurred. {}'.format(e.args))
            log.error(
                'invaild type for method not_only_backslash: {}'.format(
                    type(self.cell_value)))
            self.breaked = True
            return self


if __name__ == '__main__':
    # test for BaseCellHandler
    cell1 = BaseCellHandler(500000)
    cell1.not_only_backslash().in_supported_type((np.str_, str))
    log.info("cell1.cell_value: {}".format(cell1.cell_value))

    cell2 = BaseCellHandler('\n')
    cell2.not_only_backslash().in_supported_type((np.str_, str))
    log.info("cell2.cell_value: {}".format(cell2.cell_value))

    cell2_diff_order = BaseCellHandler('\n')
    cell2_diff_order.in_supported_type((np.str_, str)).not_only_backslash()
    log.info("cell2_diff_order.cell_value: {}".format(cell2_diff_order.cell_value))

    cell3 = BaseCellHandler('500000\n')
    cell3.not_only_backslash().in_supported_type((np.str_, str))
    log.info("cell3.cell_value: {}".format(cell3.cell_value))
