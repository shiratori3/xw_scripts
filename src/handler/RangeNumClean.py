#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   RangeNumClean.py
@Author  :   Billy Zhou
@Time    :   2021/08/30
@Desc    :   None
'''


import sys
from pathlib import Path
cwdPath = Path(__file__).parents[2]  # the num depend on your filepath
sys.path.append(str(cwdPath))

from src.manager.LogManager import logmgr
log = logmgr.get_logger(__name__)

import re
import numpy as np
from typing import Any
from xlwings import Range
from decimal import Decimal, ROUND_DOWN
from src.utils.basic import count_sum
from src.utils.cell import str2num
from src.handler.BaseCellHandler import break_where
from src.handler.BaseCellHandler import BaseCellHandler
from src.handler.BaseRangeHandler import BaseRangeHandler


class CellNumClean(BaseCellHandler):
    """change the sign of a numberic cell

    Attrs:
    -----
        cell_value: Any
            the value of cell
        pattern:
            the pattern for clean method to finditer in cell_value
        num_formatter: Decimal
            the number formatter for cell_value
        digit_div: Decimal
            the divisor for cleaned num
        truncate: bool
            whether to truncate the num after required digits

    Methods:
    -------
        in_supported_type:
            check whether can change the sign of cell value directly
        not_only_backslash:
            return break while value equal to \\n
        value2deciaml:
            try to convert the value to deciaml
        clean:
            clean the value of cell
    """
    def __init__(self, cell_value: Any, pattern, num_formatter: Decimal, digit_div: Decimal, truncate: bool, cell_format: str = ''):
        super().__init__(cell_value, cell_format=cell_format)
        # prepare pattern fo matching
        self.pattern = pattern

        # required number of places after the decimal point
        self.digit_div = digit_div
        self.num_formatter = num_formatter
        self.truncate = truncate

        # whether cleaned
        self.cleaned = False

    @break_where
    def in_supported_type(self, supported_type: tuple = (np.str_, str)):
        """check whether need to clean the cell value"""
        if self.breaked:
            return self
        if isinstance(self.cell_value, supported_type):
            log.debug('the type of cell value: {}'.format(
                type(self.cell_value)))
            return self
        else:
            # not str or np.str_, not need to clean
            self.breaked = True
            return self

    @break_where
    def not_only_backslash(self):
        """return break while cell_value equal to \\n"""
        if self.breaked:
            return self
        try:
            if self.cell_value.strip():
                self.cell_value = self.cell_value.strip()
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

    @break_where
    def value2deciaml(self):
        """try to convert the value to deciaml

        The cleaned data will divide by digit_div, which is 10 ** point_digit.
        Also the cell_value will be format to {:,.Nf}, which N is point_digit
        """
        if self.breaked:
            return self
        self.cell_value = str2num(self.cell_value)
        if not isinstance(self.cell_value, Decimal):
            # need to clean
            return self
        else:
            if self.cleaned:
                self.cell_value = self.cell_value / self.digit_div
            # format the value and break
            if self.truncate:
                self.cell_value = self.cell_value.quantize(self.num_formatter, rounding=ROUND_DOWN).to_eng_string()
            else:
                self.cell_value = self.cell_value.quantize(self.num_formatter).to_eng_string()
            self.breaked = True
            return self

    @break_where
    def clean(self):
        """clean the value of cell

        Capture all nums in the string and check whether minus sign exists, if exists, then
        add it to the begin of result_list.
        Return a string of joined result_list as result"""
        if self.breaked:
            return self
        # clean data
        result_iter = re.finditer(self.pattern, self.cell_value)
        if result_iter:
            # capture minus_sign and num
            result_list = []
            minus_list = []
            for i, result_i in enumerate(result_iter):
                log.debug("result_groups: {}".format(result_i.groups()))

                # capture minus sign: - or — or () or （）
                if i == 0 and result_i.group(2):
                    minus_list.append(result_i.group(2))
                if result_i.group(6) and set(["(", "（"]) & set(minus_list):
                    minus_list.append(result_i.group(6))

                # capture nums
                result_list.append(result_i.group(5))
            log.debug("result_list original: {}".format(result_list))

            # add minus sign to result_list
            if minus_list:
                count_left = count_sum(minus_list, "(", "（")
                count_right = count_sum(minus_list, ")", "）")
                if count_sum(minus_list, "-", "—") > 0 or (
                        count_left == count_right == 1):
                    result_list.insert(0, "-")

            # discard the digit infomation in cleaning
            self.cell_value = "".join(result_list)
            self.cleaned = True
            return self
        else:
            self.breaked = True
            return self


class RangeNumClean(BaseRangeHandler):
    """change the sign of numberic cells in a selected range

    Args:
    ----
        rang: str or Range, default None
            the address of range to handler.
            If none, use the xw.apps.active.selection
        pre_formatter: str
            whether to format the cells in rang before handling
        point_digit: int
            the digit to keep after point.
            10 ** point_digit is the divisor for cleaned num '12212' of dirty data '1.2S2,12'
    """
    def __init__(
            self, rang: str or Range = '',
            pre_formatter: str = '', point_digit: int = 2, truncate: bool = False) -> None:
        super().__init__(rang=rang, pre_formatter=pre_formatter)

        if int(point_digit) < 0:
            raise ValueError('Invaild point_digit[{}] for RangeNumClean'.format(point_digit))
        self.point_digit = int(point_digit)
        self.truncate = truncate

    def cell_handler(self):
        """try to clean the value to number for each cell in a selected range"""
        # prepare pattern fo matching
        pattern_plain = """([^-—()（）\d\\n]*((\(|（)|(-|—))?(\d+)[^-—()（）\d\\n]*(\)|）)?(\\n)?)"""  # noqa: W605
        """group(0): original
        group(1): ([^-—()（）\d\\n]*((\(|（)|(-|—))?(\d+)[^-—()（）\d\\n]*(\)|）)?(\\n)?)
        group(2): 匹配负号或左括号: ((\(|（)|(-|—))
        group(3): 匹配左括号: (\(|（)
        group(4): 匹配负号: (-|—)
        group(5): 匹配数字: (\d+)
        group(6): 匹配右括号: (\)|）)
        group(7): 匹配换行符: (\\n)
        """  # noqa: W605
        pattern = re.compile(pattern_plain, re.VERBOSE)

        # prepare digit format
        num_formatter = Decimal(10) ** -self.point_digit
        digit_div = Decimal(10) ** self.point_digit

        for cell_value in self.rang_np_list:
            cell = CellNumClean(cell_value, pattern, num_formatter, digit_div, self.truncate)
            cell.in_supported_type().not_only_backslash().value2deciaml().clean().value2deciaml()
            self.result_list.append(cell.cell_value)


if __name__ == '__main__':
    testdata_for_clean = [
        "35,396,293.85", "964,978,645.05", "5.,392.224.39", " 1,591 ,250.00 ", " 2, 100,000.33 ", " 3,791, 128.16 ", " 379, ,560.00 ", " -787, 500.00 ", " (682, 500.00 )", " 700,000. 12 ", " 2,88--9,482. .50| ", " 14,995,796.67 ", " 518,545.44 ", "", " 931 ,560 24 ", " 295,053. 19 ", " 23 1,000.00 ", " 1,599,999 96 ", " 381 ,259,440.31 ", " .513,191.49 ", " 1,120,583.33| ", "", " 3,156,862.75| ", " 262, 500.00 ", " 166,666.32| ", " 2,999 998.00 ", "", " 20,259,072 .20| ", " 401 ,518,555,555,555,512.  51| ", " 1.00 |", " 8888.8888", " 850.214.10554", " 85010554", "", "sssssss", "", "", ""]
    testdata_large_for_clean = [
        "35,396,293.85", "964,978,645.05", "5.,392.224.39", "132,948.003.96", "77,809.480.46", "104,104.749.74", "89,903,364.41", "850.214.10554", "3.365,228.932.10", "262.711.437.72", "78.023.386.44", "90.006,558.96", "29.308,153.96", "3,908.479.60", "20.960,002.39", "48,683.915.80", "18.817,893.16", "19,714,345,604.00", "92.023.333.22", "62,307.555.50", "7,122.814.00", "35,415,665.85", "52.549.874.38", "154.075.529.73", "29.379.098.45", "8.684.800.37", "1,729.916.33", "12.527,055.30", "598,256.962.12", "7,429,611,688.82", " 65,232,565.78 ", " 1,006,268.448.63 ", " 16.254.184.60 ", " 152,780,195,17 ", " 67,075,129.57 ", " 78.011,195,71 ", " 84,492,811.28 ", " 636,714.076.53 ", " 3,006,422.142.96 ", " 289.619.756.87 ", " 66.688.222.39 ", " 92,001,042.68 ", " 21.812,696.78 ", " 9,964,349.72 ", " 18.617,163.67 ", " 62.437,644.46 ", " 27,755.813.92 ", " 259.040,159.69 ", " 95.688.281.06 ", " 65,378,665.73 ", " 8,059,361,12 ", " 47,256,002.23 ", " 66,091.533.70 ", " 124.386.207.98 ", " 5.316,024.97 ", " 27,622.239.29 ", " 4,773.420.16 ", " 11.438.446.54 ", " 475,790,181.41 ", " 6,892,987,964.60 ", " 1,100,000.00 ", " 5,850,000.00 ", " 11,151,681.84 ", " 8,676,047.55 ", " 24,000,000.08 ", " 760,000.08 ", " 6,000,000.00 ", " 51,667.00 ", " 1,591 ,250.00 ", " 408,333.00 ", " 165,000.00 ", " 813,333.00 ", " 2,738,007.67 ", " 329,999.67 ", " 5,000,500.00 ", " 2,060,000.00 ", " 280,000.00 ", " 2,581,250.00 ", " 933,333.33 ", " 280,000.00 ", " 2, 100,000.33 ", " 9,441,000.00 ", " ssss", "", "", " faskfajjwr", " fauyfahnfh", " 6,435,439.52 ", " 3,791, 128.16 ", " 30,000.00 ", "", " 180,000.00 ", " 180,000.00 ", " 379, ,560.00 ", " 960,000.00 ", " 4,549,999.92 ", " 666,672.64 ", " 1,049,337.00 ", " 787, 500.00 ", " 8,799,999.84 ", " 364,800.00 ", " 682, 500.00 ", " 1,150,000.00 ", " 700,000. 12 ", " 206,250.00 ", " 1,958,976.00 ", " 4,426,240.00 ", " 2,889,482. .50| ", " 14,995,796.67 ", " 90,000.00 ", " 18,460,000.00 ", "", " 300,000.00 ", " 53,646,000.00 ", " 1,533,334.00 ", " 200,000.00 ", " 1,482,180.61 ", " 518,545.44 ", "", "", "", "", "", "", "", "", " 931 ,560 24 ", " 1,600,000.00 ", " 1,920,000.00 ", " 400,000.00 ", " 120,000.00 ", " 960,000.00 ", " 560,000.00 ", " 2,400,000.00 ", " 240,000.00 ", " 4,516,060.73 ", " 295,053. 19 ", "", "", " 3,750,000.00 ", " 5,250,000.00 ", " 550,000.00 ", " 3,200,000.00 ", " 23 1,000.00 ", " 9,456,060.22 ", " 1,599,999 96 ", "", " 166,750.00 ", " 250,000.00 ", " 525,000.00 ", " 760,000.00 ", " 1,555,000.00 ", " 25,500,000.00 ", " 200,000.00 ", "", " 185,000.00 ", " 5,000.000.00 ", "", " 8,121,000.00 ", "", "", " 560,000.00 ", " 4,500,000.00 ", " 30,000,000.00 ", " 500,000.00 ", " 9,802,260.00 ", " 2,919,550.00 ", " 11,140,000.00 ", " 200,000.00 ", "", "", "", " 381 ,259,440.31 ", "", " 2,782,698.20 ", " 2,170,000.00 ", " .513,191.49 ", " 2,120,689.66 ", " 150,000.00 ", " 900,000.00 ", " 245,000.00 ", " 445,205.00 ", " 300,000.00 ", " 475,000.00 ", " 1,120,583.33| ", "", "", "", "", "", "", "", " 25,000.00 ", " 3,156,862.75| ", " 500,000.00 ", " 262, 500.00 ", " 310,000.00 ", " 15,677.46 ", " 166,666.32| ", " 30,000.00 ", " 19,999.99 ", " 2,999 998.00 ", "\n", " 1,000,000.00 ", " 100,000.00 ", "", "", "", " 20,259,072 .20|", " 401 ,51www8,555,555,555,512. 51| ", " 1.00 |", " 8888.8888", " 850.214.10554", ""]
    if True:
        # test for CellNumClean
        pattern_plain = """([^-—()（）\d\\n]*((\(|（)|(-|—))?(\d+)[^-—()（）\d\\n]*(\)|）)?(\\n)?)"""  # noqa: W605
        pattern = re.compile(pattern_plain, re.VERBOSE)
        point_digit = 2
        num_formatter = str(point_digit).join(['{:,.', 'f}'])
        digit_div = Decimal(10) ** point_digit
        if False:
            for num in testdata_for_clean:
                log.info('unclean: {}'.format(num))
                cell = CellNumClean(num, pattern, num_formatter, digit_div)
                cell.in_supported_type().not_only_backslash().value2deciaml().clean().value2deciaml()
                log.info('cleaned: {}'.format(cell.cell_value))
        if False:
            for num in testdata_large_for_clean:
                log.info('unclean: {}'.format(num))
                cell = CellNumClean(num, pattern, num_formatter, digit_div)
                cell.in_supported_type().not_only_backslash().value2deciaml().clean().value2deciaml()
                log.info('cleaned: {}'.format(cell.cell_value))
