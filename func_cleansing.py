#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   func_cleansing.py
@Author  :   Billy Zhou
@Time    :   2021/03/03
@Version :   1.2.0
@Desc    :   None
'''


import logging
import re
import numpy as np
from func_basic import count_sum
from func_basic import dict_append


def array_regex_cleansing(
        data_list, point_exist=True,
        point_digit=2, point_overflow_accept=False, drop_fail=False):
    dict_out = {
        'success': [],
        'failure': [],
    }

    # prepare pattern fo matching
    pattern_plain = """([^-—()（）\d\\n]*((\(|（)|(-|—))?(\d+)[^-—()（）\d\\n]*(\)|）)?(\\n)?)""" # noqa
    # group(0),original
    # group(1): ([^-—()（）\d\\n]*((\(|（)|(-|—))?(\d+)[^-—()（）\d\\n]*(\)|）)?(\\n)?) # noqa
    # group(2),匹配负号或左括号: ((\(|（)|(-|—))
    # group(3),匹配左括号: (\(|（)
    # group(4),匹配负号: (-|—)
    # group(5),匹配数字: (\d+)
    # group(6),匹配右括号: (\)|）)
    # group(7),匹配换行符: (\\n)
    logging.debug("pattern: %s", pattern_plain)
    pattern = re.compile(pattern_plain, re.VERBOSE)

    for line in data_list:
        logging.debug("line: %s", line)
        logging.debug("type(line): %s", type(line))
        if not isinstance(line, (np.str_, str)):
            # cleaned num
            dict_append(dict_out, "Cleaned num", line, line, 'DEBUG')
        else:
            if not line.strip():
                # skip for "\n" row
                dict_append(dict_out, "Blank cell.", "", "", 'DEBUG')
            else:
                logging.debug("Orginal: %s", line.strip())
                result_iter = re.finditer(pattern, line)
                if not result_iter:
                    # output non-matched data
                    dict_append(
                        dict_out, "Non-matched data: %s" % line.strip(),
                        '' if drop_fail else line, line,
                        info_lvl='ERROR')
                else:
                    # cleansing data
                    # capture "-" and num into result_list
                    result_list = []
                    minus_list = []
                    for i, result_i in enumerate(result_iter):
                        logging.debug("result_groups: %s", result_i.groups())

                        # cleansing data - minus_exist
                        if i == 0 and result_i.group(2):
                            logging.debug("group_2: %s" % (result_i.group(2)))
                            minus_list.append(result_i.group(2))

                        if result_i.group(6) and set(
                                ["(", "（"]).intersection(set(minus_list)):
                            logging.debug("group_6: %s" % (result_i.group(6)))
                            minus_list.append(result_i.group(6))

                        group_num = result_i.group(5)
                        logging.debug("group_5,num: %s", group_num)
                        result_list.append(group_num)

                    logging.debug('result_list original: %s' % result_list)
                    # cleansing data - point_exist and point_digit
                    if result_list:
                        if not point_exist or point_digit == 0:
                            if result_list[-1] == "0":
                                result_list.pop()
                        elif point_exist:
                            logging.debug(
                                "length of mantissa: %s", len(result_list[-1]))
                            if len(result_list[-1]) <= point_digit:
                                result_list.pop()
                                result_list.append("." + group_num)
                            else:
                                # point_digit over the threshold
                                logging.debug(
                                    "Num after point overflow: %s",
                                    group_num)
                                if not point_overflow_accept:
                                    logging.warning(
                                        "Overflow! Discard: %s"
                                        % result_list)
                                    result_list = []
                                else:
                                    result_list.pop()
                                    if point_digit >= 0:
                                        result_list.append(
                                            group_num[:-point_digit] + "." +
                                            group_num[-point_digit:]
                                            )
                                    else:
                                        result_list.append(group_num)

                    logging.debug('result_list .cleaned: %s' % result_list)
                    logging.debug('minus_list: %s' % minus_list)
                    # cleansing data - minus_add
                    if minus_list:
                        count_left = count_sum(minus_list, "(", "（")
                        count_right = count_sum(minus_list, ")", "）")
                        if count_sum(minus_list, "-", "—") > 0 or (
                                count_left == 1 and
                                count_left == count_right):
                            result_list.insert(0, "-")

                    logging.debug('result_list -cleaned: %s' % result_list)
                    # output
                    if "".join(result_list):
                        # cleaned data
                        dict_append(
                            dict_out, "Cleaned: %s" % ("".join(result_list)),
                            "".join(result_list), "", info_lvl='DEBUG')
                    else:
                        # uncleaned data while point_overflow_accept = false
                        # Don't append to success whil drop_fail = True
                        dict_append(
                            dict_out, "Uncleaned: %s" % (line.strip()),
                            '' if drop_fail else line, line, info_lvl='ERROR')

    return dict_out['success'], dict_out['failure']


def file_cleansing(
        src_file, dst_file_cleaned, dst_file_uncleaned,
        point_exist=True, point_digit=2,
        point_overflow_accept=False, drop_fail=False):
    data_to_clean = []

    with open(src_file, "r") as f_dirty:
        for line in f_dirty.readlines():
            data_to_clean.append(line)

    if data_to_clean:
        cleaned_data_s, cleaned_data_f = array_regex_cleansing(
            data_to_clean, point_exist, point_digit,
            point_overflow_accept, drop_fail)

        # cleansing the data in src_file and output to dst_files
        with open(dst_file_cleaned, 'w+') as f_cleaned:
            for i in range(len(cleaned_data_s)):
                f_cleaned.write(cleaned_data_s[i] + "\n")
        with open(dst_file_uncleaned, 'w+') as f_uncleaned:
            for i in range(len(cleaned_data_f)):
                f_uncleaned.write(cleaned_data_f[i] + "\n")


if __name__ == '__main__':
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(name)s %(levelname)s %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S')
    logging.debug('start DEBUG')
    logging.debug('==========================================================')

    # test for array_regex_cleansing
    # from test_data.cleansing.test_data import testdata_for_cleansing
    # from test_data.cleansing.test_data import testdata_large_for_cleansing

    # success, failure = array_regex_cleansing(
    #     testdata_for_cleansing, point_exist=True, point_digit=2,
    #     point_overflow_accept=False)

    # logging.info('success: %s' % success)
    # logging.info('failure: %s' % failure)

    # test for file_cleansing
    # from pathlib import Path
    # filepath = Path.cwd().joinpath('test_data\\cleansing')
    # testfile = filepath.joinpath('test_file.txt')
    # logging.info("testfile: %s" % str(testfile))
    # logging.info(testfile.exists())
    # destfile_s = str(filepath.joinpath('test_file_cleaned.txt'))
    # destfile_f = str(filepath.joinpath('test_file_uncleaned.txt'))

    # file_cleansing(
    #     testfile, destfile_s, destfile_f,
    #     point_exist=True, point_digit=4,
    #     point_overflow_accept=True, drop_fail=True)

    logging.debug('==========================================================')
    logging.debug("end DEBUG")
