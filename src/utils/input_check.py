#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   input_check.py
@Author  :   Billy Zhou
@Time    :   2021/08/22
@Desc    :   None
'''


import getpass


def input_pwd(tip_words='Please input your password:'):
    """Add tip_words to getpass.getpass()"""
    return getpass.getpass(tip_words)


def input_default(default_word='', tip_words='Please input words.'):
    """return default_word while input() return blank"""
    input_data = input(tip_words + '(default: ' + default_word + ')\n')
    if input_data.strip() == '':
        print('Blank input. Using the default value: {0}'.format(default_word))
        return default_word
    else:
        return input_data


def input_checking_list(
        input_list: list,
        tip_words: str = 'Please input words.', case_sens: bool = False,
        min_num: int = 0, default_list: list = ['Y', 'N']):
    """Force input() to return value in input_list, if not in, reinput

    Args:
        input_list: List[str]
            the list to check whether inputed str in
        tip_words: str, default 'Please input words.'
            tip_words added to input()
        case_sens: bool, default False
            whether compare the str in input_list and inputed str in str.upper()
        min_num: int, default 0
            the minimum num of input_list, if less than min_num, use default_list
        default_list: list, default ['Y', 'N']
            if input_list is invaild, use default_list instead
    """
    input_list_str = ''
    # check input_list vaild or not
    if not (type(input_list) == list and len(input_list) > min_num):
        default_list_str = str(default_list)
        print('Invaild input list. Using the default list of ' + default_list_str + '.')
        input_list = default_list

    # construct the tip_words
    for num, value in enumerate(input_list):
        if num == 0:
            input_list_str = '[' + value + ']'
            default_value = value
        else:
            input_list_str = input_list_str + '/' + value
    tip_words = tip_words + '(' + input_list_str + '): '

    if case_sens:
        input_value = input_default(default_value, tip_words)
        while not (set([input_value]) & set(input_list)):
            print('Unexpect input! Please input words in ' + input_list_str + '.')
            input_value = input_default(default_value, tip_words)
    else:
        input_value = input_default(default_value.upper(), tip_words).upper()
        while not (set([input_value]) & set([i.upper() for i in input_list])):
            print('Unexpect input! Please input words in ' + input_list_str + '.')
            input_value = input_default(default_value.upper(), tip_words).upper()

    return input_value


def input_checking_YN(tip_words='Please input words.', default_Y=True):
    """Force input() to return value in ['Y', 'N'], if not in, reinput"""
    input_list = ['Y', 'N']
    if not default_Y:
        input_list = ['N', 'Y']
    return input_checking_list(input_list, tip_words, case_sens=False)


if __name__ == "__main__":
    import logging
    logging.basicConfig(level=logging.NOTSET)

    logging.info("return : {}".format(input_default('abc')))
    logging.info("return : {}".format(input_pwd()))
    logging.info("return : {}".format(input_checking_YN()))
    logging.info("return : {}".format(input_checking_list(['a', 'b', 'c', 'd'])))
    logging.info("return : {}".format(input_checking_list(['a'])))
    logging.info("return : {}".format(input_checking_list([])))
    logging.info("return : {}".format(input_checking_list('a')))
