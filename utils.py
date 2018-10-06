#!/usr/bin/python
# -*- coding:utf-8 -*-

"""
  公用工具
"""


def col_index_changer(col_index):
    """
    将如A这类的列索引转换成数字
    :param col_index:
    :return:
    """
    return ord(col_index) - ord("A") + 1
