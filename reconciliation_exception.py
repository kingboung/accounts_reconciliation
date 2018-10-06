#!/usr/bin/python
# -*- coding:utf-8 -*-

"""
  自定义异常类
"""

from exceptions import Exception


class DateTypeError(Exception):
    def __init__(self, msg):
        Exception.__init__(self)
        self.error_msg = msg

    def __str__(self):
        return self.error_msg


class DateFormatError(Exception):
    def __init__(self, msg):
        Exception.__init__(self)
        self.error_msg = msg

    def __str__(self):
        return self.error_msg
