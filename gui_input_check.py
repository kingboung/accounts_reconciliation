#!/usr/bin/python
# -*- coding:utf-8 -*-

"""
  校验界面用户输入是否正确
"""

from utils import col_index_changer
from reconciliation_main import datetime_autoformat


def check_excel_file_input(excel_file_path):
    if excel_file_path.endswith(".xls") or excel_file_path.endswith(".xlsx"):
        return True
    return False


def check_row_input(workbook_table, row_number):
    row_total = workbook_table.nrows
    row_number = int(row_number)
    if row_number > row_total or row_number < 1:
        return False
    else:
        return True


def check_col_input(workbook_table, col_index):
    col_number = col_index_changer(col_index)
    col_total = workbook_table.ncols
    if col_number > col_total:
        return False
    else:
        return True


def check_col_list_input(workbook_table, col_list_str):
    if not col_list_str:
        return True

    col_total = workbook_table.ncols
    col_list = col_list_str.rstrip(",").split(",")
    for col_index in col_list:
        col_number = col_index_changer(col_index)
        if col_number > col_total:
            return False

    return True


def check_date_format_fit(workbook_table, start_row, date_col):
    row = int(start_row) - 1
    col = col_index_changer(date_col) - 1

    date = workbook_table.cell_value(row, col)
    try:
        datetime_autoformat(date)
        return True
    except Exception:
        return False


def check_money_format(workbook_table, start_row, money_col):
    row = int(start_row) - 1
    col = col_index_changer(money_col) - 1

    money = workbook_table.cell_value(row, col)
    if money == "":
        return True

    try:
        float(money)
    except ValueError:
        try:
            float(str(money).replace(",", ""))
        except ValueError:
            return False

    return True
