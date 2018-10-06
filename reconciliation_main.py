#!/usr/bin/python
# -*- coding:utf-8 -*-

"""
  主逻辑
"""

import re
import xlrd
import xlwt
from datetime import datetime
from reconciliation_exception import DateTypeError
from reconciliation_exception import DateFormatError


def datetime_autoformat(excel_time):
    """
    excel时间自适应
    :param time:
    :return:
    """
    date_autoformat = "%Y-%m-%d"

    if isinstance(excel_time, float):
        format_time = xlrd.xldate_as_datetime(excel_time, 0).strftime(date_autoformat)
    elif isinstance(excel_time, basestring):
        date_pattern1 = r"^\d{6,8}$"
        date_pattern2 = r"^\d{4}-\d{1,2}-\d{1,2}$"
        date_format = "%Y%m%d"
        if re.match(date_pattern1, excel_time):
            format_time = datetime.strptime(excel_time, date_format).strftime(date_autoformat)
        elif re.match(date_pattern2, excel_time):
            format_time = datetime.strptime(excel_time, date_autoformat).strftime(date_autoformat)
        elif re.match(date_pattern1, excel_time[:8]):
            format_time = datetime.strptime(excel_time[:8], date_format).strftime(date_autoformat)
        elif re.match(date_pattern2, excel_time[:10]):
            format_time = datetime.strptime(excel_time[:10], date_autoformat).strftime(date_autoformat)
        else:
            raise DateFormatError("错误的时间模式")
    else:
        raise DateTypeError("错误的时间格式 %s" % type(excel_time))

    return format_time


def get_money_dict(work_book, start_row, money_col, time_col, keep_col=[], opposite_flag=False, keep_opposite_flag=True):
    """

    :param work_book:
    :param start_row:
    :param money_col:
    :param time_col:
    :param keep_col:
    :param opposite_flag: 是否取money列的相反数，默认不取
    :param keep_opposite_flag: 是否保留负数部分，默认保留
    :return:
    """
    workbook = xlrd.open_workbook(work_book)
    table = workbook.sheet_by_index(0)

    length = 2 + len(keep_col)
    seq_id = 0
    total_row = table.nrows
    money_dict = {}
    for row in xrange(start_row, total_row):
        data_list = table.row_values(row)
        money = data_list[money_col]
        time = data_list[time_col]

        try:
            time = datetime_autoformat(time)
        except DateTypeError:
            continue
        except DateFormatError:
            continue
        except Exception:
            continue

        if money:
            try:
                money = float(money)
            except ValueError:
                try:
                    money = float(str(money).replace(",", ""))
                except ValueError:
                    continue
            if opposite_flag:   # 去相反数
                money = -money
            if not keep_opposite_flag and money < 0:   # 小于0的数据不保留
                continue
            if money not in money_dict:
                money_dict[money] = []

            seq_id += 1
            keep_data_list = []
            keep_data_list.append("%05d" % seq_id)
            keep_data_list.append(time)
            for col in keep_col:
                keep_data = data_list[col]
                keep_data_list.append(keep_data)
            keep_data_list.append(money)

            money_dict[money].append(keep_data_list)

    return length, money_dict


def match_repeat_money(source_value, target_value, source_len, target_len):
    result = []
    source_result = []
    target_result = []

    # 时间当天匹配
    for source_data_list in source_value:
        source_time = source_data_list[1]
        for target_data_list in target_value:
            target_time = target_data_list[1]
            if source_time[:10] == target_time[:10]:
                result_data_list = []

                # 去掉序列号
                source_data_list.pop(0)
                target_data_list.pop(0)

                result_data_list.extend(source_data_list)
                result_data_list.extend(["", ""])
                result_data_list.extend(target_data_list)

                source_value.remove(source_data_list)
                target_value.remove(target_data_list)
                result.append(result_data_list)
                break

    # 时间非当天匹配的，按照最近时间匹配的原则
    match_list = []
    for source_data_list in source_value:
        for target_data_list in target_value:
            source_time = source_data_list[1]
            target_time = target_data_list[1]

            source_date = source_time[:10]
            target_date = target_time[:10]
            source_date = datetime.strptime(source_date, "%Y-%m-%d")
            target_date = datetime.strptime(target_date, "%Y-%m-%d")
            interval = abs((target_date - source_date).days)   # 时间间隔

            match_list.append([interval, source_data_list, target_data_list])

    # 根据时间间隔来排序
    sorted_match_list = sorted(match_list, key=lambda x: x[0], reverse=False)
    for match_cp in sorted_match_list:
        source_data_list = match_cp[1]
        target_data_list = match_cp[2]

        if source_data_list in source_value and target_data_list in target_value:
            result_data_list = []

            # 去掉序列号
            source_data_list.pop(0)
            target_data_list.pop(0)

            result_data_list.extend(source_data_list)
            result_data_list.extend(["", ""])
            result_data_list.extend(target_data_list)

            source_value.remove(source_data_list)
            target_value.remove(target_data_list)
            result.append(result_data_list)

    if source_value:
        for source_data_list in source_value:
            result_data_list = []

            # 去掉序列号
            source_data_list.pop(0)

            result_data_list.extend(source_data_list)
            result_data_list.extend(["", ""])
            result_data_list.extend([""] * target_len)
            source_result.append(result_data_list)

    if target_value:
        for target_data_list in target_value:
            result_data_list = []

            # 去掉序列号
            target_data_list.pop(0)

            result_data_list.extend([""] * source_len)
            result_data_list.extend(["", ""])
            result_data_list.extend(target_data_list)
            target_result.append(result_data_list)

    return result, source_result, target_result


def write_work_book(table, row, data_list):
    for col, data in enumerate(data_list):
        table.write(row, col, data)


def compare(work_book, source_money_dict, target_money_dict, source_len, target_len, sheet_name):
    row = 0
    table = work_book.add_sheet(sheet_name)

    result = []         # 能够匹配上
    source_result = []  # 只有源
    target_result = []  # 只有目标
    for key, source_value in source_money_dict.items():
        target_value = target_money_dict.get(key, [])

        if len(source_value) > 0 or len(target_value) > 0:
            result_temp, source_result_temp, target_result_temp = match_repeat_money(source_value, target_value, source_len, target_len)
            result.extend(result_temp)
            source_result.extend(source_result_temp)
            target_result.extend(target_result_temp)

        else:
            for source_data_list in source_value:
                result_data_list = []

                # 去掉序列号
                source_data_list.pop(0)

                result_data_list.extend(source_data_list)
                result_data_list.extend(["", ""])
                result_data_list.extend([""] * target_len)
                source_result.append(result_data_list)

    for key, target_value in target_money_dict.items():
        if key not in source_money_dict:
            for target_data_list in target_value:
                result_data_list = []

                # 去掉序列号
                target_data_list.pop(0)

                result_data_list.extend([""] * source_len)
                result_data_list.extend(["", ""])
                result_data_list.extend(target_data_list)
                target_result.append(result_data_list)

    for data_list in result:
        write_work_book(table, row, data_list)
        row += 1

    for data_list in source_result:
        write_work_book(table, row, data_list)
        row += 1

    for data_list in target_result:
        write_work_book(table, row, data_list)
        row += 1


if __name__ == '__main__':
    workbook = xlwt.Workbook()
    source_len, source_money_dict = get_money_dict(u"对账/刷卡做账8月.xlsx", 6, 6, 1, [2, 4, 9])
    target_len, target_money_dict = get_money_dict(u"对账/刷卡账单8月.xlsx", 1, 4, 0, [])
    compare(workbook, source_money_dict, target_money_dict, source_len, target_len, u"做账借方-账单交易金额")

    workbook.save(u"刷卡2.xls")
