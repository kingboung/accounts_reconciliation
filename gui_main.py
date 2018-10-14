#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
  主界面
"""

import os
import re
import xlrd
import xlwt
from Tkinter import Tk
from Tkinter import Label
from Tkinter import Entry
from Tkinter import Button
from Tkinter import StringVar
from Tkinter import BooleanVar
from Tkinter import Checkbutton
from tkFileDialog import askopenfilename
from tkFileDialog import asksaveasfilename
from gui_input_check import check_excel_file_input
from gui_input_check import check_row_input
from gui_input_check import check_col_input
from gui_input_check import check_col_list_input
from gui_input_check import check_date_format_fit
from gui_input_check import check_money_format
from reconciliation_main import get_money_dict
from reconciliation_main import compare
from utils import col_index_changer

gui_main = Tk()
gui_main.title("对账小程序")

# 界面居中显示
screen_width = gui_main.winfo_screenwidth()
screen_height = gui_main.winfo_screenheight()
cursor_width = 800
cursor_height = 320

gui_main.geometry("{}x{}+{}+{}".format(cursor_width,
                                       cursor_height,
                                       screen_width / 2 - cursor_width / 2,
                                       screen_height / 2 - cursor_height / 2))
gui_main.resizable(False, False)    # 主窗口大小不可调整


def validate_row_input(content):
    if re.match("^\d*$", content):
        return True
    else:
        return False


def validate_col_input(content):
    if re.match("^[A-Z]{0,1}$", content):
        return True
    else:
        return False


def validate_col_list_input(content):
    if re.match("^[A-Z]{1}(,[A-Z]{1})*[,]?$", content) or content == "":
        return True
    else:
        return False


def choose_file(file_type):
    file_path = askopenfilename(filetypes=[('all excel files', '*.xls?'), ('excel file1', '*.xls'), ('excel file2', '*.xlsx')])
    if file_type == "people" and file_path:
        people_file_path.set(file_path)
    if file_type == "bank" and file_path:
        bank_file_path.set(file_path)


def save_file():
    file_name = asksaveasfilename(defaultextension=".xls", initialfile="对账.xls")
    return file_name


def check_input_main():
    bank_excel_file_input = bank_file_path.get()
    people_excel_file_input = people_file_path.get()
    bank_start_row_input = bank_start_row.get()
    people_start_row_input = people_start_row.get()
    bank_money_col_input = bank_money_col.get()
    people_money_col_input = people_money_col.get()
    bank_time_col_input = bank_time_col.get()
    people_time_col_input = people_time_col.get()
    bank_keep_cols_input = bank_keep_cols.get()
    people_keep_cols_input = people_keep_cols.get()

    # 空检查
    if not bank_excel_file_input:
        error_msg.set("尚未选择账单文件，请点击选择账单文件")
        return False
    if not people_excel_file_input:
        error_msg.set("尚未选择做账文件，请点击选择账单文件")
        return False
    if not bank_start_row_input:
        error_msg.set("账单文件 对账起始行 不能为空")
        return False
    if not people_start_row_input:
        error_msg.set("做账文件 对账起始行 不能为空")
        return False
    if not bank_money_col_input:
        error_msg.set("账单文件 交易金额所在列 不能为空")
        return False
    if not people_money_col_input:
        error_msg.set("做账文件 借方金额/贷方金额所在列 不能为空")
        return False
    if not bank_time_col_input:
        error_msg.set("账单文件 日期所在列 不能为空")
        return False
    if not people_time_col_input:
        error_msg.set("做账文件 日期所在列 不能为空")
        return False

    # 文件存在检查
    if not os.path.exists(bank_excel_file_input):
        error_msg.set("所选账单文件不存在，可能已被删除！请重新选择")
        return False
    if not os.path.exists(people_excel_file_input):
        error_msg.set("所选做账文件不存在，可能已被删除！请重新选择")
        return False

    # 文件后缀名检查
    if not check_excel_file_input(bank_excel_file_input):
        error_msg.set("所选账单文件不是excel文件！请重新选择")
        return False
    if not check_excel_file_input(people_excel_file_input):
        error_msg.set("所选做账文件不是excel文件！请重新选择")
        return False

    bank_workbook = xlrd.open_workbook(bank_excel_file_input)
    bank_table = bank_workbook.sheet_by_index(0)
    people_workbook = xlrd.open_workbook(people_excel_file_input)
    people_table = people_workbook.sheet_by_index(0)

    # 行/列是否超出文件最大行/列检查
    if not check_row_input(bank_table, bank_start_row_input):
        error_msg.set("账单文件 对账起始行 超出文件最大行数或小于1")
        return False
    if not check_col_input(bank_table, bank_money_col_input):
        error_msg.set("账单文件 交易金额所在列 超出文件最大列数")
        return False
    if not check_col_input(bank_table, bank_time_col_input):
        error_msg.set("账单文件 日期所在列 超出文件最大列数")
        return False
    if not check_col_list_input(bank_table, bank_keep_cols_input):
        error_msg.set("账单文件 需要保留的列 某一列超出文件最大列数")
        return False

    if not check_row_input(people_table, people_start_row_input):
        error_msg.set("做账文件 对账起始行 超出文件最大行数或小于1")
        return False
    if not check_col_input(people_table, people_money_col_input):
        error_msg.set("做账文件 借方金额/贷方金额所在列 超出文件最大列数")
        return False
    if not check_col_input(people_table, people_time_col_input):
        error_msg.set("做账文件 日期所在列 超出文件最大列数")
        return False
    if not check_col_list_input(people_table, people_keep_cols_input):
        error_msg.set("做账文件 需要保留的列 某一列超出文件最大列数")
        return False

    # 所选日期是否符合系统能够处理的日期格式检查
    if not check_date_format_fit(bank_table, bank_start_row_input, bank_time_col_input):
        error_msg.set("账单文件 日期所在列 不是系统能够适配的日期格式")
        return False
    if not check_date_format_fit(people_table, people_start_row_input, people_time_col_input):
        error_msg.set("做账文件 日期所在列 不是系统能够适配的日期格式")
        return False

    # 所选金额是否符合金额格式检查
    if not check_money_format(bank_table, bank_start_row_input, bank_money_col_input):
        error_msg.set("账单文件 交易金额所在列 无法正常处理，请检查")
        return False
    if not check_money_format(people_table, people_start_row_input, people_money_col_input):
        error_msg.set("做账文件 借方金额/贷方金额所在列 无法正常处理，请检查")
        return False

    error_msg.set("")
    return True


def reconciliation_main():
    bank_excel_file_path = bank_file_path.get()
    people_excel_file_path = people_file_path.get()

    bank_start_row_number= int(bank_start_row.get()) - 1
    people_start_row_number = int(people_start_row.get()) - 1

    bank_money_col_number = col_index_changer(bank_money_col.get()) - 1
    people_money_col_number = col_index_changer(people_money_col.get()) - 1

    bank_time_col_number = col_index_changer(bank_time_col.get()) - 1
    people_time_col_number = col_index_changer(people_time_col.get()) - 1

    bank_keep_col_list = []
    bank_keep_cols_input = bank_keep_cols.get()
    if bank_keep_cols_input:
        bank_keep_col_list = bank_keep_cols_input.rstrip(",").split(",")
        bank_keep_col_list = [col_index_changer(item) - 1 for item in bank_keep_col_list]

    people_keep_col_list = []
    people_keep_cols_input = people_keep_cols.get()
    if people_keep_cols_input:
        people_keep_col_list = people_keep_cols_input.rstrip(",").split(",")
        people_keep_col_list = [col_index_changer(item) - 1 for item in people_keep_col_list]

    source_len, source_money_dict = get_money_dict(bank_excel_file_path, bank_start_row_number, bank_money_col_number, bank_time_col_number, bank_keep_col_list, bank_opposite_flag.get(), bank_keep_opposite_money_flag.get())
    target_len, target_money_dict = get_money_dict(people_excel_file_path, people_start_row_number, people_money_col_number, people_time_col_number, people_keep_col_list, people_opposite_flag.get(), people_keep_opposite_money_flag.get())

    workbook = xlwt.Workbook()
    if sheet_name.get():
        compare(workbook, source_money_dict, target_money_dict, source_len, target_len, sheet_name.get())
    else:
        compare(workbook, source_money_dict, target_money_dict, source_len, target_len, "Sheet1")

    file_name = save_file()
    if file_name:
        workbook.save(file_name)


def start_reconciliation():
    sure_button["state"] = "disabled"
    if check_input_main():
        reconciliation_main()
    sure_button["state"] = "normal"


def main():
    # ----------选择文件----------
    global bank_file_path
    bank_file_path = StringVar()
    Label(gui_main, text="账单文件：").place(x=20, y=25, width=70, height=25)
    Entry(gui_main, textvariable=bank_file_path, state="readonly").place(x=90, y=25, width=100, height=25)
    Button(gui_main, text="选择账单文件", command=lambda: choose_file("bank")).place(x=200, y=25, width=100, height=25)

    global people_file_path
    people_file_path = StringVar()
    Label(gui_main, text="做账文件：").place(x=400, y=25, width=70, height=25)
    Entry(gui_main, textvariable=people_file_path, state="readonly").place(x=470, y=25, width=100, height=25)
    Button(gui_main, text="选择做账文件", command=lambda: choose_file("people")).place(x=580, y=25, width=100, height=25)
    # ----------选择文件----------

    validate_row_func = gui_main.register(validate_row_input)

    # ---------对账起始行---------
    global bank_start_row
    bank_start_row = StringVar()
    Label(gui_main, text="对账起始行：").place(x=20, y=60, width=80, height=25)
    Entry(gui_main, textvariable=bank_start_row, validate="key", validatecommand=(validate_row_func, "%P")).place(x=100, y=60, width=80, height=25)
    Label(gui_main, text="数字，如：1", fg="orange").place(x=190, y=60, width=80, height=25)

    global people_start_row
    people_start_row = StringVar()
    Label(gui_main, text="对账起始行：").place(x=400, y=60, width=80, height=25)
    Entry(gui_main, textvariable=people_start_row, validate="key", validatecommand=(validate_row_func, "%P")).place(x=480, y=60, width=80, height=25)
    Label(gui_main, text="数字，如：1", fg="orange").place(x=575, y=60, width=80, height=25)
    # ---------对账起始行---------

    validate_col_func = gui_main.register(validate_col_input)

    # ---------交易金额/借方金额所在列---------
    global bank_money_col
    bank_money_col = StringVar()
    Label(gui_main, text="交易金额所在列：").place(x=20, y=95, width=105, height=25)
    Entry(gui_main, textvariable=bank_money_col, validate="key", validatecommand=(validate_col_func, "%P")).place(x=125, y=95, width=62, height=25)
    Label(gui_main, text="大写字母，如：A", fg="orange").place(x=185, y=95, width=130, height=25)

    global people_money_col
    people_money_col = StringVar()
    Label(gui_main, text="借方金额/贷方金额所在列：").place(x=400, y=95, width=165, height=25)
    Entry(gui_main, textvariable=people_money_col, validate="key", validatecommand=(validate_col_func, "%P")).place(x=565, y=95, width=62, height=25)
    Label(gui_main, text="大写字母，如：A", fg="orange").place(x=625, y=95, width=130, height=25)

    global bank_opposite_flag
    bank_opposite_flag = BooleanVar()
    Checkbutton(gui_main, text="取相反数", variable=bank_opposite_flag).place(x=40, y=120, width=80, height=25)

    global people_opposite_flag
    people_opposite_flag = BooleanVar()
    Checkbutton(gui_main, text="取相反数", variable=people_opposite_flag).place(x=420, y=120, width=80, height=25)

    global bank_keep_opposite_money_flag
    bank_keep_opposite_money_flag = BooleanVar()
    bank_check_button = Checkbutton(gui_main, text="保留负数部分", variable=bank_keep_opposite_money_flag)
    bank_check_button.select()
    bank_check_button.place(x=150, y=120, width=120, height=25)

    global people_keep_opposite_money_flag
    people_keep_opposite_money_flag = BooleanVar()
    people_check_button = Checkbutton(gui_main, text="保留负数部分", variable=people_keep_opposite_money_flag)
    people_check_button.select()
    people_check_button.place(x=530, y=120, width=120, height=25)
    # ---------交易金额/借方金额所在列---------

    # ---------日期所在列---------
    global bank_time_col
    bank_time_col = StringVar()
    Label(gui_main, text="日期所在列：").place(x=20, y=155, width=80, height=25)
    Entry(gui_main, textvariable=bank_time_col, validate="key", validatecommand=(validate_col_func, "%P")).place(x=100, y=155, width=80, height=25)
    Label(gui_main, text="大写字母，如：A", fg="orange").place(x=180, y=155, width=130, height=25)

    global people_time_col
    people_time_col = StringVar()
    Label(gui_main, text="日期所在列：").place(x=400, y=155, width=80, height=25)
    Entry(gui_main, textvariable=people_time_col, validate="key", validatecommand=(validate_col_func, "%P")).place(x=480, y=155, width=80, height=25)
    Label(gui_main, text="大写字母，如：A", fg="orange").place(x=560, y=155, width=130, height=25)
    # ---------日期所在列---------

    validate_col_list_func = gui_main.register(validate_col_list_input)

    # ---------需要保留的列---------
    global bank_keep_cols
    bank_keep_cols = StringVar()
    Label(gui_main, text="需要保留的列：").place(x=20, y=190, width=90, height=25)
    Entry(gui_main, textvariable=bank_keep_cols, validate="key", validatecommand=(validate_col_list_func, "%P")).place(x=110, y=190, width=60, height=25)
    Label(gui_main, text="大写字母，逗号分隔，如：A,B,C", fg="orange").place(x=180, y=190, width=200, height=25)

    global people_keep_cols
    people_keep_cols = StringVar()
    Label(gui_main, text="需要保留的列：").place(x=400, y=190, width=90, height=25)
    Entry(gui_main, textvariable=people_keep_cols, validate="key", validatecommand=(validate_col_list_func, "%P")).place(x=490, y=190, width=60, height=25)
    Label(gui_main, text="大写字母，逗号分隔，如：A,B,C", fg="orange").place(x=560, y=190, width=200, height=25)
    # ---------需要保留的列---------

    # ---------sheet名---------
    global sheet_name
    sheet_name = StringVar()
    Label(gui_main, text="sheet重命名：").place(x=20, y=225, width=90, height=25)
    Entry(gui_main, textvariable=sheet_name).place(x=110, y=225, width=150, height=25)
    Label(gui_main, text="此处可重命名，为空将取默认命名", fg="orange").place(x=270, y=225, width=200, height=25)
    # ---------sheet名---------

    # ---------错误信息提醒---------
    global error_msg
    error_msg = StringVar()
    Label(gui_main, textvariable=error_msg, fg="red").place(x=200, y=255, width=400, height=25)
    # ---------错误信息提醒---------

    # ---------确认进行对账，参数校验---------
    global sure_button
    sure_button = Button(gui_main, text="开始对账", command=start_reconciliation)
    sure_button.place(x=350, y=280, width=100, height=25)
    # ---------确认进行对账，参数校验---------

    gui_main.mainloop()


if __name__ == '__main__':
    main()
