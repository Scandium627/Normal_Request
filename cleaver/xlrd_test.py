# -*- coding: utf-8 -*-
"""
@Time ： 2022/5/27 13:28
@Auth ： Scandiuum
@File ：xlrd_test.py
@IDE ：PyCharm
"""
#
# excel = xlrd.open_workbook(filename=excel_path)
#
# excel.sheet_names()
#
# sheet1 = excel.sheets()[0]
# rows = sheet1.get_rows()
# l1 = []
# num = 0
# for i in rows:
#     if num <= 3:
#         l1.append(i)
#         num += 1
#
# l3 = l1[2]
# 44637 -44630
# l3[9].ctype
# print(l3[9])
# if "xldate" in l3[9]:
#     print("done")
#
# cell_value = sheet1
#
# date_tuple = xlrd.xldate_as_tuple(42429, excel.datemode)
#
# delay_dic = {
#     "k":7,
#     "l":15,
#     "m":30,
#     "n":5,
#     "o":5,
#     "p":5,
#     "q":5
# }
# delay_list = [7, 15, 30, 5, 5, 5, 5]
# date_now = datetime.date.today()
#
# print(date_now.day)
#
# def judge_delay_time(input_row_list,input_now_date):
#     date_range = input_row_list[9:17]
#     print(date_range)
#     for i in range(7):
#         print(date_range[i].ctype)
#         if date_range[i].ctype == 0:
#             print(input_row_list[8+i].value)
#             judge_cell = input_row_list[8+i]
#             if judge_cell.ctype == 3:
#                 date_tuple = xlrd.xldate_as_tuple(judge_cell.value, excel.datemode)
#             elif judge_cell.ctype == 2:
#                 date_tuple = transform_xldate(judge_cell.value)
#             print(compare_date(date_tuple,input_now_date) , delay_list[i])
#             if compare_date(date_tuple,input_now_date) >= delay_list[i]:
#                 return False
#     return True
#
#
# t1 = judge_delay_time(l3, date_now)
#
#
#
#
# def transform_xldate(input_cell_value):
#     if type(input_cell_value) == str:
#         match =re.search(r"(\d{4}.\d{1,2}.\d{1,2})",input_cell_value)
#         date_str = input_cell_value[match.regs[0][0]:match.regs[0][1]]
#         date_list = date_str.split('/')
#         return date_list
# #
# def compare_date(input_date_tripe,now_date):
#     if input_date_tripe and now_date:
#         spot_date = datetime.date(input_date_tripe[0], input_date_tripe[1],input_date_tripe[2])
#         now_datetime = datetime.date(now_date.year, now_date.month,now_date.day)
#         days_delay = now_datetime.__sub__(spot_date).days
#         return days_delay
#
# t1 = judge_delay_time(l3,date_now)
#
# def judge_from_dic(input_row):
#     for i in input_row:
#
#
# print(l3[9:17])
# date_tuple = xlrd.xldate_as_tuple(42429, excel.datemode)
#
# print(date_tuple)
# y = datetime.date(date_tuple[0],date_tuple[1],28)
# x = datetime.date(date_tuple[0],date_tuple[1],date_tuple[1])
# aa = y.__sub__(x)
# aa.days
# a1 = detetime
# 2016,2,29
# import re
# a1= re.compile("(3[0-1]|[12][0-9]|0?[0-9])/(1[0-2]|0?[1-9])")
# '2022/5/28一拍'.
# mat = re.search(r"(\d{4}.\d{1,2}.\d{1,2})",'2022/5/12一拍')
# mat.regs
# [mat.regs[0][0]:mat.regs[0][1]]
#
# transform_xldate('2022/5/28一拍')