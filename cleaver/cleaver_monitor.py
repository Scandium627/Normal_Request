# -*- coding: utf-8 -*-
"""
@Time ： 2022/5/24 17:30
@Auth ： Scandiuum
@File ：cleaver_monitor.py
@IDE ：PyCharm
"""
import datetime
import os
import re
import PyQt5
import xlrd

excel_path = "cleaver_sister.xls"

class Model_excel():
    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.date_now = datetime.date.today()
        self.delay_title_list = ["上门时间", "提评估", "接收评估报告", "送达评估报告", "制作拍卖审批表", "提拍卖", "开拍时间"]
        self.delay_list = [7, 15, 30, 5, 5, 5, 5]
        self.excel = xlrd.open_workbook(filename=self.excel_path)
        self.sheet = self.excel[0]


    def judge_delay_time(self,input_row,input_now_date):
        date_range = input_row[9:17]
        print(date_range)
        for i in range(7):
            print(date_range[i].ctype)
            if date_range[i].ctype == 0:
                print(input_row[8+i].value)
                judge_cell = input_row[8+i]
                if judge_cell.ctype == 3:
                    date_tuple = xlrd.xldate_as_tuple(judge_cell.value, self.excel.datemode)
                elif judge_cell.ctype == 2:
                    date_tuple = self.transform_xldate(judge_cell.value)
                else:
                    continue
                if self.compare_date(date_tuple,input_now_date) >= self.delay_list[i]:
                    out_row = [cell.value for cell in input_row]
                    out_row.append(self.delay_title_list[i])
                    return out_row
        return False


    def transform_xldate(self,input_cell_value):
        if type(input_cell_value) == str:
            match =re.search(r"(\d{4}.\d{1,2}.\d{1,2})",input_cell_value)
            date_str = input_cell_value[match.regs[0][0]:match.regs[0][1]]
            date_list = date_str.split('/')
            return date_list
    #
    def compare_date(self,input_date_tripe,now_date):
        if input_date_tripe and now_date:
            spot_date = datetime.date(input_date_tripe[0], input_date_tripe[1],input_date_tripe[2])
            now_datetime = datetime.date(now_date.year, now_date.month,now_date.day)
            days_delay = now_datetime.__sub__(spot_date).days
            return days_delay

    def scan_rows(self):
        alert_out_list = []
        rows = self.sheet.get_rows()
        for i in rows:
            scan_row_result = self.judge_delay_time(i,self.date_now)
            if scan_row_result:
                alert_out_list.append(scan_row_result)
        return alert_out_list


class Pyqt():
    def __init__(self, excel_path):
        self.excel_path = excel_path



a1 = Model_excel(excel_path)
a2 = a1.scan_rows()


## 此行之下结尾测试
