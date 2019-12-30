'''
@Descripttion: 
@version: 
@Author: cjh (492795090@qq.com)
@Date: 2019-12-30 14:14:29
@LastEditors  : cjh (492795090@qq.com)
@LastEditTime : 2019-12-30 14:18:08
'''
import xlrd
import pandas as pd
import random

def parse_xlsx(xlsx_path):
    raw_excel = xlrd.open_workbook(xlsx_path, encoding_override="utf-8")
    # 根据sheet索引或者名称获取sheet内容
    Data_sheet = raw_excel.sheets()[0]  # 通过索引获取
    print(raw_excel)

def get_data_from_date(begin_date, end_date):
    pass

if __name__ == '__main__':
    parse_xlsx('./static/huizong.xlsx')
