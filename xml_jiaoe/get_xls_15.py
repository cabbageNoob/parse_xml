# -*- coding:utf-8 -*-
import xlrd
import pandas as pd
import random
import time

"""
    4.1.15　司法舆情简报推送请求
"""


raw_excel = xlrd.open_workbook('huizong.xlsx',encoding_override="utf-8")

# 根据sheet索引或者名称获取sheet内容
Data_sheet = raw_excel.sheets()[0]  # 通过索引获取
raw_event_list = Data_sheet.col_values(1)[1:]
get_event_list = list(set(raw_event_list))
get_event_list.sort(key=raw_event_list.index)    # 事件名称列表
event_number = len(get_event_list)
# 司法舆情简报序号
BriefReportID_list = []
# 司法舆情简报产生时间
BriefReportWriteTime_list = []
# 司法舆情简报类型
BriefReportType_list= []
# 司法舆情简报名称
FileName_list = []
# 司法舆情简报长度
FileLength_list = []
# 司法舆情简报内容
FileContent_list = []
FileName_list = get_event_list
for i in range(event_number):
    BriefReportID_list.append(i)
    BriefReportWriteTime_list.append(time.strftime('%Y/%m/%d %H/%M/%S', time.localtime(time.time())))
    BriefReportType_list.append(random.randint(1,4))
    FileLength_list.append(10000)
    FileContent_list.append('无')
FileName_list = get_event_list
all_dict_main={'PubInfoBriefReportID':BriefReportID_list,
               'PubInfoBriefReportWriteTime':BriefReportWriteTime_list,
               'PubInfoBriefReportType':BriefReportType_list,
               'PubInfoFileName':FileName_list,
               'PubInfoFileLength':FileLength_list,
               'PubInfoFileContent':FileContent_list}
main_order = ['PubInfoBriefReportID', 'PubInfoBriefReportWriteTime', 'PubInfoBriefReportType', 'PubInfoFileName', 'PubInfoFileLength', 'PubInfoFileContent']
df_main = pd.DataFrame(all_dict_main)
df_main = df_main[main_order]

with pd.ExcelWriter('new15.xls') as Writer:
    df_main.to_excel(Writer,'Sheet1',index=False, encoding='utf_8_sig')