'''
@Descripttion: 
@version: 
@Author: cjh (492795090@qq.com)
@Date: 2019-12-30 18:40:45
@LastEditors: cjh (492795090@qq.com)
@LastEditTime: 2019-12-30 18:58:29
'''

# -*- coding:utf-8 -*-
import xlrd
import pandas as pd
import random


raw_excel = xlrd.open_workbook('./xml_jiaoe/huizong.xlsx',encoding_override="utf-8")

# 根据sheet索引或者名称获取sheet内容
Data_sheet = raw_excel.sheets()[0]  # 通过索引获取

raw_event_list = Data_sheet.col_values(1)[1:]
get_event_list = list(set(raw_event_list))
get_event_list.sort(key=raw_event_list.index)    # 事件名称列表
event_number = len(get_event_list)
# 订阅服务流水号
SerialNum_list = []
# 要搜索舆情的案件名称
CaseName_list = get_event_list
# 要搜索舆情的办案法官名称
CaseJudgeName_list = []
# 订阅服务结果
PlatCfgRslt_list = []
# 订阅服务失败原因
PlatInfoArea_list = []
for i in range(event_number):
    SerialNum_list.append(i)
    CaseJudgeName_list.append(i)
    result = random.randint(1, 2)
    PlatCfgRslt_list.append(result)
    if result == 1:
        PlatInfoArea_list.append('')
    else:
        PlatInfoArea_list.append(random.randint(1,3))
all_dict_main={'PubInfoSerialNum':SerialNum_list,
               'PubInfoCaseName':CaseName_list,
               'PubInfoCaseJudgeName':CaseJudgeName_list,
               'PlatCfgRslt':PlatCfgRslt_list,
               'PlatInfoArea':PlatInfoArea_list}
main_order = ['PubInfoSerialNum', 'PubInfoCaseName', 'PubInfoCaseJudgeName', 'PlatCfgRslt', 'PlatInfoArea']
df_main = pd.DataFrame(all_dict_main)
df_main = df_main[main_order]

with pd.ExcelWriter('./xml_jiaoe/xls/new4.xls') as Writer:
    df_main.to_excel(Writer,'Sheet1',index=False, encoding='utf_8_sig')