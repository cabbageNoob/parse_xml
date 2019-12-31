# -*- coding:utf-8 -*-
import xlrd
import pandas as pd
import random

raw_excel = xlrd.open_workbook('huizong.xlsx', encoding_override="utf-8")

# 根据sheet索引或者名称获取sheet内容
Data_sheet = raw_excel.sheets()[0]  # 通过索引获取

raw_event_list = Data_sheet.col_values(1)[1:]
get_event_list = list(set(raw_event_list))
get_event_list.sort(key=raw_event_list.index)  # 事件名称列表
event_number = len(get_event_list)
InfoNum_dict = {}
for event in raw_event_list:
    InfoNum_dict[event] = InfoNum_dict.get(event, 0) + 1
InfoNum_list = list(InfoNum_dict.values())
no_list = list(range(0, event_number))
all_id_list = list(range(0, Data_sheet.nrows - 1))
all_no_list = []
for i in range(1, Data_sheet.nrows):
    this_event = Data_sheet.row_values(i)[1]
    for r in range(len(get_event_list)):
        if this_event == get_event_list[r]:
            all_no_list.append(r)

# 舆情话题标识
TopicID = ['5']
# 舆情话题事件关联信息数目
raw_TopicCorrelationNum = 5
TopicCorrelationNum = [str(raw_TopicCorrelationNum)]

all_dict_main = {'PubInfoTopicID': TopicID,
                 'PubInfoTopicCorrelationNum': TopicCorrelationNum}
main_order = ['PubInfoTopicID', 'PubInfoTopicCorrelationNum']
df_main = pd.DataFrame(all_dict_main)
df_main = df_main[main_order]

# 舆情话题事件关联信息区
# 关联话题标识
CorrelationTopicID_list = random.sample(range(0, event_number), raw_TopicCorrelationNum)
# 关联话题舆情事件数目
CorrelationTopicEventNum_list = []
for id in CorrelationTopicID_list:
    for i in range(len(no_list)):
        if id == no_list[i]:
            CorrelationTopicEventNum_list.append(InfoNum_list[i])
# 关联系数
CorrelationCoefficient_list = []
for m in range(raw_TopicCorrelationNum):
    CorrelationCoefficient_list.append(random.random())
all_dict_TopicCorrelation = {'CorrelationTopicID': CorrelationTopicID_list,
                             'CorrelationTopicEventNum': CorrelationTopicEventNum_list,
                             'CorrelationCoefficient': CorrelationCoefficient_list}
TopicCorrelation_order = ['CorrelationTopicID', 'CorrelationTopicEventNum', 'CorrelationCoefficient']
df_TopicCorrelation = pd.DataFrame(all_dict_TopicCorrelation)
df_TopicCorrelation = df_TopicCorrelation[TopicCorrelation_order]

# 关联话题舆情事件标识信息区
# 关联话题标识
all_CorrelationTopicID_list = []
# 舆情事件信息标识
PubInfoID_list = []
for m in range(len(CorrelationTopicID_list)):
    for n in range(len(all_no_list)):
        if CorrelationTopicID_list[m] == all_no_list[n]:
            all_CorrelationTopicID_list.append(CorrelationTopicID_list[m])
            PubInfoID_list.append(all_id_list[n])
all_dict_CorrelationTopicEventIDArea = {'CorrelationTopicID': all_CorrelationTopicID_list,
                             'PubInfoID': PubInfoID_list}
CorrelationTopicEventIDArea_order = ['CorrelationTopicID', 'PubInfoID']
df_CorrelationTopicEventIDArea = pd.DataFrame(all_dict_CorrelationTopicEventIDArea)
df_CorrelationTopicEventIDArea = df_CorrelationTopicEventIDArea[CorrelationTopicEventIDArea_order]


with pd.ExcelWriter('new14.xls') as Writer:
    df_main.to_excel(Writer, 'Sheet1', index=False, encoding='utf_8_sig')
    df_TopicCorrelation.to_excel(Writer, 'Sheet2', index=False, encoding='utf_8_sig')
    df_CorrelationTopicEventIDArea.to_excel(Writer, 'Sheet3', index=False, encoding='utf_8_sig')
