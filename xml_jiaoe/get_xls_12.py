'''
@Descripttion: 
@version: 
@Author: nlpir
@Date: 2019-12-31 09:16:26
@LastEditors  : nlpir
@LastEditTime : 2019-12-31 12:37:02
'''
# -*- coding:utf-8 -*-

import xlrd
import pandas as pd
import random


# 舆情话题标识
PubInfoTopicID = ['5']
# 舆情话题传播路径数目
PubInfoTopicRouteNum = ['4']

all_dict_main = {'PubInfoTopicID': PubInfoTopicID,
                 'PubInfoTopicRouteNum': PubInfoTopicRouteNum}
main_order = ['PubInfoTopicID', 'PubInfoTopicRouteNum']
df_main = pd.DataFrame(all_dict_main)
df_main = df_main[main_order]

# 舆情话题传播路径信息区
# 舆情事件信息标识
PubInfoID = ['97', '98', '99', '100']
# 舆情事件发生时间
PubInfoStartTime = ['2018-05-24 11:02:00', '2018-05-23 10:18:00', '2018-05-22 15:09:00', '2018-05-22 12:46:00']

all_dict_TopicRoute= {'PubInfoID': PubInfoID,
                 'PubInfoStartTime': PubInfoStartTime}
TopicRoute_order = ['PubInfoID', 'PubInfoStartTime']
df_TopicRoute = pd.DataFrame(all_dict_TopicRoute)
df_TopicRoute = df_TopicRoute[TopicRoute_order]


# 舆情事件传播者信息区
# 传播者名称
CommunicatorName = ['黄冈网警巡查执法', '辽视第一时间', '艾较真', '王松筠']
# 传播者地区
CommunicatorRegion = ['1', '2', '3', '4']
# 传播者地址
CommunicatorAdress = ['暂无', '暂无', '暂无', '暂无']
# 传播者类型
CommunicatorType = ['1', '1', '1', '1']
# 传播者参与时间
CommunicatorJoinTime = ['2018/05/24', '2018/05/23', '2018/05/22', '2018/05/22']
# 传播者好友数目
CommunicatorFriendNum = ['100', '100', '100', '100']
# 传播者粉丝数目
CommunicatorFanNum = ['100', '100', '100', '100']
all_dict_Communicator= {'CommunicatorName': CommunicatorName,
                 'CommunicatorRegion': CommunicatorRegion,
                        'CommunicatorAdress': CommunicatorAdress,
                        'CommunicatorType': CommunicatorType,
                        'CommunicatorJoinTime': CommunicatorJoinTime,
                        'CommunicatorFriendNum': CommunicatorFriendNum,
                        'CommunicatorFanNum': CommunicatorFanNum}
Communicator_order = ['CommunicatorName', 'CommunicatorRegion', 'CommunicatorAdress', 'CommunicatorType',
                      'CommunicatorJoinTime', 'CommunicatorFriendNum', 'CommunicatorFanNum']
df_Communicator = pd.DataFrame(all_dict_Communicator)
df_Communicator = df_Communicator[Communicator_order]

# 舆情事件被传播者信息区
# 传播者名称
BeCommunicatorName = ['aaa']
# 传播者地区
BeCommunicatorRegion = ['1']
# 传播者地址
BeCommunicatorAdress = ['暂无']
# 传播者类型
BeCommunicatorType = ['1']
# 传播者参与时间
CommunicatorJoinTime = ['2018/05/20']
# 传播者好友数目
CommunicatorFriendNum = ['100']
# 传播者粉丝数目
CommunicatorFanNum = ['100']
all_dict_BeCommunicator= {'BeCommunicatorName': BeCommunicatorName,
                 'BeCommunicatorRegion': BeCommunicatorRegion,
                        'BeCommunicatorAdress': BeCommunicatorAdress,
                        'BeCommunicatorType': BeCommunicatorType,
                        'CommunicatorJoinTime': CommunicatorJoinTime,
                        'CommunicatorFriendNum': CommunicatorFriendNum,
                        'CommunicatorFanNum': CommunicatorFanNum}
BeCommunicator_order = ['BeCommunicatorName', 'BeCommunicatorRegion', 'BeCommunicatorAdress', 'BeCommunicatorType',
                      'CommunicatorJoinTime', 'CommunicatorFriendNum', 'CommunicatorFanNum']
df_BeCommunicator = pd.DataFrame(all_dict_BeCommunicator)
df_BeCommunicator = df_BeCommunicator[BeCommunicator_order]

# 舆情事件传播详情信息区（PubInfoCommunicateDetail）
# 传播主题
CommunicateTopic = ['a', 'b', 'c', 'd']
# 传播内容
CommunicateContent = ['aa', 'bb', 'cc', 'dd']
# 传播观点
CommunicateViewpoint = ['a1', 'b2', 'c3', 'd4']
# 情感
CommunicateEmotion = ['1', '2', '3', '1']
# 转载数
CommunicateBeforwardNum = ['5', '20', '10', '3']
# 信息类型
CommunicateContentType = ['1', '1', '1', '1']
# 业务类型
CommunicateWorkType = ['5', '5', '5', '3']
# 涉及部门
CommunicateDepartment = ['3', '3', '1', '2']

all_dict_CommunicateDetail= {'CommunicateTopic': CommunicateTopic,
                 'CommunicateContent': CommunicateContent,
                        'CommunicateViewpoint': CommunicateViewpoint,
                        'CommunicateEmotion': CommunicateEmotion,
                        'CommunicateBeforwardNum': CommunicateBeforwardNum,
                        'CommunicateContentType': CommunicateContentType,
                             'CommunicateWorkType':CommunicateWorkType,
                             'CommunicateDepartment':CommunicateDepartment
                        }
CommunicateDetail_order = ['CommunicateTopic', 'CommunicateContent', 'CommunicateViewpoint', 'CommunicateEmotion',
                      'CommunicateBeforwardNum', 'CommunicateContentType', 'CommunicateWorkType', 'CommunicateDepartment']
df_CommunicateDetail = pd.DataFrame(all_dict_CommunicateDetail)
df_CommunicateDetail = df_CommunicateDetail[CommunicateDetail_order]


with pd.ExcelWriter('new12.xls') as Writer:
    df_main.to_excel(Writer, 'Sheet1', index=False, encoding='utf_8_sig')
    df_TopicRoute.to_excel(Writer, 'Sheet2', index=False, encoding='utf_8_sig')
    df_Communicator.to_excel(Writer, 'Sheet3', index=False, encoding='utf_8_sig')
    df_BeCommunicator.to_excel(Writer, 'Sheet4', index=False, encoding='utf_8_sig')
    df_CommunicateDetail.to_excel(Writer, 'Sheet5', index=False, encoding='utf_8_sig')
