# -*- coding:utf-8 -*-
import sys,os
import xlrd
import pandas as pd
import random
sys.path.insert(0, os.getcwd())

raw_excel = xlrd.open_workbook('./xml_jiaoe/huizong.xlsx',encoding_override="utf-8")

# 根据sheet索引或者名称获取sheet内容
Data_sheet = raw_excel.sheets()[0]  # 通过索引获取

raw_event_list = Data_sheet.col_values(1)[1:]
get_event_list = list(set(raw_event_list))
get_event_list.sort(key=raw_event_list.index)    # 事件名称列表
event_number = len(get_event_list)
InfoNum_dict = {}
for event in raw_event_list:
    InfoNum_dict[event] = InfoNum_dict.get(event, 0) + 1
print(InfoNum_dict)
InfoNum_list = list(InfoNum_dict.values())
# print(get_event_list)
SendType_list = []
for i in range(event_number):
    SendType_list.append(random.randint(1, 2))
print(SendType_list)
SerialNum_list = []
for i in range(event_number):
    SerialNum_list.append(i)
JudgeName_list = []
for i in range(event_number):
    JudgeName_list.append(i)

all_dict_main={'PubInfoSendType':SendType_list,
               'PubInfoSerialNum':SerialNum_list,
               'PubInfoCaseName':get_event_list,
               'PubInfoCaseJudgeName':JudgeName_list,
               'PublicSentimentInfoNum':InfoNum_list}
main_order = ['PubInfoSendType', 'PubInfoSerialNum', 'PubInfoCaseName', 'PubInfoCaseJudgeName', 'PublicSentimentInfoNum']
df_main = pd.DataFrame(all_dict_main)
df_main = df_main[main_order]


# 舆情事件信息区:要搜索舆情的案件名称, 舆情事件信息标识, 舆情事件信息标题, 舆情事件信息内容, 舆情事件信息产生时间, 舆情事件信息来源网站类型, 舆情事件信息来源网站名称, 舆情事件信息发起网民, 舆情信事件息转发数量, 舆情事件信息网站链接, 舆情附件数量
InfoArea_event_list = raw_event_list
PubInfoID_list = []
for i in range(len(InfoArea_event_list)):
    PubInfoID_list.append(i)
InfoTitile_list = []
for i in range(len(InfoArea_event_list)):
    InfoTitile_list.append(i)
InfoContent_list = Data_sheet.col_values(7)[1:]
InfoTime_list = Data_sheet.col_values(6)[1:]
InfoSrcType_list = []
for i in range(len(InfoArea_event_list)):
    InfoSrcType_list.append('1')
InfoSrcWebsiteName_list = []
for i in range(len(InfoArea_event_list)):
    InfoSrcWebsiteName_list.append('新浪微博')
SrcNetizenName_list = Data_sheet.col_values(3)[1:]
InfoForwardNum_list = Data_sheet.col_values(8)[1:]
SrcWebsiteLink_list = []
for i in range(len(InfoArea_event_list)):
    SrcWebsiteLink_list.append('无')
InfoFileNum_list = []
for i in Data_sheet.col_values(11)[1:]:
    InfoFile_list = i.replace('[', '').replace(']', '').replace("'", '').replace("'", '').split(',')
    if InfoFile_list[0] != '':
        InfoFileNum_list.append(len(InfoFile_list))
    else:
        InfoFileNum_list.append(0)
InfoCorType_list = []
for i in range(len(InfoArea_event_list)):
    InfoCorType_list.append(random.randint(1, 2))
InfoDepartmentNum_list = []
for i in range(len(InfoArea_event_list)):
    InfoDepartmentNum_list.append(random.randint(0, 2))
InfoDepartmentWorkNum_list = []
for i in range(len(InfoArea_event_list)):
    InfoDepartmentWorkNum_list.append(random.randint(0, 2))

all_dict_PublicSentimentInfo={'PublicSentimentInfo_PubInfoCaseName':InfoArea_event_list,
                              'PubInfoID':PubInfoID_list,
                              'PubInfoTitile':InfoTitile_list,
                              'PubInfoContent':InfoContent_list,
                              'PubInfoTime':InfoTime_list,
                              'PubInfoSrcType':InfoSrcType_list,
                              'PubInfoSrcWebsiteName':InfoSrcWebsiteName_list,
                              'PubInfoSrcNetizenName':SrcNetizenName_list,
                              'PubInfoForwardNum':InfoForwardNum_list,
                              'PubInfoSrcWebsiteLink':SrcWebsiteLink_list,
                              'PubInfoFileNum':InfoFileNum_list,
                              'PubInfoCorType':InfoCorType_list,
                              'PubInfoDepartmentNum':InfoDepartmentNum_list,
                              'PubInfoDepartmentWorkNum':InfoDepartmentWorkNum_list
                            }
PublicSentimentInfo_order = ['PublicSentimentInfo_PubInfoCaseName', 'PubInfoID', 'PubInfoTitile', 'PubInfoContent',
                             'PubInfoTime', 'PubInfoSrcType', 'PubInfoSrcWebsiteName', 'PubInfoSrcNetizenName',
                             'PubInfoForwardNum','PubInfoSrcWebsiteLink', 'PubInfoFileNum', 'PubInfoCorType',
                             'PubInfoDepartmentNum', 'PubInfoDepartmentWorkNum']
df_PublicSentimentInfo = pd.DataFrame(all_dict_PublicSentimentInfo)
df_PublicSentimentInfo = df_PublicSentimentInfo[PublicSentimentInfo_order]


# 舆情附件信息区
PubInfoFile_PubInfoID_list = []    # 舆情事件信息标识
PubInfoFileID_list = []    # 舆情附件序号
PubInfoFileName_list = []    # 舆情附件名称
PubInfoFileType_list = []    # 舆情附件类型
PubInfoFileLength_list = []    # 舆情附件长度
PubInfoFileContent_list = []    # 舆情附件内容
id = -1
for i in range(len(Data_sheet.col_values(11)[1:])):
    InfoFile_list = Data_sheet.col_values(11)[1:][i].replace('[', '').replace(']', '').replace("'", '').replace("'", '').split(',')
    if InfoFile_list[0] != '':
        for j in InfoFile_list:
            id = id + 1
            PubInfoFile_PubInfoID_list.append(i)
            PubInfoFileID_list.append(id)
            PubInfoFileName_list.append(j)
            PubInfoFileType_list.append('图像')
            PubInfoFileLength_list.append('无')
            PubInfoFileContent_list.append('无')

all_dict_PubInfoFile={'PubInfoFile_PubInfoID':PubInfoFile_PubInfoID_list,
                              'PubInfoFileID':PubInfoFileID_list,
                              'PubInfoFileName':PubInfoFileName_list,
                              'PubInfoFileType':PubInfoFileType_list,
                              'PubInfoFileLength':PubInfoFileLength_list,
                              'PubInfoFileContent':PubInfoFileContent_list
                            }
PubInfoFile_order = ['PubInfoFile_PubInfoID', 'PubInfoFileID', 'PubInfoFileName', 'PubInfoFileType', 'PubInfoFileLength',
                     'PubInfoFileContent']
df_PubInfoFile = pd.DataFrame(all_dict_PubInfoFile)
df_PubInfoFile = df_PubInfoFile[PubInfoFile_order]


# 舆情涉及部门信息区
PubInfoDepartment_PubInfoID_list = []    # 舆情事件信息标识
PubInfoDepartmentType_list = []   # 舆情涉及部门类型
PubInfoDepartmentLevel_list = []    # 舆情涉及部门级别
PubInfoDepProvince_list = []    # 舆情涉及部门所在省份/直辖市
PubInfoDepCity_list = []    # 舆情涉及部门所在市
PubInfoDepartmentArea_list = []    # 舆情涉及部门区域
PubInfoDepartmentName_list = []    # 舆情涉及部门名称
for k in range(len(InfoDepartmentNum_list)):
    if InfoDepartmentNum_list[k] != 0:
        for l in range(InfoDepartmentNum_list[k]):
            PubInfoDepartment_PubInfoID_list.append(k)
            PubInfoDepartmentType_list.append(random.randint(1, 3))
            PubInfoDepartmentLevel_list.append(random.randint(1, 6))
            PubInfoDepProvince_list.append(random.randint(1, 6))
            PubInfoDepCity_list.append('地理识别调用')
            PubInfoDepartmentArea_list.append('地理识别调用')
            PubInfoDepartmentName_list.append('舆情涉及部门名称暂无')

all_dict_PubInfoDepartment={'PubInfoDepartment_PubInfoID':PubInfoDepartment_PubInfoID_list,
                              'PubInfoDepartmentType':PubInfoDepartmentType_list,
                              'PubInfoDepartmentLevel':PubInfoDepartmentLevel_list,
                              'PubInfoDepProvince':PubInfoDepProvince_list,
                              'PubInfoDepCity':PubInfoDepCity_list,
                              'PubInfoDepartmentArea':PubInfoDepartmentArea_list,
                              'PubInfoDepartmentName':PubInfoDepartmentName_list
                            }
PubInfoDepartment_order = ['PubInfoDepartment_PubInfoID', 'PubInfoDepartmentType', 'PubInfoDepartmentLevel', 'PubInfoDepProvince',
                           'PubInfoDepCity', 'PubInfoDepartmentArea', 'PubInfoDepartmentName']
df_PubInfoDepartment = pd.DataFrame(all_dict_PubInfoDepartment)
df_PubInfoDepartment = df_PubInfoDepartment[PubInfoDepartment_order]


# 4.1.1.1.3　舆情涉及部门业务信息区
PubInfoDepartmentWork_PubInfoID_list = []    # 舆情事件信息标识
PubInfoDepartmentWorkType_list = []    # 舆情涉及部门业务类型
PubInfoDepartmentWorkContent_list = []    # 舆情涉及部门业务内容
for k in range(len(InfoDepartmentWorkNum_list)):
    if InfoDepartmentWorkNum_list[k] != 0:
        for l in range(InfoDepartmentWorkNum_list[k]):
            PubInfoDepartmentWork_PubInfoID_list.append(k)
            PubInfoDepartmentWorkType_list.append(random.randint(1, 39))
            PubInfoDepartmentWorkContent_list.append('舆情涉及部门业务内容暂无')
all_dict_PubInfoDepartmentWork={'PubInfoDepartmentWork_PubInfoID':PubInfoDepartmentWork_PubInfoID_list,
                              'PubInfoDepartmentWorkType':PubInfoDepartmentWorkType_list,
                              'PubInfoDepartmentWorkContent':PubInfoDepartmentWorkContent_list,
                            }
PubInfoDepartmentWork_order = ['PubInfoDepartmentWork_PubInfoID', 'PubInfoDepartmentWorkType', 'PubInfoDepartmentWorkContent']
df_PubInfoDepartmentWork = pd.DataFrame(all_dict_PubInfoDepartmentWork)
df_PubInfoDepartmentWork = df_PubInfoDepartmentWork[PubInfoDepartmentWork_order]


with pd.ExcelWriter('./xml_jiaoe/xls/new1.xls') as Writer:
    df_main.to_excel(Writer,'Sheet1',index=False, encoding='utf_8_sig')
    df_PublicSentimentInfo.to_excel(Writer, 'Sheet2', index=False, encoding='utf_8_sig')
    df_PubInfoFile.to_excel(Writer, 'Sheet3', index=False, encoding='utf_8_sig')
    df_PubInfoDepartment.to_excel(Writer, 'Sheet4', index=False, encoding='utf_8_sig')
    df_PubInfoDepartmentWork.to_excel(Writer, 'Sheet5', index=False, encoding='utf_8_sig')





