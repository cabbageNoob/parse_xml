'''
@Descripttion: 
@version: 
@Author: nlpir
@Date: 2019-12-30 13:34:26
@LastEditors  : nlpir
@LastEditTime : 2019-12-31 12:35:48
'''
# -*- coding: utf-8 -*-
import time,os
from xml.dom import minidom
import xml.etree.ElementTree as ET

import xlrd

pwd_path = os.path.abspath(os.path.dirname(__file__))
# xls file
XLS_1_FILE = os.path.join(pwd_path, './xls/new1.xls')

"""
    4.1.1　实时舆情事件信息通知
"""


def create_xml():
    sheet_main = xlrd.open_workbook(
        XLS_1_FILE, encoding_override="utf-8").sheets()[0]
    sheet_PublicSentimentInfo = xlrd.open_workbook(
        XLS_1_FILE, encoding_override="utf-8").sheets()[1]
    sheet_PubInfoFile = xlrd.open_workbook(
        XLS_1_FILE, encoding_override="utf-8").sheets()[2]
    sheet_PubInfoDepartment = xlrd.open_workbook(
        XLS_1_FILE, encoding_override="utf-8").sheets()[3]
    sheet_PubInfoDepartmentWork = xlrd.open_workbook(
        XLS_1_FILE, encoding_override="utf-8").sheets()[4]
    number = sheet_main.nrows - 1
    now_time = time.strftime('%Y/%m/%d %H/%M/%S', time.localtime(time.time()))
    dom = minidom.Document()
    root_node = dom.createElement('Message')
    dom.appendChild(root_node)
    head_node = dom.createElement('MessageHead')
    head_node.setAttribute("ProtocolType", '3')
    head_node.setAttribute("MsgType", '1')
    head_node.setAttribute("MsgID", "0-65535")
    head_node.setAttribute("OperateType", '')
    head_node.setAttribute("SendType", '')
    head_node.setAttribute("Priority", '')
    head_node.setAttribute("DayTime", now_time)
    root_node.appendChild(head_node)

    for n in range(1, sheet_main.nrows):
        print(n)
        content_node = dom.createElement('MessageContent')
        root_node.appendChild(content_node)
        # 舆情上报类型
        PubInfoSendType = int(sheet_main.row_values(n)[0])
        key_node = dom.createElement('PubInfoSendType')
        content_node.appendChild(key_node)
        key_text = dom.createTextNode(str(int(sheet_main.row_values(n)[0])))
        key_node.appendChild(key_text)
        # 订阅服务流水号
        if PubInfoSendType == 2:
            PubInfoSerialNum = str(int(sheet_main.row_values(n)[1]))
        else:
            PubInfoSerialNum = ''
        key_node = dom.createElement('PubInfoSerialNum')
        content_node.appendChild(key_node)
        key_text = dom.createTextNode(PubInfoSerialNum)
        key_node.appendChild(key_text)
        # 要搜索舆情的案件名称
        if PubInfoSendType == 2:
            PubInfoCaseName = sheet_main.row_values(n)[2]
        else:
            PubInfoCaseName = ''
        key_node = dom.createElement('PubInfoCaseName')
        content_node.appendChild(key_node)
        key_text = dom.createTextNode(PubInfoCaseName)
        key_node.appendChild(key_text)
        # 要搜索舆情的办案法官名称
        if PubInfoSendType == 2:
            PubInfoCaseJudgeName = str(int(sheet_main.row_values(n)[3]))
        else:
            PubInfoCaseJudgeName = ''
        key_node = dom.createElement('PubInfoCaseJudgeName')
        content_node.appendChild(key_node)
        key_text = dom.createTextNode(PubInfoCaseJudgeName)
        key_node.appendChild(key_text)
        # 舆情事件信息数量
        key_node = dom.createElement('PublicSentimentInfoNum')
        content_node.appendChild(key_node)
        key_text = dom.createTextNode(str(int(sheet_main.row_values(n)[4])))
        key_node.appendChild(key_text)
        # PublicSentimentInfoArea舆情事件信息区
        if sheet_main.row_values(n)[4] > 0:
            this_event = sheet_main.row_values(n)[2]
            for m in range(1, sheet_PublicSentimentInfo.nrows):
                if this_event == sheet_PublicSentimentInfo.row_values(m)[0]:
                    PublicSentimentInfo_node = dom.createElement('PublicSentimentInfoArea')
                    content_node.appendChild(PublicSentimentInfo_node)
                    # 舆情事件信息标识
                    key_node = dom.createElement('PubInfoID')
                    PublicSentimentInfo_node.appendChild(key_node)
                    key_text = dom.createTextNode(str(int(sheet_PublicSentimentInfo.row_values(m)[1])))
                    key_node.appendChild(key_text)
                    # 舆情事件信息标题
                    key_node = dom.createElement('PubInfoTitile')
                    PublicSentimentInfo_node.appendChild(key_node)
                    key_text = dom.createTextNode(str(int(sheet_PublicSentimentInfo.row_values(m)[2])))
                    key_node.appendChild(key_text)
                    # 舆情事件信息内容
                    key_node = dom.createElement('PubInfoContent')
                    PublicSentimentInfo_node.appendChild(key_node)
                    key_text = dom.createTextNode(sheet_PublicSentimentInfo.row_values(m)[3])
                    key_node.appendChild(key_text)
                    # 舆情事件信息产生时间
                    key_node = dom.createElement('PubInfoTime')
                    PublicSentimentInfo_node.appendChild(key_node)
                    key_text = dom.createTextNode(sheet_PublicSentimentInfo.row_values(m)[4])
                    key_node.appendChild(key_text)
                    # 舆情事件信息来源网站类型
                    key_node = dom.createElement('PubInfoSrcType')
                    PublicSentimentInfo_node.appendChild(key_node)
                    key_text = dom.createTextNode(str(int(sheet_PublicSentimentInfo.row_values(m)[5])))
                    key_node.appendChild(key_text)
                    # 舆情事件信息来源网站名称
                    key_node = dom.createElement('PubInfoSrcWebsiteName')
                    PublicSentimentInfo_node.appendChild(key_node)
                    key_text = dom.createTextNode(sheet_PublicSentimentInfo.row_values(m)[6])
                    key_node.appendChild(key_text)
                    # 舆情事件信息发起网民
                    key_node = dom.createElement('PubInfoSrcNetizenName')
                    PublicSentimentInfo_node.appendChild(key_node)
                    key_text = dom.createTextNode(sheet_PublicSentimentInfo.row_values(m)[7])
                    key_node.appendChild(key_text)
                    # 舆情信事件息转发数量
                    key_node = dom.createElement('PubInfoForwardNum')
                    PublicSentimentInfo_node.appendChild(key_node)
                    key_text = dom.createTextNode(str(int(sheet_PublicSentimentInfo.row_values(m)[8])))
                    key_node.appendChild(key_text)
                    # 舆情事件信息网站链接
                    key_node = dom.createElement('PubInfoSrcWebsiteLink')
                    PublicSentimentInfo_node.appendChild(key_node)
                    key_text = dom.createTextNode(sheet_PublicSentimentInfo.row_values(m)[9])
                    key_node.appendChild(key_text)
                    # 舆情附件数量
                    key_node = dom.createElement('PubInfoFileNum')
                    PublicSentimentInfo_node.appendChild(key_node)
                    key_text = dom.createTextNode(str(int(sheet_PublicSentimentInfo.row_values(m)[10])))
                    key_node.appendChild(key_text)
                    # 舆情附件信息区
                    if int(sheet_PublicSentimentInfo.row_values(m)[10]) > 0:
                        for p in range(1, sheet_PubInfoFile.nrows):
                            # print('*************', int(sheet_PublicSentimentInfo.row_values(m)[1]))
                            # print('++++++++++++', int(sheet_PubInfoFile.row_values(p)[0]))
                            if int(sheet_PublicSentimentInfo.row_values(m)[1]) == int(sheet_PubInfoFile.row_values(p)[0]):
                                PubInfoFileArea_node = dom.createElement('PubInfoFileArea')
                                PublicSentimentInfo_node.appendChild(PubInfoFileArea_node)
                                # 舆情附件序号
                                key_node = dom.createElement('PubInfoFileID')
                                PubInfoFileArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(str(int(sheet_PubInfoFile.row_values(p)[1])))
                                key_node.appendChild(key_text)
                                # 舆情附件名称
                                key_node = dom.createElement('PubInfoFileName')
                                PubInfoFileArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(sheet_PubInfoFile.row_values(p)[2])
                                key_node.appendChild(key_text)
                                # 舆情附件类型
                                key_node = dom.createElement('PubInfoFileType')
                                PubInfoFileArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(sheet_PubInfoFile.row_values(p)[3])
                                key_node.appendChild(key_text)
                                # 舆情附件长度
                                key_node = dom.createElement('PubInfoFileLength')
                                PubInfoFileArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(sheet_PubInfoFile.row_values(p)[4])
                                key_node.appendChild(key_text)
                                # 舆情附件内容
                                key_node = dom.createElement('PubInfoFileContent')
                                PubInfoFileArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(sheet_PubInfoFile.row_values(p)[5])
                                key_node.appendChild(key_text)

                    # 舆情协同类型
                    key_node = dom.createElement('PubInfoCorType')
                    PublicSentimentInfo_node.appendChild(key_node)
                    key_text = dom.createTextNode(str(int(sheet_PublicSentimentInfo.row_values(m)[11])))
                    key_node.appendChild(key_text)
                    # 舆情涉及部门数量
                    key_node = dom.createElement('PubInfoDepartmentNum')
                    PublicSentimentInfo_node.appendChild(key_node)
                    key_text = dom.createTextNode(str(int(sheet_PublicSentimentInfo.row_values(m)[12])))
                    key_node.appendChild(key_text)
                    # 舆情涉及部门信息区
                    if int(sheet_PublicSentimentInfo.row_values(m)[12]) > 0:
                        for q in range(1, sheet_PubInfoDepartment.nrows):
                            if int(sheet_PublicSentimentInfo.row_values(m)[12]) == int(sheet_PubInfoDepartment.row_values(q)[0]):
                                PubInfoDepartmentArea_node = dom.createElement('PubInfoDepartmentArea')
                                PublicSentimentInfo_node.appendChild(PubInfoDepartmentArea_node)
                                # 舆情涉及部门类型
                                key_node = dom.createElement('PubInfoDepartmentType')
                                PubInfoDepartmentArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(str(int(sheet_PubInfoDepartment.row_values(q)[1])))
                                key_node.appendChild(key_text)
                                # 舆情涉及部门级别
                                key_node = dom.createElement('PubInfoDepartmentType')
                                PubInfoDepartmentArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(str(int(sheet_PubInfoDepartment.row_values(q)[2])))
                                key_node.appendChild(key_text)
                                # 舆情涉及部门所在省份/直辖市
                                key_node = dom.createElement('PubInfoDepProvince')
                                PubInfoDepartmentArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(str(int(sheet_PubInfoDepartment.row_values(q)[3])))
                                key_node.appendChild(key_text)
                                # 舆情涉及部门所在市
                                key_node = dom.createElement('PubInfoDepCity')
                                PubInfoDepartmentArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(sheet_PubInfoDepartment.row_values(q)[4])
                                key_node.appendChild(key_text)
                                # 舆情涉及部门区域
                                key_node = dom.createElement('PubInfoDepartmentArea')
                                PubInfoDepartmentArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(sheet_PubInfoDepartment.row_values(q)[5])
                                key_node.appendChild(key_text)
                                # 舆情涉及部门名称
                                key_node = dom.createElement('PubInfoDepartmentName')
                                PubInfoDepartmentArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(sheet_PubInfoDepartment.row_values(q)[6])
                                key_node.appendChild(key_text)
                    # 舆情涉及部门业务数量
                    key_node = dom.createElement('PubInfoDepartmentWorkNum')
                    PublicSentimentInfo_node.appendChild(key_node)
                    key_text = dom.createTextNode(str(int(sheet_PublicSentimentInfo.row_values(m)[13])))
                    key_node.appendChild(key_text)
                    # 舆情涉及部门业务信息区
                    if int(sheet_PublicSentimentInfo.row_values(m)[13]) > 0:
                        for k in range(1, sheet_PubInfoDepartmentWork.nrows):
                            if str(int(sheet_PublicSentimentInfo.row_values(m)[1])) == str(int(sheet_PubInfoDepartmentWork.row_values(k)[0])):
                                PubInfoDepartmentWorkArea_node = dom.createElement('PubInfoDepartmentWorkArea')
                                PublicSentimentInfo_node.appendChild(PubInfoDepartmentWorkArea_node)
                                # 舆情涉及部门业务类型
                                key_node = dom.createElement('PubInfoDepartmentWorkType')
                                PubInfoDepartmentWorkArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(str(int(sheet_PubInfoDepartmentWork.row_values(k)[1])))
                                key_node.appendChild(key_text)
                                # 舆情涉及部门业务内容
                                key_node = dom.createElement('PubInfoDepartmentWorkContent')
                                PubInfoDepartmentWorkArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(sheet_PubInfoDepartmentWork.row_values(k)[2])
                                key_node.appendChild(key_text)

    xml_str=dom.toxml()
    return xml_str
    # try:
    #     with open('./xml_jiaoe/xml/4.1.1实时舆情事件信息通知.xml', 'w', encoding='UTF-8') as fh:
    #         dom.writexml(fh, indent='', addindent='\t', newl='\n', encoding='UTF-8')
    #         print('写实时舆情事件信息通知文件完毕!')
    # except Exception as err:
    #     print('错误信息：{0}'.format(err))

if __name__=='__main__':
    create_xml()


