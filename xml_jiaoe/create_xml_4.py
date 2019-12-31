'''
@Descripttion: 
@version: 
@Author: cjh (492795090@qq.com)
@Date: 2019-12-30 18:40:49
@LastEditors  : cjh (492795090@qq.com)
@LastEditTime : 2019-12-31 10:49:59
'''

# -*- coding: utf-8 -*-
import time,os
from xml.dom import minidom
import xml.etree.ElementTree as ET
import random
import xlrd

pwd_path = os.path.abspath(os.path.dirname(__file__))
print(pwd_path)
# xls file
XLS_4_FILE = os.path.join(pwd_path, './xls/new4.xls')
"""
    4.1.4　司法案件舆情精准搜索服务订阅结果
"""


def create_xml(input_xml):
    sheet_main = xlrd.open_workbook(
        XLS_4_FILE, encoding_override="utf-8").sheets()[0]
    now_time = time.strftime('%Y/%m/%d %H/%M/%S', time.localtime(time.time()))
    dom = minidom.Document()
    root_node = dom.createElement('Message')
    dom.appendChild(root_node)
    head_node = dom.createElement('MessageHead')
    head_node.setAttribute("ProtocolType", '3')
    head_node.setAttribute("MsgType", '4')
    head_node.setAttribute("MsgID", "0-65535")
    head_node.setAttribute("OperateType", '')
    head_node.setAttribute("SendType", '')
    head_node.setAttribute("Priority", '')
    head_node.setAttribute("DayTime", now_time)
    root_node.appendChild(head_node)
    content_node = dom.createElement('MessageContent')
    root_node.appendChild(content_node)
    tree = ET.parse(input_xml)
    root = tree.getroot()
    message = root.find('MessageContent')
    PubInfoSerialNum = message.find('PubInfoSerialNum')
    # 订阅服务流水号
    if PubInfoSerialNum.text is None:
        PubInfoSerialNum = ''
    else:
        PubInfoSerialNum = PubInfoSerialNum.text
    key_node = dom.createElement('PubInfoSerialNum')
    content_node.appendChild(key_node)
    key_text = dom.createTextNode(PubInfoSerialNum)
    key_node.appendChild(key_text)

    # 要搜索舆情的案件名称
    PubInfoCaseName = message.find('PubInfoCaseName')
    if PubInfoCaseName.text is None:
        PubInfoCaseName = ''
    else:
        PubInfoCaseName = PubInfoCaseName.text
    key_node = dom.createElement('PubInfoCaseName')
    content_node.appendChild(key_node)
    key_text = dom.createTextNode(PubInfoCaseName)
    key_node.appendChild(key_text)

    # 要搜索舆情的办案法官名称
    PubInfoCaseJudgeName = message.find('PubInfoCaseJudgeName')
    if PubInfoCaseJudgeName.text is None:
        PubInfoCaseJudgeName = ''
    else:
        PubInfoCaseJudgeName = PubInfoCaseJudgeName.text
    key_node = dom.createElement('PubInfoCaseJudgeName')
    content_node.appendChild(key_node)
    key_text = dom.createTextNode(PubInfoCaseName)
    key_node.appendChild(key_text)

    # 订阅服务结果
    PlatCfgRslt = random.randint(1, 2)
    key_node = dom.createElement('PlatCfgRslt')
    content_node.appendChild(key_node)
    key_text = dom.createTextNode(str(PlatCfgRslt))
    key_node.appendChild(key_text)

    # 订阅服务失败原因
    if PlatCfgRslt == 1:
        PlatInfoArea = ''
    else:
        PlatInfoArea = random.randint(1, 3)
    key_node = dom.createElement('PlatInfoArea')
    content_node.appendChild(key_node)
    key_text = dom.createTextNode(str(PlatInfoArea))
    key_node.appendChild(key_text)

    xml_str = dom.toxml()
    return xml_str
    
    # try:
    #     with open('./xml_jiaoe/xml/4.1.4司法案件舆情精准搜索服务订阅结果.xml', 'w', encoding='UTF-8') as fh:
    #         dom.writexml(fh, indent='', addindent='\t', newl='\n', encoding='UTF-8')
    #         print('写司法案件舆情精准搜索服务订阅结果完毕!')
    # except Exception as err:
    #     print('错误信息：{0}'.format(err))

if __name__=='__main__':
    create_xml('./xml_jiaoe/get_4.xml')
