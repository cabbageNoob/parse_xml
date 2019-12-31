# -*- coding: utf-8 -*-
import time
from xml.dom import minidom
import xlrd

"""
    4.1.15　司法舆情简报推送请求
"""


def create_xml():
    sheet_main = xlrd.open_workbook('./xml_jiaoe/xls/new15.xls', encoding_override="utf-8").sheets()[0]
    now_time = time.strftime('%Y/%m/%d %H/%M/%S', time.localtime(time.time()))
    dom = minidom.Document()
    root_node = dom.createElement('Message')
    dom.appendChild(root_node)
    head_node = dom.createElement('MessageHead')
    head_node.setAttribute("ProtocolType", '3')
    head_node.setAttribute("MsgType", '15')
    head_node.setAttribute("MsgID", "0-65535")
    head_node.setAttribute("OperateType", '')
    head_node.setAttribute("SendType", '')
    head_node.setAttribute("Priority", '')
    head_node.setAttribute("DayTime", now_time)
    root_node.appendChild(head_node)
    for i in range(1, sheet_main.nrows):
        content_node = dom.createElement('MessageContent')
        root_node.appendChild(content_node)
        # 司法舆情简报序号
        key_node = dom.createElement('PubInfoBriefReportID')
        content_node.appendChild(key_node)
        key_text = dom.createTextNode(str(int(sheet_main.row_values(i)[0])))
        key_node.appendChild(key_text)

        # 司法舆情简报产生时间
        key_node = dom.createElement('PubInfoBriefReportWriteTime')
        content_node.appendChild(key_node)
        key_text = dom.createTextNode(sheet_main.row_values(i)[1])
        key_node.appendChild(key_text)

        # 司法舆情简报发送时间
        key_node = dom.createElement('PubInfoBriefReportSendTime')
        content_node.appendChild(key_node)
        key_text = dom.createTextNode(time.strftime('%Y/%m/%d %H/%M/%S', time.localtime(time.time())))
        key_node.appendChild(key_text)

        # 司法舆情简报类型
        key_node = dom.createElement('PubInfoBriefReportType')
        content_node.appendChild(key_node)
        key_text = dom.createTextNode(str(int(sheet_main.row_values(i)[2])))
        key_node.appendChild(key_text)

        # 司法舆情简报名称
        key_node = dom.createElement('PubInfoFileName')
        content_node.appendChild(key_node)
        key_text = dom.createTextNode(sheet_main.row_values(i)[3])
        key_node.appendChild(key_text)

        # 司法舆情简报长度
        key_node = dom.createElement('PubInfoFileLength')
        content_node.appendChild(key_node)
        key_text = dom.createTextNode(str(int(sheet_main.row_values(i)[4])))
        key_node.appendChild(key_text)

        # 司法舆情简报内容
        key_node = dom.createElement('PubInfoFileContent')
        content_node.appendChild(key_node)
        key_text = dom.createTextNode(sheet_main.row_values(i)[5])
        key_node.appendChild(key_text)

    xml_str = dom.toxml()
    return xml_str
    # try:
    #     with open('4.1.15司法舆情简报推送请求.xml', 'w', encoding='UTF-8') as fh:
    #         dom.writexml(fh, indent='', addindent='\t', newl='\n', encoding='UTF-8')
    #         print('写司法舆情简报推送请求完毕!')
    # except Exception as err:
    #     print('错误信息：{0}'.format(err))

if __name__ == '__main__':
    create_xml()
