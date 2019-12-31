# -*- coding: utf-8 -*-
import time
from xml.dom import minidom
import xml.etree.ElementTree as ET
import random
import xlrd

"""
    4.1.14 司法舆情话题事件关联信息应答
"""


def create_xml(input_xml):
    sheet_main = xlrd.open_workbook('./xml_jiaoe/xls/new14.xls', encoding_override="utf-8").sheets()[0]
    sheet_TopicCorrelation = xlrd.open_workbook('./xml_jiaoe/xls/new14.xls', encoding_override="utf-8").sheets()[1]
    sheet_TopicEventID = xlrd.open_workbook(
        './xml_jiaoe/xls/new14.xls', encoding_override="utf-8").sheets()[2]

    tree = ET.parse(input_xml)
    root = tree.getroot()
    message = root.find('MessageContent')
    PubInfoTopicID = message.find('PubInfoTopicID').text

    now_time = time.strftime('%Y/%m/%d %H/%M/%S', time.localtime(time.time()))
    dom = minidom.Document()
    root_node = dom.createElement('Message')
    dom.appendChild(root_node)
    head_node = dom.createElement('MessageHead')
    head_node.setAttribute("ProtocolType", '3')
    head_node.setAttribute("MsgType", '14')
    head_node.setAttribute("MsgID", "0-65535")
    head_node.setAttribute("OperateType", '')
    head_node.setAttribute("SendType", '')
    head_node.setAttribute("Priority", '')
    head_node.setAttribute("DayTime", now_time)
    root_node.appendChild(head_node)
    content_node = dom.createElement('MessageContent')
    root_node.appendChild(content_node)

    # PlatCfgRslt = random.randint(1, 2)
    PlatCfgRslt = 1
    key_node = dom.createElement('PlatCfgRslt')
    content_node.appendChild(key_node)
    key_text = dom.createTextNode(str(PlatCfgRslt))
    key_node.appendChild(key_text)

    if PlatCfgRslt == 1:
        PlatInfoArea = ''
        key_node = dom.createElement('PlatInfoArea')
        content_node.appendChild(key_node)
        key_text = dom.createTextNode(PlatInfoArea)
        key_node.appendChild(key_text)

        # 舆情话题标识
        key_node = dom.createElement('PubInfoTopicID')
        content_node.appendChild(key_node)
        key_text = dom.createTextNode(str(PubInfoTopicID))
        key_node.appendChild(key_text)

        # 舆情话题事件关联信息数目
        key_node = dom.createElement('PubInfoTopicCorrelationNum')
        content_node.appendChild(key_node)
        key_text = dom.createTextNode(sheet_main.row_values(1)[1])
        key_node.appendChild(key_text)

        """
            舆情话题事件关联信息区
        """
        for i in range(1, sheet_TopicCorrelation.nrows):
            PubInfoTopicCorrelationArea_node = dom.createElement('PubInfoTopicCorrelationArea')
            content_node.appendChild(PubInfoTopicCorrelationArea_node)

            # 关联话题标识
            key_node = dom.createElement('CorrelationTopicID')
            PubInfoTopicCorrelationArea_node.appendChild(key_node)
            key_text = dom.createTextNode(str(int(sheet_TopicCorrelation.row_values(i)[0])))
            key_node.appendChild(key_text)

            # 关联话题舆情事件数目
            key_node = dom.createElement('CorrelationTopicEventNum')
            PubInfoTopicCorrelationArea_node.appendChild(key_node)
            key_text = dom.createTextNode(str(int(sheet_TopicCorrelation.row_values(i)[1])))
            key_node.appendChild(key_text)

            # 关联话题舆情事件标识信息区
            for j in range(1, sheet_TopicEventID.nrows):
                if sheet_TopicEventID.row_values(j)[0] == sheet_TopicCorrelation.row_values(i)[0]:
                    CorrelationTopicEventIDArea_node = dom.createElement('CorrelationTopicEventIDArea')
                    PubInfoTopicCorrelationArea_node.appendChild(CorrelationTopicEventIDArea_node)

                    # 舆情事件信息标识
                    key_node = dom.createElement('PubInfoID')
                    CorrelationTopicEventIDArea_node.appendChild(key_node)
                    key_text = dom.createTextNode(str(int(sheet_TopicEventID.row_values(j)[1])))
                    key_node.appendChild(key_text)

    xml_str = dom.toxml()
    return xml_str
    # try:
    #     with open('4.1.14司法舆情话题事件关联信息应答.xml', 'w', encoding='UTF-8') as fh:
    #         dom.writexml(fh, indent='', addindent='\t', newl='\n', encoding='UTF-8')
    #         print('写司法舆情话题事件关联信息应答结果完毕!')
    # except Exception as err:
    #     print('错误信息：{0}'.format(err))

if __name__=='__main__':
    create_xml('./xml_jiaoe/xml/get_14.xml')
