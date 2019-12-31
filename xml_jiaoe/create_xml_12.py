# -*- coding: utf-8 -*-
import time,os
from xml.dom import minidom
import xml.etree.ElementTree as ET
import random
import xlrd

pwd_path = os.path.abspath(os.path.dirname(__file__))
# xls file
XLS_12_FILE = os.path.join(pwd_path, './xls/new12.xls')

"""
    4.1.12 司法舆情话题传播路径信息应答
"""


def create_xml(input_xml):
    sheet_main = xlrd.open_workbook(
        XLS_12_FILE, encoding_override="utf-8").sheets()[0]
    df_TopicRoute = xlrd.open_workbook(
        XLS_12_FILE, encoding_override="utf-8").sheets()[1]
    df_Communicator = xlrd.open_workbook(
        XLS_12_FILE, encoding_override="utf-8").sheets()[2]
    df_BeCommunicator = xlrd.open_workbook(
        XLS_12_FILE, encoding_override="utf-8").sheets()[3]
    df_CommunicateDetail = xlrd.open_workbook(
        XLS_12_FILE, encoding_override="utf-8").sheets()[4]

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

        # 舆情话题传播路径数目
        key_node = dom.createElement('PubInfoTopicRouteNum')
        content_node.appendChild(key_node)
        key_text = dom.createTextNode(sheet_main.row_values(1)[1])
        key_node.appendChild(key_text)

        """
            舆情话题传播路径信息区
        """
        for i in range(1, df_TopicRoute.nrows):
            PubInfoTopicRouteArea_node = dom.createElement('PubInfoTopicRouteArea')
            content_node.appendChild(PubInfoTopicRouteArea_node)

            # 舆情事件信息标识
            key_node = dom.createElement('PubInfoID')
            PubInfoTopicRouteArea_node.appendChild(key_node)
            key_text = dom.createTextNode(str(int(df_TopicRoute.row_values(i)[0])))
            key_node.appendChild(key_text)

            # 舆情事件发生时间
            key_node = dom.createElement('PubInfoStartTime')
            PubInfoTopicRouteArea_node.appendChild(key_node)
            key_text = dom.createTextNode(df_TopicRoute.row_values(i)[1])
            key_node.appendChild(key_text)

            # 舆情事件传播者信息区
            for j in range(1, df_Communicator.nrows):
                PubInfoCommunicatorArea_node = dom.createElement('PubInfoCommunicatorArea')
                PubInfoTopicRouteArea_node.appendChild(PubInfoCommunicatorArea_node)

                # 传播者名称
                key_node = dom.createElement('CommunicatorName')
                PubInfoCommunicatorArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_Communicator.row_values(j)[0])
                key_node.appendChild(key_text)

                # 传播者地区
                key_node = dom.createElement('CommunicatorRegion')
                PubInfoCommunicatorArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_Communicator.row_values(j)[1])
                key_node.appendChild(key_text)

                # 传播者地址
                key_node = dom.createElement('CommunicatorAdress')
                PubInfoCommunicatorArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_Communicator.row_values(j)[2])
                key_node.appendChild(key_text)

                # 传播者类型
                key_node = dom.createElement('CommunicatorType')
                PubInfoCommunicatorArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_Communicator.row_values(j)[3])
                key_node.appendChild(key_text)

                # 传播者参与时间
                key_node = dom.createElement('CommunicatorJoinTime')
                PubInfoCommunicatorArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_Communicator.row_values(j)[4])
                key_node.appendChild(key_text)

                # 传播者好友数目
                key_node = dom.createElement('CommunicatorFriendNum')
                PubInfoCommunicatorArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_Communicator.row_values(j)[5])
                key_node.appendChild(key_text)

                # 传播者粉丝数目
                key_node = dom.createElement('CommunicatorFanNum')
                PubInfoCommunicatorArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_Communicator.row_values(j)[6])
                key_node.appendChild(key_text)

            # 舆情事件被传播者信息区
            for k in range(1, df_BeCommunicator.nrows):
                PubInfoBeCommunicatorArea_node = dom.createElement('PubInfoBeCommunicatorArea')
                PubInfoTopicRouteArea_node.appendChild(PubInfoBeCommunicatorArea_node)

                # 被传播者名称
                key_node = dom.createElement('BeCommunicatorName')
                PubInfoBeCommunicatorArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_BeCommunicator.row_values(k)[0])
                key_node.appendChild(key_text)

                # 被传播者地区
                key_node = dom.createElement('BeCommunicatorRegion')
                PubInfoBeCommunicatorArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_BeCommunicator.row_values(k)[1])
                key_node.appendChild(key_text)

                # 被传播者地址
                key_node = dom.createElement('BeCommunicatorAdress')
                PubInfoBeCommunicatorArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_BeCommunicator.row_values(k)[2])
                key_node.appendChild(key_text)

                # 被传播者类型
                key_node = dom.createElement('BeCommunicatorType')
                PubInfoBeCommunicatorArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_BeCommunicator.row_values(k)[3])
                key_node.appendChild(key_text)

                # 传播者参与时间
                key_node = dom.createElement('CommunicatorJoinTime')
                PubInfoBeCommunicatorArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_BeCommunicator.row_values(k)[4])
                key_node.appendChild(key_text)

                # 传播者好友数目
                key_node = dom.createElement('CommunicatorFriendNum')
                PubInfoBeCommunicatorArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_BeCommunicator.row_values(k)[5])
                key_node.appendChild(key_text)

                # 传播者粉丝数目
                key_node = dom.createElement('CommunicatorFanNum')
                PubInfoBeCommunicatorArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_BeCommunicator.row_values(k)[6])
                key_node.appendChild(key_text)
            # 舆情事件传播详情信息区（PubInfoCommunicateDetail）
            for m in range(1, df_CommunicateDetail.nrows):
                PubInfoCommunicateDetailArea_node = dom.createElement('PubInfoCommunicateDetailArea')
                PubInfoTopicRouteArea_node.appendChild(PubInfoCommunicateDetailArea_node)

                # 传播主题
                key_node = dom.createElement('CommunicateTopic')
                PubInfoCommunicateDetailArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_CommunicateDetail.row_values(m)[0])
                key_node.appendChild(key_text)

                # 传播内容
                key_node = dom.createElement('CommunicateContent')
                PubInfoCommunicateDetailArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_CommunicateDetail.row_values(m)[1])
                key_node.appendChild(key_text)

                # 传播观点
                key_node = dom.createElement('CommunicateViewpoint')
                PubInfoCommunicateDetailArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_CommunicateDetail.row_values(m)[2])
                key_node.appendChild(key_text)

                # 情感
                key_node = dom.createElement('CommunicateEmotion')
                PubInfoCommunicateDetailArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_CommunicateDetail.row_values(m)[3])
                key_node.appendChild(key_text)

                # 转载数
                key_node = dom.createElement('CommunicateBeforwardNum')
                PubInfoCommunicateDetailArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_CommunicateDetail.row_values(m)[4])
                key_node.appendChild(key_text)

                # 信息类型
                key_node = dom.createElement('CommunicateContentType')
                PubInfoCommunicateDetailArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_CommunicateDetail.row_values(m)[5])
                key_node.appendChild(key_text)

                # 业务类型
                key_node = dom.createElement('CommunicateWorkType')
                PubInfoCommunicateDetailArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_CommunicateDetail.row_values(m)[6])
                key_node.appendChild(key_text)

                # 涉及部门
                key_node = dom.createElement('CommunicateDepartment')
                PubInfoCommunicateDetailArea_node.appendChild(key_node)
                key_text = dom.createTextNode(df_CommunicateDetail.row_values(m)[7])
                key_node.appendChild(key_text)

    xml_str=dom.toxml()
    return xml_str

    # try:
    #     with open('4.1.12司法舆情话题传播路径信息应答.xml', 'w', encoding='UTF-8') as fh:
    #         dom.writexml(fh, indent='', addindent='\t', newl='\n', encoding='UTF-8')
    #         print('写司法舆情话题传播路径信息应答结果完毕!')
    # except Exception as err:
    #     print('错误信息：{0}'.format(err))

if __name__=='__main__':
    create_xml('./xml_jiaoe/xml/get_12.xml')
