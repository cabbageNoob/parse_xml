'''
@Descripttion: 
@version: 
@Author: cjh (492795090@qq.com)
@Date: 2019-12-30 13:43:09
@LastEditors  : cjh (492795090@qq.com)
@LastEditTime : 2019-12-30 19:21:11
'''
import os
import xml.etree.ElementTree as ET
from xml_jiaoe import parse_xml_4_str
def parse_xml_4(xml_path):
    xml_str = parse_xml_4_str.create_xml(xml_path)
    return xml_str
def parse_xml_8(xml_path):
    tree = ET.parse(xml_path)
    root=tree.getroot()
    Message = ET.Element('Message')
    MessageContent=ET.SubElement(Message,'MessageContent')
    if tree:
        PlatCfgRslt_text = '1'
        # 舆情热点话题信息接收结果
        PlatCfgRslt = ET.SubElement(MessageContent, 'PlatCfgRslt')
        PlatCfgRslt.text=PlatCfgRslt_text
    else:
        PlatCfgRslt_text = '2'
        PlatCfgRslt = ET.SubElement(MessageContent, 'PlatCfgRslt')
        PlatCfgRslt.text = PlatCfgRslt_text
        # 舆情热点信息接收失败原因
        PlatInfoArea = ET.SubElement(MessageContent, 'PlatInfoArea')
        xml_str = ET.tostring(Message, encoding='UTF-8',
                              short_empty_elements=False)
        return xml_str
    
    PubInfoStartTime_text = list(root.iter('PubInfoStartTime'))[0].text
    PubInfoEndTime_text = list(root.iter('PubInfoEndTime'))[0].text
    # 舆情监测起始时间
    PubInfoStartTime = ET.SubElement(MessageContent, 'PubInfoStartTime')
    PubInfoStartTime.text = PubInfoStartTime_text
    # 舆情监测结束时间
    PubInfoEndTime = ET.SubElement(MessageContent, 'PubInfoEndTime')
    PubInfoEndTime.text = PubInfoEndTime_text
    # 获取的热点话题数量
    PubInfoTopicNum = ET.SubElement(MessageContent, 'PubInfoTopicNum')
    PubInfoTopicNum.text = '1'
    # 热点话题信息区
    PubInfoTopicArea = ET.SubElement(MessageContent, 'PubInfoTopicArea')

    ## 热点话题信息区
    # 舆情热点话题标识
    PubInfoTopicID = ET.SubElement(PubInfoTopicArea, 'PubInfoTopicID')
    PubInfoTopicID.text = ''
    # 舆情热点话题标题
    PubInfoTopicTitle = ET.SubElement(PubInfoTopicArea, 'PubInfoTopicTitle')
    PubInfoTopicTitle.text = ''
    # 舆情热点话题内容
    PubInfoTopicContent = ET.SubElement(PubInfoTopicArea, 'PubInfoTopicContent')
    PubInfoTopicContent.text = ''
    # 舆情热点话题摘要
    PubInfoTopicAbstract = ET.SubElement(
        PubInfoTopicArea, 'PubInfoTopicAbstract')
    PubInfoTopicAbstract.text = ''
    # 舆情热点话题产生时间
    PubInfoTopicStartTime = ET.SubElement(
        MessageContent, 'PubInfoTopicStartTime')
    PubInfoTopicStartTime.text = ''
    # 舆情热点话题结束时间
    PubInfoTopicEndTime = ET.SubElement(MessageContent, 'PubInfoTopicEndTime')
    PubInfoTopicEndTime.text=''
    # ......
    xml_str = ET.tostring(Message, encoding='UTF-8',short_empty_elements=False)
    return xml_str

def parse_xml_10(xml_path):
    tree = ET.parse(xml_path)
    root=tree.getroot()
    Message = ET.Element('Message')
    MessageContent = ET.SubElement(Message, 'MessageContent')
    
    if tree:
        PlatCfgRslt_text = '1'
        # 舆情热点话题信息接收结果
        PlatCfgRslt = ET.SubElement(MessageContent, 'PlatCfgRslt')
        PlatCfgRslt.text = PlatCfgRslt_text
    else:
        PlatCfgRslt_text = '2'
        PlatCfgRslt = ET.SubElement(MessageContent, 'PlatCfgRslt')
        PlatCfgRslt.text = PlatCfgRslt_text
        # 舆情热点信息接收失败原因
        PlatInfoArea = ET.SubElement(MessageContent, 'PlatInfoArea')
        xml_str = ET.tostring(Message, encoding='UTF-8',short_empty_elements=False)
        return xml_str
    
    # 获取的话题数量
    PubInfoTopicNum = ET.SubElement(MessageContent, 'PubInfoTopicNum')
    PubInfoTopicNum.text = '1'
    # 话题信息区
    PubInfoTopicArea = ET.SubElement(MessageContent, 'PubInfoTopicArea')

    ## 话题信息区
    PubInfoTopicID = ET.SubElement(PubInfoTopicArea, 'PubInfoTopicID')
    PubInfoTopicID.text = ''
    # 舆情热点话题标题
    PubInfoTopicTitle = ET.SubElement(PubInfoTopicArea, 'PubInfoTopicTitle')
    PubInfoTopicTitle.text = ''
    # 舆情热点话题内容
    PubInfoTopicContent = ET.SubElement(
        PubInfoTopicArea, 'PubInfoTopicContent')
    PubInfoTopicContent.text = ''
    # 舆情热点话题摘要
    PubInfoTopicAbstract = ET.SubElement(
        PubInfoTopicArea, 'PubInfoTopicAbstract')
    PubInfoTopicAbstract.text = ''
    # 舆情热点话题产生时间
    PubInfoTopicStartTime = ET.SubElement(
        MessageContent, 'PubInfoTopicStartTime')
    PubInfoTopicStartTime.text = ''
    # 舆情热点话题结束时间
    PubInfoTopicEndTime = ET.SubElement(MessageContent, 'PubInfoTopicEndTime')
    PubInfoTopicEndTime.text = ''
    # ......
    xml_str = ET.tostring(Message, encoding='UTF-8',
                          short_empty_elements=False)
    return xml_str
    
        
        


    
    
