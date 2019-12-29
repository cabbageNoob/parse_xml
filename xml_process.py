# -*- coding: utf-8 -*-
'''
@Descripttion: 
@version: 
@Author: 
@Date: 2019-12-27 11:11:48
@LastEditors  : cjh (492795090@qq.com)
@LastEditTime : 2019-12-29 22:05:42
'''
import os
import xml.etree.ElementTree as ET
from flask import Flask, render_template, redirect, request
from werkzeug.utils import secure_filename
from flask import send_from_directory

GENERATE_XML='D:/uploads/generate_xml'
if not os.path.exists(GENERATE_XML):
    os.mkdir(GENERATE_XML)
app = Flask(__name__)

app.config['GENERATE_XML'] = GENERATE_XML
app.config['SECRET_KEY'] = '123456'

class MessageType(object):
    msg_type = {'1': '实时舆情事件信息通知','2': '实时舆情事件信息接收结果', '3': '司法案件舆情精准搜索服务订阅请求',\
                '4': '司法案件舆情精准搜索服务订阅结果', '10': '司法舆情业务监测信息请求', '11': '司法舆情业务监测信息应答', \
                '12': '司法舆情热点话题监测信息请求', '13': '司法舆情热点话题监测信息应答', '14': '司法舆情话题搜索请求', \
                '15': '司法舆情话题搜索应答', '16': '司法舆情话题传播路径信息请求', '17': '司法舆情话题传播路径信息应答', \
                '18': '司法舆情话题事件关联信息请求', '19': '司法舆情话题事件关联信息应答', '20': '司法舆情简报推送请求', \
                '21': '司法舆情话题搜索应答'}


def parse_xml(xml_path):
    '''
    @Descripttion: 解析xml文件
    @param xml_path :xml文件路径
    @return: 返回需要的xml文件
    '''    
    tree = ET.parse(xml_path)
    root = tree.getroot()

    MsgType = list(root.iter('MessageHead'))[0].attrib['MsgType']
    xml_str = get_attrib_value(root, MsgType)
    root = ET.fromstring(xml_str)
    prettyXml(root, '\t', '\n')
    tree = ET.ElementTree(root)
    file_name = xml_path.split('\\')[-1]
    file_path = os.path.join(app.config['GENERATE_XML'], 'gen_'+file_name)
    tree.write(file_path,encoding='UTF-8',  xml_declaration=True)
    ET.dump(root)
    xml_str=ET.tostring(root, encoding='UTF-8', short_empty_elements=False)
    print(xml_str)
    return xml_str


def get_attrib_value(root,MsgType):
    '''
    @Descripttion: 根据MsgType获取xml信息
    @param MsgType
    @return: 
    '''    
    if (MsgType == '1'):
        pass
    elif (MsgType == '2'):
        PublicSentimentInfoNum_text = list(
            root.iter('PublicSentimentInfoNum'))[0].text
        # PlatCfgRslt_text = list(
        #     root.iter('PlatCfgRslt'))[0].text
        # PlatInfoArea_text = list(
        #     root.iter('PlatInfoArea'))[0].text
        PubInfo = ET.Element('PubInfo')
        PublicSentimentInfoNum = ET.SubElement(PubInfo, 'PublicSentimentInfoNum')
        PublicSentimentInfoNum.text = PublicSentimentInfoNum_text
        PublicSentimentInfoNum = ET.SubElement(PubInfo, 'PublicSentimentInfoNum')
        PublicSentimentInfoNum.text = PublicSentimentInfoNum_text
        # PlatCfgRslt = ET.SubElement(PubInfo, 'PlatCfgRslt')
        # PlatCfgRslt.text = PlatCfgRslt_text
        # PlatInfoArea = ET.SubElement(PubInfo, 'PlatInfoArea')
        # PlatInfoArea.text = PlatInfoArea_text
        xml_str = ET.tostring(PubInfo, encoding='UTF-8', short_empty_elements=False)
        return xml_str
            
    elif (MsgType == '3'):
        PubInfoSerialNum = list(
            root.iter('PubInfoSerialNum'))[0].text
        PubInfoCaseName = list(
            root.iter('PubInfoCaseName'))[0].text
        PubInfoCaseJudgeName = list(
            root.iter('PubInfoCaseJudgeName'))[0].text
    elif (MsgType == '4'):
        pass
    elif (MsgType == '10'):
        pass
    elif (MsgType == '11'):
        pass
    elif (MsgType == '12'):
        pass
    elif (MsgType == '13'):
        pass
    elif (MsgType == '14'):
        pass
    elif (MsgType == '15'):
        pass
    elif (MsgType == '16'):
        pass
    elif (MsgType == '17'):
        pass
    elif (MsgType == '18'):
        pass
    elif (MsgType == '19'):
        pass
    elif (MsgType == '20'):
        pass
    elif (MsgType == '21'):
        PlatCfgRslt = list(
            root.iter('PlatCfgRslt'))[0].text
        PlatInfoArea = list(
            root.iter('PlatInfoArea'))[0].text



# elemnt为传进来的Elment类，参数indent用于缩进，newline用于换行
def prettyXml(element, indent, newline, level=0):
    '''
    @Descripttion: 用于美化xml，避免只有一行
    @param element: element元素
    @param indent: 缩进
    @param newline: 换行
    @param level:  
    @return: 
    '''
    if element:  # 判断element是否有子元素
        if element.text == None or element.text.isspace():  # 如果element的text没有内容
            element.text = newline + indent * (level + 1)
        else:
            element.text = newline + indent * \
                (level + 1) + element.text.strip() + \
                newline + indent * (level + 1)
    #else:  # 此处两行如果把注释去掉，Element的text也会另起一行
        #element.text = newline + indent * (level + 1) + element.text.strip() + newline + indent * level
    temp = list(element)  # 将elemnt转成list
    for subelement in temp:
        # 如果不是list的最后一个元素，说明下一个行是同级别元素的起始，缩进应一致
        if temp.index(subelement) < (len(temp) - 1):
            subelement.tail = newline + indent * (level + 1)
        else:  # 如果是list的最后一个元素， 说明下一行是母元素的结束，缩进应该少一个
            subelement.tail = newline + indent * level
        prettyXml(subelement, indent, newline, level=level + 1)  # 对子元素进行递归操作


def main():
    xml_str = parse_xml('./origin.xml')
    root = ET.fromstring(xml_str)
    prettyXml(root, '\t', '\n')
    tree=ET.ElementTree(root)
    tree.write('a.xml', encoding='UTF-8',  xml_declaration=True)
    return ET.dump(root)
    
if __name__ == '__main__':
    main()
    



