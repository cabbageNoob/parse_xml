'''
@Descripttion: 
@version: 
@Author: nlpir
@Date: 2019-12-31 12:05:09
@LastEditors  : nlpir
@LastEditTime : 2019-12-31 12:36:10
'''
# -*- coding: utf-8 -*-

import time,os
from xml.dom import minidom
import xml.etree.ElementTree as ET
import random
import xlrd

pwd_path = os.path.abspath(os.path.dirname(__file__))
# xls file
XLS_6_FILE = os.path.join(pwd_path, './xls/new6.xls')

"""
    4.1.6 司法舆情业务监测信息应答
"""


def create_xml(input_xml):
    sheet_main = xlrd.open_workbook(
        XLS_6_FILE, encoding_override="utf-8").sheets()[0]

    now_time = time.strftime('%Y/%m/%d %H/%M/%S', time.localtime(time.time()))
    dom = minidom.Document()
    root_node = dom.createElement('Message')
    dom.appendChild(root_node)
    head_node = dom.createElement('MessageHead')
    head_node.setAttribute("ProtocolType", '3')
    head_node.setAttribute("MsgType", '11')
    head_node.setAttribute("MsgID", "0-65535")
    head_node.setAttribute("OperateType", '')
    head_node.setAttribute("SendType", '')
    head_node.setAttribute("Priority", '')
    head_node.setAttribute("DayTime", now_time)
    root_node.appendChild(head_node)
    content_node = dom.createElement('MessageContent')
    root_node.appendChild(content_node)
    # 部门舆情检索信息接收结果
    key_node = dom.createElement('PlatCfgRslt')
    content_node.appendChild(key_node)
    key_text = dom.createTextNode('1')
    key_node.appendChild(key_text)
    # 部门舆情检索信息接收结果
    key_node = dom.createElement('PlatInfoArea')
    content_node.appendChild(key_node)
    key_text = dom.createTextNode('')
    key_node.appendChild(key_text)

    tree = ET.parse(input_xml)
    root = tree.getroot()
    message = root.find('MessageContent')
    # 司法舆情业务监测信息类型
    PubInfoStaticType = message.find('PubInfoStaticType').text
    key_node = dom.createElement('PubInfoStaticType')
    content_node.appendChild(key_node)
    key_text = dom.createTextNode(str(PubInfoStaticType))
    key_node.appendChild(key_text)

    # 部门类型
    PubInfoDepartmentType = message.find('PubInfoDepartmentType').text
    key_node = dom.createElement('PubInfoDepartmentType')
    content_node.appendChild(key_node)
    key_text = dom.createTextNode(str(PubInfoDepartmentType))
    key_node.appendChild(key_text)

    # 部门级别
    PubInfoDepartmentLevel = message.find('PubInfoDepartmentLevel').text
    key_node = dom.createElement('PubInfoDepartmentLevel')
    content_node.appendChild(key_node)
    key_text = dom.createTextNode(str(PubInfoDepartmentLevel))
    key_node.appendChild(key_text)

    # 部门所在省份/直辖市
    PubInfoDepProvince = message.find('PubInfoDepProvince').text
    key_node = dom.createElement('PubInfoDepProvince')
    content_node.appendChild(key_node)
    key_text = dom.createTextNode(str(PubInfoDepProvince))
    key_node.appendChild(key_text)

    # 部门所在市
    PubInfoDepCity = message.find('PubInfoDepCity').text
    key_node = dom.createElement('PubInfoDepCity')
    content_node.appendChild(key_node)
    key_text = dom.createTextNode(str(PubInfoDepCity))
    key_node.appendChild(key_text)

    # 如果部门级别为1，则上报全国舆情综合监测信息区
    if int(PubInfoDepartmentLevel) == 1:
        # 全国舆情综合监测信息区PubInfoCountryStaticArea
        PubInfoCountryStaticArea_node = dom.createElement('PubInfoCountryStaticArea')
        content_node.appendChild(PubInfoCountryStaticArea_node)

        # 全国舆情综合监测信息区PubInfoCountryStaticArea----全国舆情事件信息发生总数量PubInfoCountryAmount
        PubInfoCountryAmount = sheet_main.nrows - 1
        key_node = dom.createElement('PubInfoCountryAmount')
        PubInfoCountryStaticArea_node.appendChild(key_node)
        key_text = dom.createTextNode(str(PubInfoCountryAmount))
        key_node.appendChild(key_text)

        # 全国舆情综合监测信息区PubInfoCountryStaticArea----全国舆情事件信息标识区PubInfoCountryIDArea
        PubInfoCountryIDArea_node = dom.createElement('PubInfoCountryIDArea')
        PubInfoCountryStaticArea_node.appendChild(PubInfoCountryIDArea_node)
        # 全国舆情事件信息标识区PubInfoCountryIDArea----舆情事件信息标识PubInfoID
        for i in range(1, sheet_main.nrows):
            key_node = dom.createElement('PubInfoID')
            PubInfoCountryIDArea_node.appendChild(key_node)
            key_text = dom.createTextNode(str(int(sheet_main.row_values(i)[0])))
            key_node.appendChild(key_text)

        """
            全国舆情综合监测信息区PubInfoCountryStaticArea----按情感统计的全国信息区PubInfoCountryEmotionArea
        """
        PubInfoCountryEmotionArea_node = dom.createElement('PubInfoCountryEmotionArea')
        PubInfoCountryStaticArea_node.appendChild(PubInfoCountryEmotionArea_node)

        # 按情感统计的全国信息区PubInfoCountryEmotionArea----正面舆情数量
        PubInfoPositiveAmount_node = dom.createElement('PubInfoPositiveAmount')
        PubInfoCountryEmotionArea_node.appendChild(PubInfoPositiveAmount_node)
        # 正面舆情信息标识区
        PubInfoPositiveIDArea_node = dom.createElement('PubInfoPositiveIDArea')
        PubInfoCountryEmotionArea_node.appendChild(PubInfoPositiveIDArea_node)
        # 负面舆情数量
        PubInfoNegativeAmount_node = dom.createElement('PubInfoNegativeAmount')
        PubInfoCountryEmotionArea_node.appendChild(PubInfoNegativeAmount_node)
        # 负面舆情信息标识区
        PubInfoNegativeIDArea_node = dom.createElement('PubInfoNegativeIDArea')
        PubInfoCountryEmotionArea_node.appendChild(PubInfoNegativeIDArea_node)
        # 中性舆情数量
        PubInfoNeutralAmount_node = dom.createElement('PubInfoNeutralAmount')
        PubInfoCountryEmotionArea_node.appendChild(PubInfoNeutralAmount_node)
        # 中性舆情信息标识区
        PubInfoNeutralIDArea_node = dom.createElement('PubInfoNeutralIDArea')
        PubInfoCountryEmotionArea_node.appendChild(PubInfoNeutralIDArea_node)

        PubInfoPositiveAmount = 0
        PubInfoNegativeAmount = 0
        PubInfoNeutralAmount = 0
        for i in range(1, sheet_main.nrows):
            if int(sheet_main.row_values(i)[1]) == 1:
                PubInfoPositiveAmount = PubInfoPositiveAmount + 1
                key_node = dom.createElement('PubInfoID')
                PubInfoPositiveIDArea_node.appendChild(key_node)
                key_text = dom.createTextNode(str(int(sheet_main.row_values(i)[0])))
                key_node.appendChild(key_text)
            if int(sheet_main.row_values(i)[1]) == -1:
                PubInfoNegativeAmount = PubInfoNegativeAmount + 1
                key_node = dom.createElement('PubInfoID')
                PubInfoNegativeIDArea_node.appendChild(key_node)
                key_text = dom.createTextNode(str(int(sheet_main.row_values(i)[0])))
                key_node.appendChild(key_text)
            if int(sheet_main.row_values(i)[1]) == 0:
                PubInfoNeutralAmount = PubInfoNeutralAmount + 1
                key_node = dom.createElement('PubInfoID')
                PubInfoNeutralIDArea_node.appendChild(key_node)
                key_text = dom.createTextNode(str(int(sheet_main.row_values(i)[0])))
                key_node.appendChild(key_text)

        PubInfoPositiveAmount_text = dom.createTextNode(str(PubInfoPositiveAmount))
        PubInfoPositiveAmount_node.appendChild(PubInfoPositiveAmount_text)
        PubInfoNegativeAmount_text = dom.createTextNode(str(PubInfoNegativeAmount))
        PubInfoNegativeAmount_node.appendChild(PubInfoNegativeAmount_text)
        PubInfoNeutralAmount_text = dom.createTextNode(str(PubInfoNeutralAmount))
        PubInfoNeutralAmount_node.appendChild(PubInfoNeutralAmount_text)


        # 全国舆情综合监测信息区PubInfoCountryStaticArea----部门业务数目PubInfoWorkNum
        raw_PubInfoWork = sheet_main.col_values(2)[1:]
        PubInfoWork = list(set(raw_PubInfoWork))
        PubInfoWork.sort(key=raw_PubInfoWork.index)
        WorkAmount_dict = {}
        for work in raw_PubInfoWork:
            WorkAmount_dict[work] = WorkAmount_dict.get(work, 0) + 1
        InfoNum_list = list(WorkAmount_dict.values())

        PubInfoWorkNum = len(PubInfoWork)
        key_node = dom.createElement('PubInfoWorkNum')
        PubInfoCountryStaticArea_node.appendChild(key_node)
        key_text = dom.createTextNode(str(PubInfoWorkNum))
        key_node.appendChild(key_text)

        """
           全国舆情综合监测信息区PubInfoCountryStaticArea----按业务统计的全国舆情数目信息区 PubInfoWorkAmountArea
        """
        for j in range(PubInfoWorkNum):
            PubInfoWorkAmountArea_node = dom.createElement('PubInfoWorkAmountArea')
            PubInfoCountryStaticArea_node.appendChild(PubInfoWorkAmountArea_node)

            # 按业务统计的全国舆情数目信息区 PubInfoWorkAmountArea----舆情部门业务类型PubInfoDepartmentWorkType
            PubInfoDepartmentWorkType = PubInfoWork[j]
            key_node = dom.createElement('PubInfoDepartmentWorkType')
            PubInfoWorkAmountArea_node.appendChild(key_node)
            key_text = dom.createTextNode(str(int(PubInfoDepartmentWorkType)))
            key_node.appendChild(key_text)

            # 按业务统计的全国舆情数目信息区 PubInfoWorkAmountArea----该类业务全国舆情发生总数量CountryPubInfoWorkAmount
            CountryPubInfoWorkAmount = InfoNum_list[j]
            key_node = dom.createElement('CountryPubInfoWorkAmount')
            PubInfoWorkAmountArea_node.appendChild(key_node)
            key_text = dom.createTextNode(str(CountryPubInfoWorkAmount))
            key_node.appendChild(key_text)

            """
                  按业务统计的全国舆情数目信息区 PubInfoWorkAmountArea----该类业务全国舆情信息标识区  CountryPubInfoWorkIDArea
            """
            CountryPubInfoWorkIDArea_node = dom.createElement('CountryPubInfoWorkIDArea')
            PubInfoWorkAmountArea_node.appendChild(CountryPubInfoWorkIDArea_node)
            for k in range(1, sheet_main.nrows):
                if sheet_main.row_values(k)[2] == PubInfoWork[j]:
                    # 该类业务全国舆情信息标识区  CountryPubInfoWorkIDArea-----该类业务舆情事件信息标识PubInfoID
                    key_node = dom.createElement('PubInfoID')
                    CountryPubInfoWorkIDArea_node.appendChild(key_node)
                    key_text = dom.createTextNode(str(int(sheet_main.row_values(k)[0])))
                    key_node.appendChild(key_text)

            """
                   按业务统计的全国舆情数目信息区 PubInfoWorkAmountArea---- 该类业务按情感统计的全国舆情信息区CountryPubInfoWorkEmotionArea
            """
            CountryPubInfoWorkEmotionArea_node = dom.createElement('CountryPubInfoWorkEmotionArea')
            PubInfoWorkAmountArea_node.appendChild(CountryPubInfoWorkEmotionArea_node)

            # 正面舆情数量
            PubInfoPositiveAmount_node = dom.createElement('PubInfoPositiveAmount')
            CountryPubInfoWorkEmotionArea_node.appendChild(key_node)
            # 正面舆情信息标识区
            PubInfoPositiveIDArea_node = dom.createElement('PubInfoPositiveIDArea')
            CountryPubInfoWorkEmotionArea_node.appendChild(PubInfoPositiveIDArea_node)
            # 负面舆情数量
            PubInfoNegativeAmount_node = dom.createElement('PubInfoNegativeAmount')
            CountryPubInfoWorkEmotionArea_node.appendChild(key_node)
            # 负面舆情信息标识区
            PubInfoNegativeIDArea_node = dom.createElement('PubInfoNegativeIDArea')
            CountryPubInfoWorkEmotionArea_node.appendChild(PubInfoNegativeIDArea_node)
            # 中性舆情数量
            PubInfoNeutralAmount_node = dom.createElement('PubInfoNeutralAmount')
            CountryPubInfoWorkEmotionArea_node.appendChild(key_node)
            # 中性舆情信息标识区
            PubInfoNeutralIDArea_node = dom.createElement('PubInfoNeutralIDArea')
            CountryPubInfoWorkEmotionArea_node.appendChild(PubInfoNeutralIDArea_node)

            PubInfoPositiveAmount = 0
            PubInfoNegativeAmount = 0
            PubInfoNeutralAmount = 0
            for l in range(1, sheet_main.nrows):
                if sheet_main.row_values(l)[2] == PubInfoWork[j]:
                    if int(sheet_main.row_values(l)[1]) == 1:
                        PubInfoPositiveAmount = PubInfoPositiveAmount + 1
                        key_node = dom.createElement('PubInfoID')
                        PubInfoPositiveIDArea_node.appendChild(key_node)
                        key_text = dom.createTextNode(str(int(sheet_main.row_values(l)[0])))
                        key_node.appendChild(key_text)
                    if int(sheet_main.row_values(l)[1]) == -1:
                        PubInfoNegativeAmount = PubInfoNegativeAmount + 1
                        key_node = dom.createElement('PubInfoID')
                        PubInfoNegativeIDArea_node.appendChild(key_node)
                        key_text = dom.createTextNode(str(int(sheet_main.row_values(l)[0])))
                        key_node.appendChild(key_text)
                    if int(sheet_main.row_values(l)[1]) == 0:
                        PubInfoNeutralAmount = PubInfoNeutralAmount + 1
                        key_node = dom.createElement('PubInfoID')
                        PubInfoNeutralIDArea_node.appendChild(key_node)
                        key_text = dom.createTextNode(str(int(sheet_main.row_values(l)[0])))
                        key_node.appendChild(key_text)

            PubInfoPositiveAmount_text = dom.createTextNode(str(PubInfoPositiveAmount))
            PubInfoPositiveAmount_node.appendChild(PubInfoPositiveAmount_text)
            PubInfoNegativeAmount_text = dom.createTextNode(str(PubInfoNegativeAmount))
            PubInfoNegativeAmount_node.appendChild(PubInfoNegativeAmount_text)
            PubInfoNeutralAmount_text = dom.createTextNode(str(PubInfoNeutralAmount))
            PubInfoNeutralAmount_node.appendChild(PubInfoNeutralAmount_text)

            # 按业务统计的全国舆情数目信息区 PubInfoWorkAmountArea----省份数量ProvinceNum
            raw_Province = []
            raw_PubInfoID = []
            raw_Emotion = []
            raw_Work = []
            raw_City = []
            for m in range(1, sheet_main.nrows):
                if sheet_main.row_values(m)[2] == PubInfoWork[j]:
                    raw_Province.append(sheet_main.row_values(m)[3])
                    raw_PubInfoID.append(sheet_main.row_values(m)[0])
                    raw_Emotion.append(sheet_main.row_values(m)[1])
                    raw_Work.append(sheet_main.row_values(m)[2])
                    raw_City.append(sheet_main.row_values(m)[4])
            ProvinceNum_dict = {}
            for province in raw_Province:
                ProvinceNum_dict[province] = ProvinceNum_dict.get(province, 0) + 1
            ProvinceNum_list = list(ProvinceNum_dict.values())
            Province = list(set(raw_Province))
            Province.sort(key=raw_Province.index)

            ProvinceNum = len(Province)
            key_node = dom.createElement('ProvinceNum')
            PubInfoWorkAmountArea_node.appendChild(key_node)
            key_text = dom.createTextNode(str(ProvinceNum))
            key_node.appendChild(key_text)

            """
                按业务统计的全国舆情数目信息区 PubInfoWorkAmountArea----该类业务省份舆情发生数量信息区ProvincePubInfoWorkAmountArea
            """
            for n in range(ProvinceNum):
                ProvincePubInfoWorkAmountArea_node = dom.createElement('ProvincePubInfoWorkAmountArea')
                PubInfoWorkAmountArea_node.appendChild(ProvincePubInfoWorkAmountArea_node)

                # 该类业务省份舆情发生数量信息区ProvincePubInfoWorkAmountArea----所在省份/直辖市PubInfoDepProvince
                PubInfoDepProvince = Province[n]
                key_node = dom.createElement('PubInfoDepProvince')
                ProvincePubInfoWorkAmountArea_node.appendChild(key_node)
                key_text = dom.createTextNode(str(PubInfoDepProvince))
                key_node.appendChild(key_text)

                # 该类业务省份舆情发生数量信息区ProvincePubInfoWorkAmountArea----该类业务省份舆情总数目ProvinceInfoWorkAmount
                ProvinceInfoWorkAmount =  ProvinceNum_list[n]
                key_node = dom.createElement('ProvinceInfoWorkAmount')
                ProvincePubInfoWorkAmountArea_node.appendChild(key_node)
                key_text = dom.createTextNode(str(ProvinceInfoWorkAmount))
                key_node.appendChild(key_text)

                for p in range(len(raw_Province)):
                    if raw_Province[p] == Province[n]:
                        # 该类业务省份舆情发生数量信息区ProvincePubInfoWorkAmountArea----该类业务省份舆情信息标识区ProvincePubInfoWorkIDArea
                        ProvincePubInfoWorkIDArea_node = dom.createElement('ProvincePubInfoWorkIDArea')
                        ProvincePubInfoWorkAmountArea_node.appendChild(ProvincePubInfoWorkIDArea_node)
                        # 该类业务省份舆情信息标识区ProvincePubInfoWorkIDArea----该类业务舆情事件信息标识PubInfoID
                        PubInfoID = raw_PubInfoID[p]
                        key_node = dom.createElement('PubInfoID')
                        ProvincePubInfoWorkIDArea_node.appendChild(key_node)
                        key_text = dom.createTextNode(str(PubInfoID))
                        key_node.appendChild(key_text)

                """
                    该类业务省份舆情发生数量信息区ProvincePubInfoWorkAmountArea----该类业务按情感统计的省份舆情信息区ProvincePubInfoWorkEmotionArea
                """
                for n in range(ProvinceNum):
                    ProvincePubInfoWorkEmotionArea_node = dom.createElement('ProvincePubInfoWorkEmotionArea')
                    ProvincePubInfoWorkAmountArea_node.appendChild(ProvincePubInfoWorkEmotionArea_node)

                    # 正面舆情数量
                    PubInfoPositiveAmount_node = dom.createElement('PubInfoPositiveAmount')
                    ProvincePubInfoWorkEmotionArea_node.appendChild(key_node)
                    # 正面舆情信息标识区
                    PubInfoPositiveIDArea_node = dom.createElement('PubInfoPositiveIDArea')
                    ProvincePubInfoWorkEmotionArea_node.appendChild(PubInfoPositiveIDArea_node)
                    # 负面舆情数量
                    PubInfoNegativeAmount_node = dom.createElement('PubInfoNegativeAmount')
                    ProvincePubInfoWorkEmotionArea_node.appendChild(key_node)
                    # 负面舆情信息标识区
                    PubInfoNegativeIDArea_node = dom.createElement('PubInfoNegativeIDArea')
                    ProvincePubInfoWorkEmotionArea_node.appendChild(PubInfoNegativeIDArea_node)
                    # 中性舆情数量
                    PubInfoNeutralAmount_node = dom.createElement('PubInfoNeutralAmount')
                    ProvincePubInfoWorkEmotionArea_node.appendChild(key_node)
                    # 中性舆情信息标识区
                    PubInfoNeutralIDArea_node = dom.createElement('PubInfoNeutralIDArea')
                    ProvincePubInfoWorkEmotionArea_node.appendChild(PubInfoNeutralIDArea_node)

                    PubInfoPositiveAmount = 0
                    PubInfoNegativeAmount = 0
                    PubInfoNeutralAmount = 0
                    for l in range(1, sheet_main.nrows):
                        if sheet_main.row_values(l)[2] == PubInfoWork[j]:
                            if int(sheet_main.row_values(l)[3]) == Province[n] and int(sheet_main.row_values(l)[1]) == 1:
                                PubInfoPositiveAmount = PubInfoPositiveAmount + 1
                                key_node = dom.createElement('PubInfoID')
                                PubInfoPositiveIDArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(str(int(sheet_main.row_values(l)[0])))
                                key_node.appendChild(key_text)
                            if int(sheet_main.row_values(l)[1]) == -1:
                                PubInfoNegativeAmount = PubInfoNegativeAmount + 1
                                key_node = dom.createElement('PubInfoID')
                                PubInfoNegativeIDArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(str(int(sheet_main.row_values(l)[0])))
                                key_node.appendChild(key_text)
                            if int(sheet_main.row_values(l)[1]) == 0:
                                PubInfoNeutralAmount = PubInfoNeutralAmount + 1
                                key_node = dom.createElement('PubInfoID')
                                PubInfoNeutralIDArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(str(int(sheet_main.row_values(l)[0])))
                                key_node.appendChild(key_text)

                    PubInfoPositiveAmount_text = dom.createTextNode(str(PubInfoPositiveAmount))
                    PubInfoPositiveAmount_node.appendChild(PubInfoPositiveAmount_text)
                    PubInfoNegativeAmount_text = dom.createTextNode(str(PubInfoNegativeAmount))
                    PubInfoNegativeAmount_node.appendChild(PubInfoNegativeAmount_text)
                    PubInfoNeutralAmount_text = dom.createTextNode(str(PubInfoNeutralAmount))
                    PubInfoNeutralAmount_node.appendChild(PubInfoNeutralAmount_text)

                # 该类业务省份舆情发生数量信息区ProvincePubInfoWorkAmountArea----管辖地市的数量PubInfoDepCityNum
                row_city = []
                for q in range(1, sheet_main.nrows):
                    if sheet_main.row_values(q)[2] == PubInfoWork[j]:
                        row_city.append(sheet_main.row_values(q)[4])
                cityAmount_dict = {}
                for city in row_city:
                    cityAmount_dict[city] = cityAmount_dict.get(city, 0) + 1
                CityNum_list = list(cityAmount_dict.values())
                City = list(cityAmount_dict.keys())
                PubInfoDepCityNum = len(list(set(row_city)))

                key_node = dom.createElement('PubInfoDepCityNum')
                ProvincePubInfoWorkAmountArea_node.appendChild(key_node)
                key_text = dom.createTextNode(str(PubInfoDepCityNum))
                key_node.appendChild(key_text)

                """
                    该类业务省份舆情发生数量信息区ProvincePubInfoWorkAmountArea----该类业务地市舆情数量信息区CityPubInfoWorkAmountArea
                """
                for r in range(PubInfoDepCityNum):
                    CityPubInfoWorkAmountAreanode = dom.createElement('CityPubInfoWorkAmountArea')
                    ProvincePubInfoWorkAmountArea_node.appendChild(CityPubInfoWorkAmountAreanode)

                    # 该类业务地市舆情数量信息区CityPubInfoWorkAmountArea----所在地市PubInfoDepCity
                    PubInfoDepCity = City[r]
                    key_node = dom.createElement('PubInfoDepCity')
                    CityPubInfoWorkAmountAreanode.appendChild(key_node)
                    key_text = dom.createTextNode(PubInfoDepCity)
                    key_node.appendChild(key_text)

                    # 该类业务地市舆情数量信息区CityPubInfoWorkAmountArea----该类业务地市舆情总数目CityInfoworkAmount
                    CityInfoworkAmount_node = dom.createElement('CityInfoworkAmount')
                    CityPubInfoWorkAmountAreanode.appendChild(CityInfoworkAmount_node)
                    # 该类业务地市舆情数量信息区CityPubInfoWorkAmountArea----该类业务地市舆情信息标识区CityPubInfoWorkIDArea
                    CityPubInfoWorkIDArea_node = dom.createElement('CityPubInfoWorkIDArea')
                    CityPubInfoWorkAmountAreanode.appendChild(CityPubInfoWorkIDArea_node)
                    # 该类业务按情感统计的地市舆情信息区
                    CityPubInfoWorkEmotionArea_node = dom.createElement('CityPubInfoWorkEmotionArea')
                    CityPubInfoWorkAmountAreanode.appendChild(CityPubInfoWorkEmotionArea_node)
                    # 该类业务按情感统计的地市舆情信息区
                    # 正面舆情数量
                    PubInfoPositiveAmount_node = dom.createElement('PubInfoPositiveAmount')
                    CityPubInfoWorkEmotionArea_node.appendChild(key_node)
                    # 正面舆情信息标识区
                    PubInfoPositiveIDArea_node = dom.createElement('PubInfoPositiveIDArea')
                    CityPubInfoWorkEmotionArea_node.appendChild(PubInfoPositiveIDArea_node)
                    # 负面舆情数量
                    PubInfoNegativeAmount_node = dom.createElement('PubInfoNegativeAmount')
                    CityPubInfoWorkEmotionArea_node.appendChild(key_node)
                    # 负面舆情信息标识区
                    PubInfoNegativeIDArea_node = dom.createElement('PubInfoNegativeIDArea')
                    CityPubInfoWorkEmotionArea_node.appendChild(PubInfoNegativeIDArea_node)
                    # 中性舆情数量
                    PubInfoNeutralAmount_node = dom.createElement('PubInfoNeutralAmount')
                    CityPubInfoWorkEmotionArea_node.appendChild(key_node)
                    # 中性舆情信息标识区
                    PubInfoNeutralIDArea_node = dom.createElement('PubInfoNeutralIDArea')
                    CityPubInfoWorkEmotionArea_node.appendChild(PubInfoNeutralIDArea_node)
                    CityInfoworkAmount = 0
                    PubInfoPositiveAmount = 0
                    PubInfoNegativeAmount = 0
                    PubInfoNeutralAmount = 0
                    for s in range(1, sheet_main.nrows):
                        if sheet_main.row_values(s)[2] == PubInfoWork[j] and sheet_main.row_values(s)[4] == City[r]:

                            CityPubInfoWorkIDArea_text = dom.createTextNode(str(int(sheet_main.row_values(s)[0])))
                            CityPubInfoWorkIDArea_node.appendChild(CityPubInfoWorkIDArea_text)
                            CityInfoworkAmount = CityInfoworkAmount + 1
                            if int(sheet_main.row_values(s)[1]) == 1:
                                PubInfoPositiveAmount = PubInfoPositiveAmount + 1
                                key_node = dom.createElement('PubInfoID')
                                PubInfoPositiveIDArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(str(int(sheet_main.row_values(s)[0])))
                                key_node.appendChild(key_text)
                            if int(sheet_main.row_values(l)[1]) == -1:
                                PubInfoNegativeAmount = PubInfoNegativeAmount + 1
                                key_node = dom.createElement('PubInfoID')
                                PubInfoNegativeIDArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(str(int(sheet_main.row_values(l)[0])))
                                key_node.appendChild(key_text)
                            if int(sheet_main.row_values(l)[1]) == 0:
                                PubInfoNeutralAmount = PubInfoNeutralAmount + 1
                                key_node = dom.createElement('PubInfoID')
                                PubInfoNeutralIDArea_node.appendChild(key_node)
                                key_text = dom.createTextNode(str(int(sheet_main.row_values(l)[0])))
                                key_node.appendChild(key_text)

                    CityInfoworkAmount_text = dom.createTextNode(str(CityInfoworkAmount))
                    CityInfoworkAmount_node.appendChild(CityInfoworkAmount_text)
                    PubInfoPositiveAmount_text = dom.createTextNode(str(PubInfoPositiveAmount))
                    PubInfoPositiveAmount_node.appendChild(PubInfoPositiveAmount_text)
                    PubInfoNegativeAmount_text = dom.createTextNode(str(PubInfoPositiveAmount))
                    PubInfoNegativeAmount_node.appendChild(PubInfoNegativeAmount_text)
                    PubInfoNeutralAmount_text = dom.createTextNode(str(PubInfoPositiveAmount))
                    PubInfoNeutralAmount_node.appendChild(PubInfoNeutralAmount_text)

        # 全国舆情综合监测信息区PubInfoCountryStaticArea----省份数量ProvinceNum
        raw_Province = sheet_main.col_values(3)[1:]
        ProvinceNum = len(list(set(raw_Province)))
        Province_dict = {}
        for province in raw_Province:
            Province_dict[province] = Province_dict.get(province, 0) + 1
        ProvinceNum_list = list(Province_dict.values())
        Province = list(Province_dict.keys())
        key_node = dom.createElement('ProvinceNum')
        PubInfoCountryStaticArea_node.appendChild(key_node)
        key_text = dom.createTextNode(str(ProvinceNum))
        key_node.appendChild(key_text)

        """
            全国舆情综合监测信息区PubInfoCountryStaticArea----按省份统计的全国舆情数量信息区ProvincePubInfoAmountArea
        """
        for u in range(ProvinceNum):
            ProvincePubInfoAmountArea_node = dom.createElement('ProvincePubInfoAmountArea')
            PubInfoCountryStaticArea_node.appendChild(ProvincePubInfoAmountArea_node)

            # 按省份统计的全国舆情数量信息区ProvincePubInfoAmountArea----所在省份 / 直辖市PubInfoDepProvince
            PubInfoDepProvince = Province[u]
            key_node = dom.createElement('PubInfoDepProvince')
            ProvincePubInfoAmountArea_node.appendChild(key_node)
            key_text = dom.createTextNode(str(PubInfoDepProvince))
            key_node.appendChild(key_text)

            # 按省份统计的全国舆情数量信息区ProvincePubInfoAmountArea----省份舆情总数目ProvinceInfoAmount
            ProvinceInfoAmount = ProvinceNum_list[u]
            key_node = dom.createElement('ProvinceInfoAmount')
            ProvincePubInfoAmountArea_node.appendChild(key_node)
            key_text = dom.createTextNode(str(ProvinceInfoAmount))
            key_node.appendChild(key_text)

            # 按省份统计的全国舆情数量信息区ProvincePubInfoAmountArea----省份舆情信息标识区ProvincePubInfoIDArea
            ProvincePubInfoIDArea_node = dom.createElement('ProvincePubInfoIDArea')
            ProvincePubInfoAmountArea_node.appendChild(ProvincePubInfoIDArea_node)

            # 省份舆情信息标识区ProvincePubInfoIDArea----舆情事件信息标识PubInfoID
            for v in range(1, sheet_main.nrows):
                if sheet_main.row_values(v)[3] == Province[u]:
                    key_node = dom.createElement('PubInfoID')
                    ProvincePubInfoIDArea_node.appendChild(key_node)
                    key_text = dom.createTextNode(str(int(sheet_main.row_values(v)[0])))
                    key_node.appendChild(key_text)

            # 按省份统计的全国舆情数量信息区ProvincePubInfoAmountArea----按情感统计的省份舆情信息区ProvincePubInfoEmotionArea
            ProvincePubInfoEmotionArea_node = dom.createElement('ProvincePubInfoEmotionArea')
            ProvincePubInfoAmountArea_node.appendChild(ProvincePubInfoEmotionArea_node)

            # 按情感统计的省份舆情信息区ProvincePubInfoEmotionArea----正面舆情数量PubInfoPositiveAmount
            PubInfoPositiveAmount_node = dom.createElement('PubInfoPositiveAmount')
            ProvincePubInfoEmotionArea_node.appendChild(key_node)
            # 按情感统计的省份舆情信息区ProvincePubInfoEmotionArea----正面舆情信息标识区PubInfoPositiveIDArea
            PubInfoPositiveIDArea_node = dom.createElement('PubInfoPositiveIDArea')
            ProvincePubInfoEmotionArea_node.appendChild(PubInfoPositiveIDArea_node)
            # 按情感统计的省份舆情信息区ProvincePubInfoEmotionArea----负面舆情数量PubInfoNegativeAmount
            PubInfoNegativeAmount_node = dom.createElement('PubInfoNegativeAmount')
            ProvincePubInfoEmotionArea_node.appendChild(key_node)
            # 按情感统计的省份舆情信息区ProvincePubInfoEmotionArea----负面舆情信息标识区PubInfoPositiveIDArea
            PubInfoNegativeIDArea_node = dom.createElement('PubInfoNegativeIDArea')
            ProvincePubInfoEmotionArea_node.appendChild(PubInfoNegativeIDArea_node)
            # 按情感统计的省份舆情信息区ProvincePubInfoEmotionArea----中性舆情数量PubInfoNeutralAmount
            PubInfoNeutralAmount_node = dom.createElement('PubInfoNeutralAmount')
            ProvincePubInfoEmotionArea_node.appendChild(key_node)
            # 按情感统计的省份舆情信息区ProvincePubInfoEmotionArea----中性舆情信息标识区PubInfoNeutralIDArea
            PubInfoNeutralIDArea_node = dom.createElement('PubInfoNeutralIDArea')
            ProvincePubInfoEmotionArea_node.appendChild(PubInfoNeutralIDArea_node)

            PubInfoPositiveAmount = 0
            PubInfoNegativeAmount = 0
            PubInfoNeutralAmount = 0
            for v in range(1, sheet_main.nrows):
                if int(sheet_main.row_values(v)[3]) == Province[u] and int(sheet_main.row_values(v)[1]) == 1:
                    PubInfoPositiveAmount = PubInfoPositiveAmount + 1
                    key_node = dom.createElement('PubInfoID')
                    PubInfoPositiveIDArea_node.appendChild(key_node)
                    key_text = dom.createTextNode(str(int(sheet_main.row_values(i)[0])))
                    key_node.appendChild(key_text)
                if int(sheet_main.row_values(v)[3]) == Province[u] and int(sheet_main.row_values(i)[1]) == -1:
                    PubInfoNegativeAmount = PubInfoNegativeAmount + 1
                    key_node = dom.createElement('PubInfoID')
                    PubInfoNegativeIDArea_node.appendChild(key_node)
                    key_text = dom.createTextNode(str(int(sheet_main.row_values(i)[0])))
                    key_node.appendChild(key_text)
                if int(sheet_main.row_values(v)[3]) == Province[u] and int(sheet_main.row_values(i)[1]) == 0:
                    PubInfoNeutralAmount = PubInfoNeutralAmount + 1
                    key_node = dom.createElement('PubInfoID')
                    PubInfoNeutralIDArea_node.appendChild(key_node)
                    key_text = dom.createTextNode(str(int(sheet_main.row_values(i)[0])))
                    key_node.appendChild(key_text)

            PubInfoPositiveAmount_text = dom.createTextNode(str(PubInfoPositiveAmount))
            PubInfoPositiveAmount_node.appendChild(PubInfoPositiveAmount_text)
            PubInfoNegativeAmount_text = dom.createTextNode(str(PubInfoNegativeAmount))
            PubInfoNegativeAmount_node.appendChild(PubInfoNegativeAmount_text)
            PubInfoNeutralAmount_text = dom.createTextNode(str(PubInfoNeutralAmount))
            PubInfoNeutralAmount_node.appendChild(PubInfoNeutralAmount_text)

        # 全国舆情综合监测信息区PubInfoCountryStaticArea----舆情检索起始时间PubInfoStartTime
        PubInfoStartTime = message.find('PubInfoStartTime').text
        key_node = dom.createElement('PubInfoStartTime')
        PubInfoCountryStaticArea_node.appendChild(key_node)
        key_text = dom.createTextNode(PubInfoStartTime)
        key_node.appendChild(key_text)

        # 全国舆情综合监测信息区PubInfoCountryStaticArea----舆情检索结束时间PubInfoEndTime
        PubInfoEndTime = message.find('PubInfoEndTime').text
        key_node = dom.createElement('PubInfoEndTime')
        PubInfoCountryStaticArea_node.appendChild(key_node)
        key_text = dom.createTextNode(PubInfoEndTime)
        key_node.appendChild(key_text)

        # 全国舆情综合监测信息区PubInfoCountryStaticArea----间隔时间PubInfoIntervelTime
        PubInfoIntervelTime = '30天'
        key_node = dom.createElement('PubInfoIntervelTime')
        PubInfoCountryStaticArea_node.appendChild(key_node)
        key_text = dom.createTextNode(PubInfoIntervelTime)
        key_node.appendChild(key_text)

        # 全国舆情综合监测信息区PubInfoCountryStaticArea----间隔时间点数量IntervelTimeNum
        IntervelTimeNum = '1'
        key_node = dom.createElement('IntervelTimeNum')
        PubInfoCountryStaticArea_node.appendChild(key_node)
        key_text = dom.createTextNode(IntervelTimeNum)
        key_node.appendChild(key_text)

    xml_str = dom.toxml()
    return xml_str

    # try:
    #     with open('4.1.6司法舆情业务监测信息应答结果.xml', 'w', encoding='UTF-8') as fh:
    #         dom.writexml(fh, indent='', addindent='\t', newl='\n', encoding='UTF-8')
    #         print('写司法舆情业务监测信息应答结果完毕!')
    # except Exception as err:
    #     print('错误信息：{0}'.format(err))

if __name__=='__main__':
    create_xml('./xml_jiaoe/xml/get_6.xml')
