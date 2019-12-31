# -*- coding:utf-8 -*-
import xlrd
import pandas as pd
import random
from xml.dom import minidom
import time


PubInfoID = [0,1,2,3,4,5,6,7,8]
Emotion = [1, 0, -1, 1, 0, -1, 1, 0, -1]
Work = [1, 2, 3, 1, 2, 3, 1, 2, 3]
Province = [3, 2, 1, 1, 2, 3, 2, 3, 1]
City = ['d', 'b', 'a','a', 'b', 'c', 'b', 'e', 'a']


all_dict_main={'PubInfoID':PubInfoID,
               'Emotion':Emotion,
               'Work':Work,
               'Province':Province,
               'City':City}
main_order = ['PubInfoID', 'Emotion', 'Work', 'Province', 'City']
df_main = pd.DataFrame(all_dict_main)
df_main = df_main[main_order]

with pd.ExcelWriter('new6.xls') as Writer:
    df_main.to_excel(Writer,'Sheet1',index=False, encoding='utf_8_sig')



