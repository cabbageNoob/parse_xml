'''
@Descripttion: 
@version: 
@Author: nlpir
@Date: 2020-01-03 15:35:53
@LastEditors  : nlpir
@LastEditTime : 2020-01-03 15:36:42
'''
import json

def writejson2file(data, filename):
    '''
    @Descripttion: 
    @param {data}:
    @param {filename}:
    @return: 
    '''
    with open(filename, 'w', encoding='utf8') as outfile:
        data = json.dumps(data, indent=4, sort_keys=True, ensure_ascii=False)
        outfile.write(data)


def readjson(filename):
    '''
    @Descripttion: 
    @param {filename} 
    @return: 
    '''
    with open(filename, 'rb') as outfile:
        return json.load(outfile)
