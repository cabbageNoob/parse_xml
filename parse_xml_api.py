'''
@Descripttion: 
@version: 
@Author: cjh (492795090@qq.com)
@Date: 2019-12-29 21:20:59
@LastEditors  : cjh (492795090@qq.com)
@LastEditTime : 2019-12-31 11:13:08
'''
import requests
import json, sys


def http_post(url,files):
    res = requests.post(url, files=files)
    return res.text


def main(host='127.0.0.1', file=""):
    way = sys.argv[1]
    if way in ['-h', '-H']:
        host = sys.argv[2]
        if sys.argv[3] not in ['-f', '-F']:
            raise Exception("please use -f or -F")
        file = sys.argv[4]
    elif way in ['-f', '-F']:
        file = sys.argv[2]
    url = "http://" + host + ":8001/parse_xml_api"
    files = {"file": open(file,"rb")}
    return http_post(url, files)


if __name__ == '__main__':
    # main()
    # url = "http://" + "127.0.0.1" + ":8001/parse_xml_api"
    # files = {"file": open("D:\\uploads\\origin.xml", "rb")}
    # # downloading with requests
    # r=requests.get(url)
    # with open('test.xml', 'wb') as code:
    #     code.write(r.content)
    # print(http_post(url,files))
    print(main())
