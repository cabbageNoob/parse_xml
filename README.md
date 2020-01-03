<!--
 * @Descripttion: 
 * @version: 
 * @Author: nlpir
 * @Date: 2019-12-29 21:48:21
 * @LastEditors  : nlpir
 * @LastEditTime : 2020-01-03 15:52:09
 -->

## python 读取xml文件并封装成服务

### 启动服务
运行manager.py文件

### 调用服务方式
- 方式一：直接在浏览器打开[网址](http://127.0.0.1:8001/parse_xml_api)，
按要求选择xml文件再上传

- 方式二：打开一个cmd窗口，输入如下类似命令
```
curl 127.0.0.1:8001/parse_xml_api -X POST -F "file"="@D:\uploads\origin.xml"
```

- 方式三：执行parse_xml_api.py文件，例如下面命令
```
python parse_xml_api.py -h 127.0.0.1 -f ./static/xml/origin.xml
```
-h或-H后输入ip地址；-f或-F后输入xml文件所在地址

### 解析json文件
- 方式：打开一个cmd窗口，输入如下类似命令
```
curl 127.0.0.1:8001/parse_json -X POST -F "time"="2016-04-0411"
```

### TODO
- 在xml_process.py文件中的处理xml文件的get_attrib_value()函数，我只写了部分xml文件处理方法，相关业务逻辑还是不太明白，需要补全逻辑