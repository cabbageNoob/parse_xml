'''
@Descripttion: 
@version: 
@Author: nlpir
@Date: 2020-01-03 16:20:37
@LastEditors  : nlpir
@LastEditTime : 2020-01-03 16:34:18
'''
import os

pwd_path = os.path.abspath(os.path.dirname(__file__))
folder_name = os.path.join(pwd_path, '../static/json')

# 获取那个文件夹中所有的文件名字：
file_names = os.listdir(folder_name)
# 进入folder_name文件夹
os.chdir(folder_name)

def edit_file_name(file_names):
    for name in file_names:
        new_name = name[:-7]+'-'+name[-7:]
        os.rename(name, new_name)

if __name__ == '__main__':
    edit_file_name(file_names)
