# -*- encoding: utf-8 -*-
'''
@File : ip2area_py3_chinaz.py
@Time : 2019/06/25 23:26:58
@Author : JE2Se 
@Version : 1.0
@Contact : admin@je2se.com
@WebSite : https://www.je2se.com
'''

import sys
import os
import requests
from bs4 import BeautifulSoup
import tablib
import socket
import re

path = "/Users/je2se/Desktop/"    # 存放路径
filename = "ip"                       # 文件名称
dataset1 = tablib.Dataset()           # 数据集合

ip_list = []                          # IP列表

# 写XLS
def into_els(new_ip,taglocality):
    headers = ('ip', '地区')        # 首行字段
    dataset1.headers = headers
    dataset1.append((new_ip,taglocality))


#匹配出IP地址函数
def matchIP (new_ip):
    url = "http://ip.tool.chinaz.com/"
    try:
        url = url+str(new_ip)
    except:
        pass

    ## 根据传入的IP地址截取出地区
    wbdata = requests.get(url).text
    soup = BeautifulSoup(wbdata, 'lxml')
    for tag in soup.find_all('span', class_='Whwtdhalf w50-0'):
        tag_extractl = tag.get_text().encode('utf-8')
        if tag_extractl.find(b"IP\xe7\x9a\x84\xe7\x89\xa9\xe7\x90\x86\xe4\xbd\x8d\xe7\xbd\xae"):     #过滤掉【IP的物理位置】这个字符
            print("%s||%s" % (new_ip,tag.get_text()))        #输出域名,IP，地区
            into_els(new_ip,tag.get_text())                        #写数据到数据集合中


#读取文件函数
def read_file(file_path):
    if not os.path.exists(file_path):
        sys.exit(0)
    else:
        with open(file_path, 'r') as source:
            for line in source:
                ip_list.append(line.rstrip('\r\n').rstrip('\n'))
    for ip in ip_list:
        matchIP(ip)

    hFile = open(path + filename + '.xls', "wb")

    hFile.write(dataset1.xls)
    hFile.close()


if __name__ == '__main__':
    file_str="/Users/je2se/Desktop/ip.txt"
    read_file(file_str)