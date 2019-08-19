# -*- encoding: utf-8 -*-
'''
@File : Ipv6toArea_py3.py
@Time : 2019/08/12 15:17:10
@Author : JE2Se 
@Version : 1.0
@Contact : admin@je2se.com
@WebSite : https://www.je2se.com
'''

import requests
import tablib
import re
import xlrd
                
dataset1 = tablib.Dataset()           

def into_els(new_ip,taglocality):
    hFile = open('result.xls', "wb")
    headers = ('ipv6地址', '地区')
    dataset1.headers = headers
    dataset1.append((new_ip,taglocality))
    hFile.write(dataset1.xls)

def matchIP (new_ip):
    url = 'http://ip.zxinc.org/ipquery/?ip='
    try:
        url = url+str(new_ip)
        wbdata = requests.get(url).text
        client_re = re.compile(u'地理位置</td>\n<td style="text-align:center;background-color:#fff;height:20px">[\s\S]*?</td>')
        client1 = client_re.findall(wbdata)[0].split('>')[-2].strip('</td').strip()
        print(client1)
        into_els(new_ip,client1)
    except:
        pass

def openxls():
    workbook = xlrd.open_workbook(u'ipv6.xlsx')
    sheet = workbook.sheet_by_index(0)
    for ip in range(sheet.nrows):
        rows = sheet.row_values(ip)
        print(rows)
        matchIP(rows[0])


if __name__ == '__main__':
    openxls()
    print('Ipv6转换完成~~')