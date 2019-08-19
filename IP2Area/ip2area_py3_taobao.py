# -*- encoding: utf-8 -*-
'''
@File : ip2area_py3_taobao.py
@Time : 2019/06/25 23:21:04
@Author : JE2Se 
@Version : 1.0
@Contact : admin@je2se.com
@WebSite : https://www.je2se.com
'''


import pandas as pd
import urllib.request
import json
import os
import time

# 设置工作目录
os.chdir('/Users/je2se/Desktop/') #自行更换文件的存放目录，建议脚本和目标文件都放在一个地方

# 数据读取
df = pd.read_excel('IP.xlsx', sheet_name=0)
df_length = len(df)

# 读取'IP'列数据放入列表
ip_data = df.IP.tolist()

city_data = []
isp_data = []
country_data = []
region_data = []
country_region_city = []
for i in ip_data:
    url = ('http://ip.taobao.com/service/getIpInfo.php?ip=%s') % i
#    print(url)
    urlobject = urllib.request.urlopen(url)
    urlcontent = urlobject.read()
    res = json.loads(urlcontent)

#    print(res)
    country = res['data']['country']
    country_data.append(country)
    region = res['data']['region']
    region_data.append(region)
    city = res['data']['city']
    city_data.append(city)
    isp = res['data']['isp']
    isp_data.append(isp)

for i in range(len(ip_data)):
#    print(country_data[i] + region_data[i] + city_data[i])
    area = country_data[i] + region_data[i] + city_data[i]+isp_data[i]
    country_region_city.append(area)
#    print(country_region_city)
# 将数据整合为Dataframe类型
ipinfo = {"IP":ip_data, "城市":country_region_city}
print("ok")

result = pd.DataFrame(ipinfo)
# Dataframe输出为excel
result.to_excel('result.xlsx')

