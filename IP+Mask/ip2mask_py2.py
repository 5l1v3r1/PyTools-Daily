#  _*_ encoding: utf-8 _*_

import xlrd
import os
import string
ipdata = []

workbook = xlrd.open_workbook(u'test.xlsx')
sheet = workbook.sheet_by_index(0)
for i in range(sheet.nrows):
    rows = sheet.row_values(i)
    ipdata.append('{}/255.255.255.255'.format(rows[0].split()[0]))

data = ''
for i in range(len(ipdata)):
    data += ipdata[i]
    if i == 0:
        data += ','
    if i % 49 != 0 and i !=0:
        data += ','
    if i % 49 == 0 and i != 0:
        data += '\n\n'
    
save_file = file(u'0614.txt', 'w+')
save_file.write(data)
save_file.seek(-1, os.SEEK_END)
save_file.truncate()
save_file.close()
print("子网掩码转换ok了")