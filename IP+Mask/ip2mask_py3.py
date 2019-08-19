#  _*_ encoding: utf-8 _*_

import xlrd
import os
ipdata = []

workbook = xlrd.open_workbook(u'1.xlsx')
sheet = workbook.sheet_by_index(0)
for i in range(sheet.nrows):
    rows = sheet.row_values(i)
    ipdata.append('{}/255.255.255.255,'.format(rows[0].split()[0]))

data = ''
for i in range(len(ipdata)):
    print(ipdata[i])
    data += ipdata[i]
    if i % 49 == 0 and i != 0:
        data += '\n'

save_file = open(u'0614.txt', 'w+')
save_file.write(data)
save_file.seek(-1, os.SEEK_END)
save_file.truncate()
save_file.close()

