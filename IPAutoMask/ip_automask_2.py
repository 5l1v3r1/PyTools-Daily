#  _*_ encoding: utf-8 _*_

from IPy import IP
import os
import xlrd

ipdata = []

workbook = xlrd.open_workbook(u'test.xlsx')
sheet = workbook.sheet_by_index(0)
for i in range(sheet.nrows):
    rows = sheet.row_values(i)
    print rows[0]
    ipmask = IP(rows[0])
    a = ipmask.netmask()
    c = ipmask.net()
    newip = "%s/%s" % (c,a)
    ipdata.append(newip)
data = ''
for i in range(len(ipdata)):
    data += ipdata[i]
    if i == 0:
        data += ','
    if i % 49 != 0 and i !=0:
        data += ','
    if i % 49 == 0 and i != 0:
        data += '\n\n'

save_file = file(u'0615.txt', 'w+')
save_file.write(data)
save_file.seek(-1, os.SEEK_END)
save_file.truncate()
save_file.close()
print("子网掩码转换ok了")




