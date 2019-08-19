""" 
字典生成脚本

"""

import itertools as its
import os
__auther__="JE2Se"
ipdemo = "12.1.91."

dic = open("ip.txt","a")
for ip in range(254):
    newip = ipdemo + str(ip)
    dic.write("".join(newip))
    dic.write("".join("\n"))

dic.close()


