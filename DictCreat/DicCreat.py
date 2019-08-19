""" 
字典生成脚本

"""
import itertools as its
import os
__auther__="JE2Se"

words = "1234568790"        #可设置为"123456789abcdefghijklmnopqrstuvwxyz大写...
password= its.product(words,repeat=6)   #repeat代表着你要创建的密码位数
dic = open("pass.txt","a")

for i in password:
    dic.write("".join(i))
    dic.write("".join("\n"))

dic.close()

