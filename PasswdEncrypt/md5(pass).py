import hashlib
import os
__author__="JE2Se"

file1 = open(r"字典位置")
file2 = open("2.txt","w")
for line in file1.readlines():
   line=line.strip('\n')
   file2.writelines(hashlib.md5(line.encode(encoding='UTF-8')).hexdigest()+'\n')

file1.close()
file2.close()