__author__="JE2Se"
import hashlib
import os

file1 = open(r"字典位置")
file2 = open("2.txt","w")
for line in file1.readlines():
   line=line.strip('\n')
   file2.writelines(hashlib.sha256(line.encode("utf-8")).hexdigest()+'\n')

file1.close()
file2.close()