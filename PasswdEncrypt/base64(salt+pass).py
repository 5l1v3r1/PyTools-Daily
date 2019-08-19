import os
import base64

file1 = open(r"pwd.txt")
file2 = open("enpwd.txt","w")
for line in file1.readlines():
    line=line.strip('\n')
    salt = "99879094ddeee" 
    file2.writelines(str(bytes.decode(base64.b64encode((salt+line).encode('utf-8')))) + '\n')

file1.close()
file2.close()

