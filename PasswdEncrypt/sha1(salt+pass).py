import os
from hashlib import sha1
 
#fetion.com.cn:123456
#sha1：871d0a9e95aba22b4604224cfba24e605fabbc34
file1 = open(r"pwd.txt")
file2 = open("enpwd.txt","w")
for line in file1.readlines():
    line = line.strip('\n')
    salt = "fetion.com.cn:" 
    newpwd =salt+line
    s1 = sha1()
    s1.update(newpwd.encode())
    result = s1.hexdigest()
    print(result)
    file2.writelines((str(result)) + '\n')

file1.close()
file2.close()


