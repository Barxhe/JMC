import os
import re

db_path = "D:\dbtool"
os.chdir(db_path)
filelist=[]

for file in os.listdir():
    if re.match('REPORT_',file):
        filelist.append(file)
print(filelist[1])       