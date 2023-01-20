import msvcrt
import  pandas as pd
import os
from tqdm import tqdm

# 读取paper指导
def getPaper(path):
    paperlist = list()
    file = open(path, encoding="utf-8")
    while True:
        line = file.readline().rstrip()
        #print(line[0])
        if line:
            if line[0] != '#':
                paperlist.append(line)
        else:
            break
    file.close()
    return paperlist


#读取花名册学生信息
file_path = r"花名册.xlsx"  #花名册文件
source_path = r"青年大学习已完成名单.txt" #已完成人员名单
print("读取花名册...")
df = pd.DataFrame(pd.read_excel(file_path))
ID = list(df["ID"]) #获得所有人的学号
NAMES = list(df['NAME'])
print("读取已完成名单...")
finishNames = getPaper(source_path) #读取文件
unfinishNames = list() #未完成名单
print("统计未完成名单...")
for name in NAMES:
    if not any(name in item for item in finishNames): #不在完成名单中
        unfinishNames.append(name) #添加到未完成列表中
print("未完成名单如下：")
for name in unfinishNames:
    print("    "+name)
print("未完成共%d"%len(unfinishNames))
print("\n未完成名单已保存到文件：”未完成.txt")
if unfinishNames: #list不为空-写入到文件中
    with open(r"未完成.txt",'w') as f:
        for name in unfinishNames:
            f.write(name+"\n")
else:
    with open("未完成.txt", 'w') as f:
        f.write("")
        f.close()
print("\n请按任意键退出")
ord(msvcrt.getch())
