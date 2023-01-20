import msvcrt
import  pandas as pd
import os
from tqdm import tqdm

#读取花名册学生信息
file_path = r"花名册.xlsx"  #源文件
res_path = r"未交名单.xlsx" #目标文件
df = pd.DataFrame(pd.read_excel(file_path))
#print(df["NAME"])
#print("转换为字典")
#print(df[["ID","NAME"]].to_dict(orient='dict'))
ID = list(df["ID"]) #获得所有人的学号
student = df.set_index("ID")["NAME"].to_dict() #ID为键，NAME为值
#print(student)

#获取文件夹下所有文件名字
base_path = os.getcwd() #当前位置的绝对路径
print("当前工作目录："+base_path)
print("正在读取当前目录下所有文件名...")
fileNames = os.listdir(base_path)
print("文件个数：%d"%len(fileNames))


unStudentID = list() #未交作业的小朋友的学号
for id in tqdm(ID):
    #print(id)

    if not any(str(id) in filename for filename in fileNames): #如果所有文件名中都没有这个学号
        unStudentID.append(id) #记录下来

unStudentID.sort() #排个序
unStudent = {str(id):student[id] for id in unStudentID}



df_w = pd.DataFrame(list(unStudent.items()),columns=["ID","NAME"])
print("未交文件同学如下：")
print(df_w)
df_w.to_excel(res_path)
print("应交：",len(ID))
print("已交：",len(ID)-len(unStudentID))
print("未交：",len(unStudentID))
print(r"统计完成，结果见文件”未交名单.xlsx“")

print("请按任意键退出")
ord(msvcrt.getch())







