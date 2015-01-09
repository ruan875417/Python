import xlrd
import os

data=xlrd.open_workbook(r'K:\工作\助教\14级C作业统计.xls') #打开excel，路径可改
table=data.sheets()[0] #得到第一个工作表
nrows=table.nrows #获取行数
map={}#保存学号及是否交作业

for i in range(nrows):
    cell=table.cell(i,1).value
    num=str(cell)[0:9]
    if num.isdigit():
        map[num]=0

fileNum=0
for filename in os.listdir(r'K:\工作\助教\20141224'):#包含学生作业的文件夹，路径可改
    fileNum+=1                                 #文件夹中文件总数
    studentID=str(filename)[0:9]               #学生作业命名方式是前面9位为学号或者后面9位为学号
    studentID2=str(filename)[-13:-4]
    if studentID in map:
        map[studentID]+=1
    if studentID2 in map:
        map[studentID2]+=1

notAdmitNum=0#没有提交人数
keys=map.keys()
print("没交：")
for key in sorted(keys):
    #print(key,map[key])
    if map[key]==0:
        print(key)
        notAdmitNum+=1

admitMoreNum=0#提交不止一次人数
print("\n提交不止一次")
for key in sorted(keys):
    if map[key]>1:
        print(key,map[key])
        admitMoreNum+=(map[key]-1)

print("\n文件夹总文件数：%d"%fileNum)
print("重复提交次数：%d"%admitMoreNum)
print("总的提交人数：%d"%(fileNum-admitMoreNum))
print("总的没交人数：%d"%notAdmitNum)








