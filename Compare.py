import math

import xlrd_compdoc_commented
import xlwt
workBook1=xlrd_compdoc_commented.open_workbook('微博积极.xls')
SheetJ=workBook1.sheet_by_index(0)
workBook2=xlrd_compdoc_commented.open_workbook('微博消极.xls')
SheetX=workBook2.sheet_by_index(0)
Jiji=[None]*500
NumJ=[0]*500
NumJ1=[None]*500
for i in range(0,500):
    rowj=SheetJ.row_values(i)
    Jiji[i]=str(rowj[0])
    NumJ1[i]=int(rowj[1])
Xiaoji=[None]*500
NumX=[None]*500
NumX1=[None]*500
for j in range(0,500):
    rowX=SheetX.row_values(j)
    Xiaoji[j]=str(rowX[0])
    NumX[j]=j+1
    NumX1[j]=int(rowX[1])
for item in range(0,len(Xiaoji)-1):
    for i1 in range(0,len(Jiji)-1):
        if(Xiaoji[item]==Jiji[i1]):
            n=int(NumX1[item])
            m=int(NumJ1[i1])
            NumJ[item]=(i1-item)+(n-m)
            break

workbook = xlwt.Workbook(encoding='utf-8')
worksheet1 = workbook.add_sheet('积极词')
worksheet2=workbook.add_sheet("消极词")
worksheet1.write(0,0,"积极词汇")
worksheet1.write(0,1,"按排名积极系数")
worksheet2.write(0,0,"消极词汇")
worksheet2.write(0,1,"按排名消极系数")
i=1
j=1
for i1 in range(0,499):
    if((i1<150 and NumJ[i1]>=3) or NumJ[i1]>15):
        worksheet2.write(i,0,Xiaoji[i1])
        worksheet2.write(i,1,NumJ[i1])
        i=i+1
    if ((i1 < 150 and NumJ[i1] <= -3) or NumJ[i1] < -15):
        worksheet1.write(j, 0, Xiaoji[i1])
        worksheet1.write(j, 1, -NumJ[i1])
        j = j + 1
workbook.save("积极消极词汇排名.xls")
print("积极",Jiji)
print('num',NumJ1)
print(NumJ)
print('消极',Xiaoji)
print(NumX)
print(NumX1)