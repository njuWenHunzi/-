import xlrd_compdoc_commented
import xlwt
import math
workBook=xlrd_compdoc_commented.open_workbook('D:/新建文件夹/2368771102/FileRecv/筛选后的微博消极积极词频统计.xlsx')
sheetJi=workBook.sheet_by_index(0)
sheetXiao=workBook.sheet_by_index(1)
JijiNum=[0]*12
XiaojiNum=[0]*12
ComNumJi=[0]*12
ComNumXiao=[0]*12
Time=[None]*12
for o in range(0,12):
    ComNumJi[o]=int(sheetJi.row_values(sheetJi.nrows-1)[o+2])
    ComNumXiao[o]=int(sheetXiao.row_values(sheetXiao.nrows-1)[o+2])
    Time[o]=sheetJi.row_values(0)[o+2]
print(ComNumJi)
# print(sheetJi.row_values(sheetJi.nrows-1)[2])
for i in range(1,100):
    key = sheetJi.row_values(i)
    key[1]=math.log(key[1])
    print(key)
    for j in range(0,12):
        JijiNum[j]+=int(key[j+2])*key[1]/ComNumXiao[j]
for i1 in range(1,100):
    key1 = sheetXiao.row_values(i1)
    key1[1]=math.log(key1[1])
    for j1 in range(0,12):
        XiaojiNum[j1]+=int(key1[j1+2])*key1[1]/ComNumXiao[j1]
Work=xlwt.Workbook(encoding='utf-8')
Sheet=Work.add_sheet('My Sheet')
Sheet.write(0,0,"时间")
Sheet.write(1,0,"积极指数")
Sheet.write(2,0,"消极指数")
for i2 in range(1,13):
    Sheet.write(0,i2,Time[i2-1])
    Sheet.write(1,i2,JijiNum[i2-1])
    Sheet.write(2,i2,XiaojiNum[i2-1])
Work.save("筛选后微博消极积极指数.xls")

print(JijiNum)
print(XiaojiNum)
