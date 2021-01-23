# encoding: utf-8
import xlsxwriter
import xlrd_compdoc_commented
import xlwt
import re


def read_excel():
    # 打开文件
    Allstr = ''
    ji=0
    for ite in {0,1,2,3}:
        workBook = xlrd_compdoc_commented.open_workbook('NewsWeibo'+str(ite)+'.xls')
        Work=xlrd_compdoc_commented.open_workbook('D:/新建文件夹/2368771102/FileRecv/weibo0_1_2_3评论筛选.xlsx')
        Sheet=Work.sheet_by_index(ji)
        a1=[None]*(Sheet.nrows-1)
        a2=[None]*(Sheet.nrows-1)
        for i in range(0,Sheet.nrows-1):
            a1[i]=int(Sheet.row_values(i+1)[0].split(' ')[0])
            a2[i]=int(Sheet.row_values(i+1)[0].split(' ')[1])
        print(a1)
        print(a2)
        # writebook=xlsxwriter.Workbook('C:/Users/文/Desktop/123.xlsx')
        # tb=writebook.add_worksheet('sheet1')
        # tb.write_string(0,0,"first")
        sheet_name = workBook.sheet_names()[0]
        sheet = workBook.sheet_by_index(0)
        # row=sheet.row_values(0)
        # nrows=sheet.nrows;
        # yiqingnum=0
        # zidian={"疫情","新增","患者"}
        # for i in zidian:
        #     obj=i
        #     for num in range(1, nrows):
        #         row = sheet.row_values(num)
        #         if (re.search(obj, row[2])):
        #             yiqingnum += 1
        #     print(obj, "出现次数=", yiqingnum)
        #
        for num in range(0,len(a1)):
            row = sheet.row_values(a1[num])
            Allstr=Allstr+row[a2[num]+3]
            # Allstr = Allstr + row[3] + '\n' + "评论     " + '编号' + str(num) + ' 1 ' + row[
            #     4] + '\n            ' + '编号' + str(num) + ' 2 ' + row[5] + '\n            ' + '编号' + str(num) + ' 3 ' + \
            #          row[6] + '\n            ' + '编号' + str(num) + ' 4 ' + row[7] + '\n            ' + '编号' + str(
            #     num) + ' 5 ' + row[8] + '\n\n'
        ji+=1
    f=open("C:/Users/文/Desktop/微博消极评论.txt",'a',encoding='utf-8')
    f.write(Allstr)
    f.close()

read_excel()