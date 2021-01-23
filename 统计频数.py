import jieba
import xlwt
import xlrd_compdoc_commented
from collections import Counter
filepath='C:/Users/文/Desktop/微博积极评论.txt'

with open(filepath,encoding='utf-8') as f:
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('My Worksheet')

    words=jieba.lcut(f.read())
    res=Counter(words)
    r=res.most_common(1000)
    for i in range(0,1000):#统计前1000的高频词
        s = str(r[i])
        e1 = s.split("'")[1]
        e2=s.split("'")[2].strip(")").strip(",")
        worksheet.write(i, 0, e1)
        worksheet.write(i, 1, e2)
    workbook.save('微博积极.xls')

