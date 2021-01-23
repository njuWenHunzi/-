import  requests
import  json
import pprint
import xlwt
import re
r=requests.get('https://api.bilibili.com/x/v2/reply?jsonp=jsonp&pn=1&type=12&oid=4620519')
data=json.loads(r.text)
# pprint.pprint(data['data']['replies'])
workbook=xlwt.Workbook(encoding='utf-8')
worksheet=workbook.add_sheet('B站评论')
worksheet.write(0,0,'评论')
MainReplyNum=1

for i in data['data']['replies']:
    pprint.pprint(i)
    worksheet.write(MainReplyNum,0,i['content']['message'])
    MainReplyNum+=1
    for j in i['replies']:
        pprint.pprint(j['content']['message'])
        worksheet.write(MainReplyNum,0,j['content']['message'])
        MainReplyNum += 1

workbook.save('B站评论爬取.xls')
