import  requests
import  json
import pprint
import xlwt
import re
import urllib.request
from bs4 import BeautifulSoup
FindLink=re.compile(r'href="//www.bilibili.com/video/(.*?)?from=search" target="_blank" title="')
head={'user-agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.132 Safari/537.36'}
for j in range(1,5):
    response = urllib.request.Request(
        'https://search.bilibili.com/all?keyword=%E7%96%AB%E6%83%85&order=totalrank&duration=0&tids_1=0&page='+str(j), headers=head) #B站搜索“疫情”页面
    r = urllib.request.urlopen(response)
    html = r.read().decode("utf-8")
    soup = BeautifulSoup(html, 'html.parser')
    for item in soup.find_all('div'):
        item = str(item)
        link = re.findall(FindLink, item)
        if (len(link) > 1):
            i = 0
            while (i < len(link)):
                print(link[i].strip('?'))  # 打印bv号
                i += 2
            break
    print('--------------------------------------------------') #分割线
