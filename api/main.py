# 사전에 pip install notion 필요

from notion.client import *
from notion.block import *
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import pandas as pd
import os
import threading

# 검색 키워드
keyword = ['IPO']


# 크롤링 하는 사이트
source = {'더벨' : 'http://www.thebell.co.kr/free/content/Search.asp?keyword=',
        '매일경제' : 'http://find.mk.co.kr/new/search.php?page=news',
        '서울경제' : 'https://www.sedaily.com/Search/Search/SEList?Page=',
        '기업공시채널' : ''}


def get_filename():
    now = datetime.now()
    filename = "%s-%02d-%02d Crawling Result" % (now.year, now.month, now.day)
    return filename

def get_collection_schema():
    return {
        "title" : {"name" : "title", "type" : "text"},
        "url" : {"name" : "url", "type" : "url"},
        "crawlingdate" : {"name" : "crawlingdate", "type" : "text"},
        "publisheddate" : {"name" : "publisheddate" , "type" : "text"},
        "source" : {"name" : "source", "type" : "text"},
        "keyword" : {"name" : "keyword", "type" : "text"}
    }
    
# 더벨 사이트 크롤링
def theBell():
    theBellNewsList = []
    for onekeyword in keyword:
        searchlink = source['더벨'] + onekeyword
        for i in range(1, 3): #페이지 3개 크롤링
            searchlink += '&page=' + str(i)
            req = requests.get(searchlink)
            html = req.text
            soup = BeautifulSoup(html, 'html.parser')

            rawnews = soup.select('div.listBox > ul > li')

            for onenews in rawnews:
                title = onenews.select('a')[0].get('title')
                newsurl= onenews.select('a')[0].get('href')
                newsurl = 'https://www.thebell.co.kr/free/content/' + newsurl
                published_date = onenews.select('.date')[0].text
                now = datetime.now()
                create_date = "%s-%02d-%02d %02d:%02d:%02d" % (now.year, now.month, now.day, now.hour, now.minute, now.second)
                
                crawling_one_news = {'기사 제목' : title,
                    '기사 링크' : newsurl,
                    '기사 날짜' : published_date,
                    '키워드' : onekeyword,
                    '출처' : 'TheBell',
                    '크롤링 날짜' : create_date}
                
                theBellNewsList.append(crawling_one_news)
    
    return theBellNewsList #리스트를 리턴한다!


# 매일경제 사이트 크롤링
def MK():
    MK_NewsList = []
    for onekeyword in keyword:
        searchlink = source['매일경제'] + '&s_keyword=' + onekeyword
        for i in range(1, 3): #페이지 3개 크롤링
            searchlink += '&pageNum=' + str(i)
            req = requests.get(searchlink)
            soup = BeautifulSoup(req.content.decode('euc-kr', 'replace'))

            rawnews = soup.select('.sub_list')

            for onenews in rawnews:
                title = onenews.select('span > a')[0].text
                newsurl = onenews.select('span > a')[0].get('href')
                published_date = onenews.select('span.art_time')[0].text
                published_date = published_date[-24:]
                published_date = published_date[0:4] + '-' + published_date[6:8] + '-' + published_date[10:12] + ' ' + published_date[-9:]
                now = datetime.now()
                create_date = "%s-%02d-%s %02d:%02d:%02d" % (now.year, now.month, now.day, now.hour, now.minute, now.second)

                crawling_one_news = {
                    '기사 제목' : title,
                    '기사 링크' : newsurl,
                    '기사 날짜' : published_date,
                    '키워드' : onekeyword,
                    '출처' : '매일경제',
                    '크롤링 날짜' : create_date
                }

                MK_NewsList.append(crawling_one_news)
    
    return MK_NewsList

                

# 서울경제 사이트 크롤링
def sedaily():
    source['서울경제'] = 'https://www.sedaily.com/Search/Search/SEList?Page='
    sedailyNewsList = []
    for onekeyword in keyword:
        for i in range(1, 3):
            searchlink = source['서울경제'] + str(i) + '&scDetail=&scOrdBy=0&catView=AL&scText=' + onekeyword + '&scPeriod=6m&scArea=tc&scTextIn=&scTextExt=&scPeriodS=&scPeriodE=&command=paging&_=1579667403619'
            req = requests.get(searchlink)
            html = req.text
            soup = BeautifulSoup(html, 'html.parser')

            rawnews = soup.select('ul.news_list > li > div')

            for onenews in rawnews:
                title = onenews.select('a')[0].text
                newsurl = 'https://www.sedaily.com' + onenews.select('a')[0].get('href')
                published_date = onenews.select('.letter')[0].text
                now = datetime.now()
                create_date = "%s-%02d-%s %02d:%02d:%02d" % (now.year, now.month, now.day, now.hour, now.minute, now.second)

                crawling_one_news = {
                    '기사 제목' : title,
                    '기사 링크' : newsurl,
                    '기사 날짜' : published_date,
                    '키워드' : onekeyword,
                    '출처' : '서울경제',
                    '크롤링 날짜' : create_date
                }

                sedailyNewsList.append(crawling_one_news)
    
    return sedailyNewsList

def get_today_news():
    newslist = []
    print("Gathering Keyword News.......")
    newslist = theBell() + sedaily() + MK()
    print("Just Ignore waring message ^_^")
    return newslist


def main():
    
    #login
    token_v2 = '459e602e06823930253e2544399eb383d6a1dd46a5b51ac6a0476557460f02c1c8563b1bade2b287e862050cc34dae987b4bb3c2bfcb4e0e1640c22592edddceb2a254a4b2961d0b502f92055d69'
    client = NotionClient(token_v2=token_v2)

    #daily news keyword crawling page
    url = 'https://www.notion.so/sociallistening/Daily-News-Keyword-Monitoring-30e047e776c44fa2a1ea6419a132817d'
    page = client.get_block(url)

    today_crawling_result = page.children.add_new(CollectionViewPageBlock)
    today_crawling_result.collection = client.get_collection(
        client.create_record("collection", parent=today_crawling_result, schema=get_collection_schema())
    )
    today_crawling_result.title = get_filename()

    todaynewslist = get_today_news()

    for onenews in todaynewslist:
        row = today_crawling_result.collection.add_row()
        row.title = onenews['기사 제목']
        row.source = onenews['출처']
        row.publisheddate = onenews['기사 날짜']
        row.crawlingdate = onenews['크롤링 날짜']
        row.url = onenews['기사 링크']
        row.keyword = onenews['키워드']
    
    view = today_crawling_result.views.add_new(view_type='table')
    print("Finish to Upload News to Notion Page")

    a = input("\n>>>>If you want to export data to excel, PRESS 0 \t\t>> ")
    if a == '0':
        print("Start Exporting!")
        df = pd.DataFrame(todaynewslist)
        filepath = os.path.dirname(os.path.abspath('__file__'))
        filename = filepath + '\\' + get_filename() + '.xlsx'
        print(filename)

        try:
            df.to_excel(filename, index=False)
            print("File Save Success!")
            print("File Created at " + filename)
        except:
            print("File create Error!")
    
    else:
        print("Have a Good Day!")
    
    
    #24h = 86400s
    #print("hihi")
    threading.Timer(86400, main).start()
    
#main
main()
