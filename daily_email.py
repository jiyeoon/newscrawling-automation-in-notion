import os
from datetime import datetime
import pandas as pd 
import numpy as np 
import win32com.client


def getDataFrame():
    dir = os.path.dirname(os.path.abspath('__file__'))
    filepath = dir + r'\asset\static\Korea GB Account CRM Extraction File_20200212.xlsx'
    print(filepath)

    try:
        df = pd.read_excel(filepath)
        return df
    except:
        print("File open error!")

def getDDE(aename):
    print("Get DDE!")
    print(aename)
    if aename == 'Jihun Kang' or aename == 'Steven Park':
        return 'Diane Kim'
    elif aename == 'Jongsu Hwang' or aename == 'Kaon Kim':
        return 'Jeegoo Kang'
    elif aename == 'Jungrye Jung' or aename == 'Alyssa Lee':
        return 'Erin Lim'
    elif aename == 'Hweejae Yim' or aename == 'Tae Seung Kim' or aename == 'Young Ae Kim':
        return 'Jeongha Lim'
    elif aename == 'Junhyung Byun' or aename == 'Sinjung Choi' or aename == 'Woojoon Chu' or aename == 'Yangkyung Lee' or aename == 'Jungsun Yoon':
        return 'Lee Jae'
    else:
        return 'N/A'

def search(search_term, df):
    #df = getDataFrame()
    onecompany = {}
    if len(df[df['ORG Name (KOR)'] == search_term]) >= 1:
        newdf = df[df['ORG Name (KOR)'] == search_term]
        print(newdf)
        onecompany['Name'] = search_term
        onecompany['BP#'] = int(newdf['ORG ID (BP#)'])
        onecompany['2018 turnover'] = int(newdf['Revenue Local Currency (2018)']/1000000)
        onecompany['2018 turnover'] = format(onecompany['2018 turnover'], ',')
        onecompany['AE'] = newdf['Name_For_Account_Owner'] + " " + newdf['Last_Name_For_Account_Owner']
        onecompany['AE'] = onecompany['AE'].values[0]
        onecompany['Industry'] = newdf['SAP_Master_Code_Text'].values[0]
        onecompany['Buying Classification'] = newdf['Gtm_Regional_Buying_Classification_Text'].values[0]
        onecompany['DDE'] = getDDE(onecompany['AE'])

        print(onecompany)
        return onecompany
    else:
        print("There is No " + search_term + " !!!!!!")
        onecompany['Name'] = search_term
        onecompany['BP#'] = 'N/A'
        onecompany['2018 turnover'] = 'N/A'
        onecompany['AE'] = 'N/A'
        onecompany['DDE'] = getDDE(onecompany['AE'])
        onecompany['Industry'] = 'N/A'
        onecompany['Buying Classification'] = 'N/A'
        ## 여기는.. 키스라인 데이터를 가져와야한다..~!!!


def getFromKISLINE(search_term):
    onecompany = {}


def getCompanyList():
    df = getDataFrame()
    search_term = input("검색할 기업명을 입력하세용::::>>>> \t")
    companylist = []
    while search_term != "":
        onecompany = search(search_term, df)
        news_title = input("\n기사 제목을 입력하세요 :::>>> \t")
        onecompany['News Title'] = news_title
        news_url = input("\n뉴스 url을 입력하세요 :::>>> \t")
        onecompany['News Url'] = news_url
        onecompany['Comment'] = []
        print("Company Comment!!")
        while True:
            input_str = input(">")
            if input_str == "":
                break
            else:
                onecompany['Comment'].append(input_str)
        
        print(onecompany)
        companylist.append(onecompany)
        search_term = input("\n---------------\n검색할 기업명을 입력하세요:::>>> \t")
    print("Ok Bye..~~~~")
    print(companylist)

    return companylist

def getHTML():
    now = datetime.now()

    html_header_text = """
        <!DOCTYPE html>
        <html>
        <head>
        <title>[Daily News Clipping] %s년 %s월 %s일 주요 Keyword Monitoring</title>
        <link href="https://fonts.googleapis.com/css?family=Noto+Sans+KR&display=swap" rel="stylesheet">
        <style>
            body{
                font-family: 'Noto Sans KR', sans-serif;
            }
        </style>
        </head>

        <body>
        <p>안녕하세요? </br>
        Commercial Sales의 Social Listening 담당 이지연입니다. </p>
        <p> %s년 %s월 %s일 주요 Keyword Monitoring 결과를 안내드립니다. </br>
        <a href='/'>이곳</a>을 클릭하시면 더 많은 뉴스를 확인하실 수 있습니다.
        </p>
    """ % (now.year, now.month, now.day, now.year, now.month, now.day)

    companylist = getCompanyList()

    html_body_text = ""
    for onecompany in companylist:
        html_body_text += "<br><p><h3><a href=\'%s'>%s</a></h3>" % (onecompany['News Url'], onecompany['News Title'])
        html_body_text += "<strong>* %s / 2018 매출액 : %s / Industry : %s / Buying Classification : %s </strong><br>" % (onecompany['Name'], onecompany['2018 turnover'], onecompany['Industry'], onecompany['Buying Classification'])
        html_body_text += "<string>* BP# : %d / AE : %s / DDE : %s </strong><br>" % (onecompany['BP#'], onecompany['AE'], onecompany['DDE'])
        for onestring in onecompany['Comment']:
            html_body_text += "* " + onestring + "<br>"
        html_body_text += "</p>"
    
    html_footter_text = "<p>감사합니다.<br>이지연 드림</p></body></html>"

    html_text = html_header_text + html_body_text + html_footter_text

    return html_text
    """
    html_file = open('%s-%02d-%02d Daily News Clipping Email.html' % (now.year, now.month, now.day), 'w')
    html_file.write(html_text)
    html_file.close()
    """

def main():
    print("func main")

def sendMail():
    now = datetime.now()
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = "[Daily News Clipping] %s년 %s월 %s일 주요 Keyword News" % (now.year, now.month, now.day)
    newMail.HTMLBody = getHTML()
    newMail.To = "jiyeon.lee@sap.com"
    newMail.Send()
    


if __name__ == "__main__":
    main()
