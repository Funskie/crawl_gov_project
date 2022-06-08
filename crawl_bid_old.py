# coding: utf-8
# In[ ]:
import requests
from selenium import webdriver
from pyquery import PyQuery as pq
from selenium.webdriver.common.keys import Keys
from apscheduler.schedulers.blocking import BlockingScheduler
import time
from time import sleep
import smtplib
from email.mime.text import MIMEText
from email import encoders, utils
from email.mime.multipart import MIMEMultipart
bidHead = ['關鍵字','機關名稱','案號','標案名稱','傳輸次數','公告日期','截標日期','預算金額','標案連結']
scheduler = BlockingScheduler()
mailTo = ['chongjing3370@gmail.com', 'tusty9292@gmail.com']

def sendMail(mailList, mailBody):
    #本區設定smtp server
    # smtp_server = 'smtp.office365.com:587'
    me = 'tusty9292@gmail.com'
    mail = MIMEMultipart()
    mail['Subject'] = "政府採購公告網 "+str(time.strftime("%Y/%m/%d-%H%M"))
    mail['From'] = 'tusty9292@gmail.com'
    mail['Date'] = utils.formatdate(localtime = 1)
    mail['Message-ID'] = utils.make_msgid()
    mail['Content-Type'] = "text/calendar; charset=utf-8"
    sever = smtplib.SMTP('smtp.gmail.com',587)
    body = MIMEText(mailBody, 'html', "utf-8")
    body.set_charset("utf-8")
    mail.attach(body)
    #本區為實際把信寄出去
    #必須要先ehlo再stattls
    sever.ehlo()
    sever.starttls()
    sever.login('chongjing3370@gmail.com', 'Cj16684351')
    sever.sendmail(me, mailList, mail.as_string())
    sever.quit()

def mailBody(head,bidList):
    a = []
    rowSpan = {}
    for s in bidList:
        a.append(s[0])
    b = set(a)
    for s in b:
        count = a.count(s)
        rowSpan[s] = count+1
    mailText = '<table border="1"><tbody><tr>'
    for i in range(0,len(head)): mailText += "<th><b>" + str(head[i]) + "</b></th>"
    for s in range(0,len(bidList)):
        if s > 0:
            if bidList[s][0] != bidList[s-1][0]: mailText += '<tr><th rowspan="'+str(rowSpan[bidList[s][0]])+'">'+bidList[s][0]+'</tr>'
        else: mailText += '<tr><th rowspan="'+str(rowSpan[bidList[s][0]])+'">'+bidList[s][0]+'</tr>'
        mailText += "<tr>"
        for i in range(1,len(bidList[s])):
            if i < 7: mailText += "<td>" + str(bidList[s][i]) + "</td>"
            elif i == 7: mailText += '<td align="right">' + str(bidList[s][i]) + "</td>"
            else: mailText += '<td><a href="'+str(bidList[s][i])+'"><b>LINK</b><a></td>'
        mailText += "</tr>"
    mailText += "</tbody></table>"
    return(mailText)

def crawlBid(keywords, orgKW):
    driver = webdriver.PhantomJS("phantomjs.exe")
    bidList = []
    bidList_1 = []
    linkList = []

    #Table標題列
    for s in keywords:
        bidSystemUrl = "https://web.pcc.gov.tw/tps/pss/tender.do?method=goSearch&searchMode=common&searchType=basic"
        browse = driver.get(bidSystemUrl)
        homepage = driver.page_source
        query = driver.find_element_by_name("tenderName")
        query.clear()
        query.send_keys(s)
        if s == "智慧":
            query.send_keys(Keys.RETURN)
        else:
            queryCate = driver.find_element_by_css_selector("#radProctrgCate3")
            queryCate.click()
            query.send_keys(Keys.RETURN)
        sleep(1) #等頁面回傳
        bidpage = driver.page_source
        querypage = pq(bidpage)
        querypage.make_links_absolute(base_url=bidSystemUrl)
        #tbody下的所有onmouseover='overcss(this);'的物件，屬性選擇器用法很複雜要小心
        #如果關鍵字搜尋不到就不會有此物件，會出現onmouseover="overcss(this)"的物件，少一個分號;，物件不一樣
        checkQuery = querypage("[onmouseover='overcss(this);']").text()  
        #print('===關鍵字：',s,"\n")
        if checkQuery:
            for bidCases in querypage("tbody [onmouseover='overcss(this);']").items():
                bidList_ins = []
                for eachBid in bidCases("td").items():
                    bidList_ins.append(eachBid.text())
                bidLink = bidCases("td:nth-of-type(3) a").attr("href")
                bidList_ins.insert(3,bidCases("td:nth-of-type(3) u").text())
                bidList_ins[2] = bidList_ins[2].split("\n\n")[0]
                bidList_ins.pop(0)
                bidList_ins.pop(4)
                bidList_ins.pop(4)
                bidList_ins.append(bidLink)
                a = bidList_ins[:]
                a.insert(0,s)
                if bidList_1.count(bidList_ins) == 0:
                    bidList_1.append(bidList_ins)
                    bidList.append(a)
    for m in orgKW:
        bidSystemUrl = "https://web.pcc.gov.tw/tps/pss/tender.do?method=goSearch&searchMode=common&searchType=basic"
        browse = driver.get(bidSystemUrl)
        homepage = driver.page_source
        query = driver.find_element_by_name("orgName")
        query.clear()
        query.send_keys(m)
#         queryCate = driver.find_element_by_css_selector("#radProctrgCate3")
#         queryCate.click()
        query.send_keys(Keys.RETURN)
        sleep(1) #等頁面回傳
        bidpage = driver.page_source
        querypage = pq(bidpage)
        querypage.make_links_absolute(base_url=bidSystemUrl)
        #tbody下的所有onmouseover='overcss(this);'的物件，屬性選擇器用法很複雜要小心
        #如果關鍵字搜尋不到就不會有此物件，會出現onmouseover="overcss(this)"的物件，少一個分號;，物件不一樣
        checkQuery = querypage("[onmouseover='overcss(this);']").text()  
        #print('===關鍵字：',s,"\n")
        if checkQuery:
            for bidCases in querypage("tbody [onmouseover='overcss(this);']").items():
                bidList_ins = []
                for eachBid in bidCases("td").items():
                    bidList_ins.append(eachBid.text())
                bidLink = bidCases("td:nth-of-type(3) a").attr("href")
                bidList_ins.insert(3,bidCases("td:nth-of-type(3) u").text())
                bidList_ins[2] = bidList_ins[2].split("\n\n")[0]
                bidList_ins.pop(0)
                bidList_ins.pop(4)
                bidList_ins.pop(4)
                bidList_ins.append(bidLink)
                a = bidList_ins[:]
                if m == "水利署": a.insert(0,'WRA')
                elif m == "大地工程處": a.insert(0,'GEO')
                else: a.insert(0,'NCDR')
                if bidList_1.count(bidList_ins) == 0:
                    bidList_1.append(bidList_ins)
                    bidList.append(a)
    driver.quit()
    return(bidList)

mailtext = ""
bid_list = []

def crawl_job():
    keywords= ['聯網','監測','自動','智慧','監控','灌溉','節水','調控','逕流','出流管制','塔寮坑']
    orgKW= ['水利署']
    bid1 = crawlBid(keywords, orgKW)
    text1 = '<b>水專管今日標案　　</b><a href="http://140.112.10.220/allbidWRA.html">等標期內標案 按此<a><br>'
    text1 += mailBody(bidHead, bid1)
    text1 += '標案關鍵字：'+str(keywords)+'<br>機關關鍵字：'+str(orgKW)
    keywords= ['智慧','溪溝','水土保持','坡地','物聯','監測','可利用','土地利用','邊坡','社區','水資源','崩塌','土石流','防災','太陽能','大數據','綠能','土砂','雲端','韌性','巨量資料','地理','液化','維護管理','監控','電信','系統整合','汚水','三維','管線','BIM','灌溉','資訊','管理','下水道','農村社區','森林','火災','養殖','漏水']
    orgKW= ['國家災害防救科技中心','大地工程處']
    bid2 = crawlBid(keywords, orgKW)
    text1 += '<br><br><b>團隊相關今日標案　　</b><a href="http://140.112.10.220/allbid.html">等標期內標案 按此<a><br>'
    text1 += mailBody(bidHead, bid2)
    text1 += '標案關鍵字：'+str(keywords)+'<br>機關關鍵字：'+str(orgKW)
    # if bid1 != []: sendMail(mailTo,text1)
    print("Today is DONE!")
    global mailtext
    global bid_list
    mailtext = text1
    bid_list = bid1


def mail_job():
    if bid_list != []: sendMail(mailTo,mailtext)

scheduler.add_job(crawl_job, 'cron',day_of_week='0-6', hour='07',minute='55',second='00')
scheduler.add_job(mail_job, 'cron',day_of_week='0-6', hour='08',minute='00',second='00')
scheduler.add_job(crawl_job, 'cron',day_of_week='0-6', hour='17',minute='25',second='00')
scheduler.add_job(mail_job, 'cron',day_of_week='0-6', hour='17',minute='30',second='00')
scheduler.start()
