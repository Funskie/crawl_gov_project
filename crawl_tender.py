import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup

import time
from time import sleep
from datetime import datetime
import re

import numpy as np
import pandas as pd

import smtplib
from email.mime.text import MIMEText
from email import encoders, utils
from email.mime.multipart import MIMEMultipart

keywords = ["邊坡", "崩塌", "復建"]
orgkws = ["水土保持局臺北分局", "羅東林區管理處", "第四區養護工程處", "宜蘭縣"]
loc_list = ["臺北", "新北", "桃園", "新竹", "宜蘭"]
# phantomjs_path = r"D:\PyProjects\projcrawler\phantomjs-2.1.1-windows\bin\phantomjs.exe"
mailList = [
    'chongjing3370@gmail.com', 
    'tusty9292@gmail.com', 
    'fish892555@gmail.com'
    ]

def get_tender_by_kw(kws, loc):
    driver = webdriver.Chrome()
    tender_list = []
    for k in kws:
        bidSystemUrl = "https://web.pcc.gov.tw/pis/"
        browse = driver.get(bidSystemUrl)
        homepage = driver.page_source
        query = driver.find_element_by_name("tenderName")
        query.clear()
        query.send_keys(k)
        queryCate = driver.find_element_by_css_selector("#basicRadProctrgCate1")
        queryCate.click()
        queryTime = driver.find_element_by_css_selector("#basicIsSpdtDateTypeId")
        queryTime.click()
        query_loc = driver.find_element_by_name("orgName")
        query_loc.send_keys(loc)
        driver.execute_script('return basicTenderSearch()')
        # query.send_keys(Keys.RETURN)
        sleep(1) #等頁面回傳
        bidpage = driver.page_source
        soup = BeautifulSoup(bidpage, 'html.parser')
        tender_table = soup.find_all("tbody")[-1].find_all("tr")
        if (len(tender_table) == 1) and (tender_table[0].text == "無符合條件資料"):
            continue
        else:
            tender_list += tender_table
    driver.close()
    return list(set(tender_list))

def get_tender_by_org(orgs):
    driver = webdriver.Chrome()
    tender_list = []
    for o in orgs:
        bidSystemUrl = "https://web.pcc.gov.tw/pis/"
        browse = driver.get(bidSystemUrl)
        homepage = driver.page_source
        query = driver.find_element_by_name("orgName")
        query.clear()
        query.send_keys(o)
        queryCate = driver.find_element_by_css_selector("#basicRadProctrgCate1")
        queryCate.click()
        queryTime = driver.find_element_by_css_selector("#basicIsSpdtDateTypeId")
        queryTime.click()
        driver.execute_script('return basicTenderSearch()')
        # query.send_keys(Keys.RETURN)
        sleep(1) #等頁面回傳
        bidpage = driver.page_source
        soup = BeautifulSoup(bidpage, 'html.parser')
        tender_table = soup.find_all("tbody")[-1].find_all("tr")
        if (len(tender_table) == 1) and (tender_table[0].text == "無符合條件資料"):
            continue
        else:
            tender_list += tender_table
    driver.close()
    return list(set(tender_list))

def get_tender_info(tender):
    return {
        "機關名稱": "".join(tender.find_all("td")[1].text.split()), 
        "標案案號&標案名稱": " ".join(tender.find_all("td")[2].text.split()), 
        "傳輸次數": "".join(tender.find_all("td")[3].text.split()), 
        "招標方式": "".join(tender.find_all("td")[4].text.split()), 
        "採購性質": "".join(tender.find_all("td")[5].text.split()), 
        "公告日期": "".join(tender.find_all("td")[6].text.split()), 
        "截止投標": "".join(tender.find_all("td")[7].text.split()), 
        "預算金額": "".join(tender.find_all("td")[8].text.split()), 
        "url": tender.find_all("td")[2].find('a').get('href').replace("/prkms/urlSelector/common/tpam?pk", "https://web.pcc.gov.tw/tps/QueryTender/query/searchTenderDetail?pkPmsMain"),
           }

if __name__ == '__main__':
    t_ks = []
    for loc in loc_list:
        t = get_tender_by_kw(keywords, loc)
        t_ks += t
    t_orgs = get_tender_by_org(orgkws)
    t_list = list(set(t_ks+t_orgs))

    tender_df = pd.DataFrame()
    for t in t_list:
        temp_df = pd.DataFrame()
        t_info = get_tender_info(t)
        for k, v in t_info.items():
            temp_df[k] = [v]
        tender_df = pd.concat([tender_df, temp_df], ignore_index=True)
    tender_df.sort_values("截止投標", inplace=True, ignore_index=True)
    tender_df.drop_duplicates(keep='last', ignore_index=True, inplace=True)
    tender_df['url'] = '<a href=' + tender_df['url'] + '><div>連結</div></a>'

    now = datetime.now().strftime("%Y/%m/%d")
    text = f"<h>{now}等標期內標案</h>"
    text += tender_df.to_html(escape=False).replace("\n", "")

    me = 'tusty9292@gmail.com'
    mail = MIMEMultipart()
    mail['Subject'] = "政府採購公告網 "+str(time.strftime("%Y/%m/%d-%H%M"))
    mail['From'] = 'chongjing3370@gmail.com'
    mail['Date'] = utils.formatdate(localtime = 1)
    mail['Message-ID'] = utils.make_msgid()
    mail['Content-Type'] = "text/calendar; charset=utf-8"
    body = MIMEText(text, 'html', "utf-8")
    body.set_charset("utf-8")
    mail.attach(body)

    with smtplib.SMTP(host="smtp.gmail.com", port="587") as smtp:  # 設定SMTP伺服器
        try:
            smtp.ehlo()  # 驗證SMTP伺服器
            smtp.starttls()  # 建立加密傳輸
            smtp.login('chongjing3370@gmail.com', 'njhxfuhvwdmdjgds')  # 登入寄件者gmail
            smtp.sendmail(me, mailList, mail.as_string())  # 寄送郵件
            print("Complete!")
        except Exception as e:
            print("Error message: ", e)
