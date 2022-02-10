import numpy as np
import pandas as pd

import requests
from bs4 import BeautifulSoup

import json

from datetime import datetime, timedelta
import os

def get_proj_list_by_date(date):
    unit_keywords = ['宜蘭', '羅東', '第四區養護工程處', '水土保持局臺北分局']
    date_url = f"https://pcc.g0v.ronny.tw/api/listbydate?date={date}"
    # print(date_url)
    r = requests.get(date_url)
    d = json.loads(r.text)
    if not d:
        return []
    p = [i for i in d['records'] if (i['brief']['type'] == '公開招標公告') and ('工程類' in i['brief']['category'])]
    
    projs = []
    for k in unit_keywords:
        for i in p:
            try:
                if k in i['unit_name']:
                    projs.append(i)
            except:
                continue
    return projs

def get_proj_info(unit_id, job_number):
    proj_url = f"https://pcc.g0v.ronny.tw/api/tender?unit_id={unit_id}&job_number={job_number}"
    # print(proj_url)
    r = requests.get(proj_url)
    d = json.loads(r.text)
    return {
        "unit_name": d['records'][-1]['unit_name'], 
        "proj_name": d['records'][-1]['brief']['title'], 
        "proj_type": d['records'][-1]['brief']['type'], 
        "proj_category": d['records'][-1]['brief']['category'], 
        "proj_budget": d['records'][-1]['detail']['採購資料:預算金額'], 
        "proj_loc": d['records'][-1]['detail']['其他:履約地點'], 
        "proj_pub_time":d['records'][-1]['detail']['招標資料:公告日'], 
        "proj_end_time":d['records'][-1]['detail']['領投開標:截止投標'], 
        "proj_open_time":d['records'][-1]['detail']['領投開標:開標時間'], 
        "url": d['records'][-1]['detail']['url']
        }

if __name__ == '__main__':
    # Monday is 0, Sunday is 6
    weekday = datetime.today().weekday()
    day_list = [i.strftime('%Y%m%d') for i in pd.date_range(datetime.today()-timedelta(days=7+weekday), datetime.today()-timedelta(days=2+weekday))][:-1]

    proj_df = pd.DataFrame()
    for date in day_list:
        print(date)
        projs = get_proj_list_by_date(date)
        for proj in projs:
            try:
                temp_df = pd.DataFrame()
                proj_info = get_proj_info(proj['unit_id'], proj['job_number'])
                for k, v in proj_info.items():
                    temp_df[k] = [v]
                proj_df = pd.concat([proj_df, temp_df], ignore_index=True)
            except Exception as e:
                print(proj['unit_id'], proj['job_number'], " : ", e)
                continue
    proj_df.drop_duplicates(keep='last', ignore_index=True, inplace=True)
    proj_df.rename(columns={"unit_name":"機關名稱", "proj_name":"標案名稱", "proj_type":"招標方式", "proj_category":"標的分類", 
                            "proj_budget":"預算金額", "proj_loc":"履約地點", "proj_pub_time":"公告日", "proj_end_time":"截止投標", 
                            "proj_open_time":"開標時間"}, inplace=True)

    saved_name = os.path.join("z:/", "00_標案資料", f"{datetime.today().strftime('%Y%m%d')}.xlsx")
    proj_df.to_excel(saved_name, index=None)
