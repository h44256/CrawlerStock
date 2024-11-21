import datetime
import requests 
from bs4 import BeautifulSoup
import pandas as pd
import time
from datetime import datetime as dtime
from dateutil.relativedelta import relativedelta

def getStockInfo():
    urlAjax = "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY_AVG?date=20241119&stockNo=2330&response=json&_=1732023234816"
    headers = {'User-Agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36'}
    response = requests.get(urlAjax, headers = headers)

    if response.status_code==200:
        data = response.json()
        stock_list = []
        print("正在爬取......");
        #title = data["title"]
        datas = data["data"]
        for d in datas:
            date = d[0]
            price = d[1]
            data_info=[date,price]
            stock_list.append(data_info)
        df = pd.DataFrame(stock_list, columns=["日期","收盤價"])
        df.to_excel('stock.xlsx', index=False, engine="openpyxl")
        print("完成資料蒐集")

def calculate_months(date_start, date_end):
    start = dtime.strptime(date_start, "%Y%m%d")
    end = dtime.strptime(date_end, "%Y%m%d")
    # 計算年份差和月份差
    year_diff = end.year - start.year
    month_diff = end.month - start.month
    # 總月份數
    total_months = year_diff * 12 + month_diff
    return total_months
def input_date(prompt):
    while True:
        date_str = input(prompt)
        # 先檢查長度是否為 8
        if len(date_str) != 8:
            print("日期長度錯誤！請按照 YYYYMMDD 格式輸入，例如：20240701")
            continue
        try:
            # 嘗試將輸入轉換為日期
            date = dtime.strptime(date_str, "%Y%m%d")
            return date_str  # 如果格式正確，返回字串
        except ValueError:
            # 如果格式錯誤，提示重新輸入
            print("日期格式錯誤！請按照 YYYYMMDD 格式輸入，例如：20240701")


stockName = input("請輸入股票代號或名稱")
dateStart = input_date("請輸入想查詢的\"起始\"年月日 (格式: YYYYMMDD): ")
dateEnd = input_date("請輸入想查詢的\"結束\"年月日 (格式: YYYYMMDD): ")
url = "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY_AVG"
headers = {'User-Agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36'}
#設定日期跟總月份
months = calculate_months(dateStart, dateEnd)
current_date = dtime.strptime(dateStart, "%Y%m%d")
stock_list = []#紀錄日期跟收盤價資料
for i in range(months+1):
    print("正在爬取......");
    time.sleep(3)#遵守證交所的頻率, 不要對server造成衝擊, 不然要被鎖IP
    dateName = current_date.strftime("%Y%m%d")#帶入參數日期
    params = {"response":"json",
                "date": f"{dateName}",
                "stockNo": f"{stockName}"}
    response = requests.get(url, headers=headers, params=params)
    if response.status_code==200:
        data = response.json()
        #title = data["title"]
        datas = data["data"]#從撈到的資料裡面取得data(證交所剛好設定這資料名叫做data)
        for d in datas:
            date = d[0] #當天
            price = d[1] #當天價錢
            data_info=[date,price]
            stock_list.append(data_info)#放到list裡面
        print(f"完成 {current_date.strftime('%Y-%m')} 資料蒐集")
    current_date += relativedelta(months=1)
df = pd.DataFrame(stock_list, columns=["日期","收盤價"])
df.to_excel(f'{stockName}.xlsx', index=False, engine="openpyxl")
print("完成所有資料蒐集")