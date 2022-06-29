import json
import pdfkit
import os
import subprocess
from threading import Thread
import calendar
import threading
from requests import delete

# import requests

"""

輸入使用
text = input('')
li = ['2022-01', '2020-02', '2020-03', '2020-04', '2020-05', '2020-06', '2020-07', '2020-08', '2020-09', '2020-10']
[i for i in li if '2022-01' == text][0]

每日批次API結構
http://210.61.217.104/Report6BatchAll/237/2022-05-01/2022-05-31/237/248
每周
http://210.61.217.104/Report6BatchAll/214/2022-05-01/2022-05-31/214/225

分3條線程
    消防
    電力
    水力

data資料結構
    消防
    電力
        每日
        每周
        每月
    水力
        每日
        每周
        每月
寫好電 水就copy就好

call法，在cmd上 顯示文字 下載消防報表 選擇A，電力報表選擇B、水力報表選擇C，全部選擇all

限制性輸入取當前年分01~12月如 2020-01

水：每日、每周、月份，各一個funtion
電：每日、每周、月份，各一個funtion

輸出資料夾結構
6月消防報表
6月電力報表
6月水力報表

目前先想功能就好，數據慢慢手動建立
打包後，寫一個下介面選單，12個月分並選擇後判斷當前是否是當月或是低於月份，超過則跳錯誤
輸出完成跳出資料夾
將data.json放到github web上，再用request讀取，方便新增資料
新增功能
    當前路徑建立資料夾，判斷資料夾是否存在，以月分建檔名，5月消防月報表
colab 上處理，全部轉換後並一次上傳到drive，使用rclone工具
高階一點的寫法
    5個線程處理
wkhtmltopdf doc
 https://wkhtmltopdf.org/usage/wkhtmltopdf.txt
套件 
 https://pypi.org/project/pdfkit/
不更改環境變數
import pdfkit
path_wkhtmltopdf = r'C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe' 取當前路徑並接合 os.path.join("root", "directory1", "directory2")
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
pdfkit.from_url("http://google.com", "out.pdf", configuration=config)
"""

class htmltopdf:
    def __init__(self, File_folder='水電消防報表',  fire='消防', electricity_folder='電力', drain='排水'):
        self.File_folder = File_folder # Share
        self.fire_folder = fire
        self.electricity_folder = electricity_folder
        self.drain_folder = drain

    # wkhtmltopdf Path setting
    path_wkhtmltopdf = os.path.join(os.getcwd(), 'wkhtmltox\\bin', 'wkhtmltopdf.exe')
    config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

    # PDF output settings
    options = {
    'no-background': None
    }

    # open all data
    # url json https://www.delftstack.com/zh-tw/howto/python/python-get-json-from-url/
    with open("test_data.json", encoding="utf-8") as f: #Test usefile testdata.json
        open_data = json.load(f) # json data

    def folder(self, folderpath_name):
    # 使用相對路徑且資料夾都在根目錄
        folderpath = folderpath_name 
        # 檢查目錄是否存在 
        if os.path.isdir(folderpath):
            print('{} 資料夾存在。'.format(folderpath))
        else:
            print('資料夾不存在。建立{}資料夾'.format(folderpath))
            os.mkdir(folderpath)
            print('{}建立完成'.format(folderpath))
    
    #消防設備
    def Fire_call(self, date:str):
        date = date  # ex : '2022-05'

        #建立資料夾
        self.folder(os.path.join(self.File_folder))  #水電消防報表
        self.folder(os.path.join(self.File_folder,self.fire_folder))  #消防
        self.folder(os.path.join(self.File_folder,self.fire_folder,date))  # 月

        # with open("data.json", encoding="utf-8") as f: #Test usefile testdata.json
        #     p = json.load(f) # json data

    
        # 寫一個匿名funtion 計算多少筆檔案，抓name(使用len)
        count_file = []
        for i in self.open_data['Fire_Equipment']:
            count_file.append(i['name'])
        print(f'共有{len(count_file)}個PDF')

        try:
            for i in self.open_data['Fire_Equipment']:
                name = i['name']
                url_api_1 = i['api_1']
                url_api_2 = i['api_2']
                print(f'http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}' + f'  {name}')
                url = f'http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}'
                pdfkit.from_url(url, os.path.join(self.File_folder, self.fire_folder, date, name) + '.pdf', options=htmltopdf.options, configuration=htmltopdf.config)
        except:
            print('請求失敗', url)

        # open folder
        start_directory = os.path.join(self.File_folder, self.fire_folder, date)
        self.startfile(start_directory)

    # 電力設備
    def electricity(self, date:str):
        """
        api 結構
        /Report6BatchAll/237/2022-05-01/2022-05-31/237/248
        /Report6BatchAll/237/{2022-05}-01/{2022-05}-31/237/248

        /Report6BatchAll/237/{2022-05}-01/{2022-05}-{getmotn}/237/248

        使用批次檔案產生
        """

        self.folder(os.path.join("test", date))  #建立資料夾

        date = date # ex : '2022-05'

        try:
            for i in self.open_data['electricity_every_day']:
                name = i['name']
                url_api_1 = i['api_1']
                url_api_2 = i['api_2']
                print(f'http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}' + f'  {name}')
                url = f'http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
                pdfkit.from_url(url, os.path.join("test", date, name) + '.pdf', options=htmltopdf.options, configuration=htmltopdf.config)
        except:
            print('請求失敗', url)

        # open folder
        start_directory = os.path.join("test", date) 
        self.startfile(start_directory)

    # Finish Open the folder
    def startfile(sele, filename):
        try:
            os.startfile(filename)
        except:
            subprocess.Popen(['xdg-open', filename])

    # get month of day   
    def get_monthrange(sele,date):
        month = sele.delete_left_zero(date)
        year = date.split('-')[0]
        day = calendar.monthrange(int(year),int(month))
        print(f'{date} 月份天數 {day[1]}')
        return day[1]
        

    # Remove the left zero ex:'08' to '8' ， Use in get month of day
    def delete_left_zero(sele,date:str):
        str = date
        del_zero = str.split('-')[1].lstrip('0')
        # print(del_zero)
        return del_zero
    
    def input_(self):
        while True:
            text=input('輸入格式，範例 2022-06: ')
            if len(text)==7 and int(text[:4]) >= 2022 and text[4] == '-' and text[:4].isnumeric() \
                and len(text[:4]) == 4 and text[5:7].isnumeric()\
                and len(text[5:7]) == 2 and int(text[5:7]) >= 1 and int(text[5:7]) <= 12:
                break
            else:
                print('輸入錯誤，請輸入正確格式')
        print('Valid input')
        return text


    # thread
    def start_thread(self, thread_number, call, thread_name:str):
        """
        重寫
        """
        

        try:
            threading.Thread(target=call, args=(thread_number,))
            Thread()
        except:
            print(f'失敗第{thread_name}線程') 

if __name__ == '__main__':
    my = htmltopdf()
    my.Fire_call('2022-05')
    # my.Fire_call('2022-03')
    # my.get_monthrange('2022-06')
    # my.electricity('2022-05')
    
    # date = input('')

    # my.delete_left_zero('2022-06')
    # pdfkit.from_url('http://210.61.217.104/Report6/82/2022-05/82/98', os.path.join("test", 'abc11') + '.pdf', options=options, configuration=config)


## 測試資料正確性
    # with open("data.json", encoding="utf-8") as f:
    #     p = json.load(f) # json data

    # for i in p['Fire_Equipment']:
    #     name = i['name']
    #     url_api_1 = i['api_1']
    #     url_api_2 = i['api_2']

    #     if url_api_1 == 'input' or url_api_2 == 'input':
    #         continue
    #     else:
    #         print(url_api_1)

        # print(url_api_2)

"""
異部
open file一個區塊
pdfkit一個區塊，達成條件就call
計算每跑20筆資料一個線程
"""