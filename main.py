from distutils import ccompiler
import json
from turtle import back
# from typing_extensions import Self
import pdfkit
import os
import subprocess
from threading import Thread
import threading
import calendar
import _thread
from concurrent.futures import ThreadPoolExecutor
import datetime
import time
import colorama
from colorama import Fore
from colorama import Style
import re
import sys
import PyPDF2

"""
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
    排水
        每日
        每周
        每月

wkhtmltopdf doc
 https://wkhtmltopdf.org/usage/wkhtmltopdf.txt
套件 
 https://pypi.org/project/pdfkit/
"""


class htmltopdf():
    def __init__(self, File_folder='水電消防報表',  fire='消防', electricity_folder='電力', drain='排水'):
        self.File_folder = File_folder # Share
        self.fire_folder = fire
        self.electricity_folder = electricity_folder
        self.drain_folder = drain

    colorama.init(autoreset=True)
    # wkhtmltopdf Path setting
    # path_wkhtmltopdf = os.path.join(os.getcwd(), 'wkhtmltox\\bin', 'wkhtmltopdf.exe')
    path_wkhtmltopdf = os.path.join('wkhtmltox\\bin','wkhtmltopdf.exe')
    config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

    # PDF output settings
    options = {
    'no-background': None
    }
    
    # open all data
    # url json https://www.delftstack.com/zh-tw/howto/python/python-get-json-from-url/
    
    data_file = 'data.json' # #Test usefile testdata.json

    with open(data_file , encoding="utf-8") as f:
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
        """
        消防設備
        輸入值 YYYY-DD
        """

        date = date  # ex : '2022-05'

        #建立資料夾
        self.folder(os.path.join(self.File_folder,self.fire_folder))  #消防
        self.folder(os.path.join(self.File_folder,self.fire_folder,date))  # 月

        # with open("data.json", encoding="utf-8") as f: #Test usefile testdata.json
        #     p = json.load(f) # json data

    
        # count file
        count_file = []
        for i in self.open_data['Fire_Equipment']:
            count_file.append(i['name'])
        print(f'{self.fire_folder} 共有 {len(count_file)} 個PDF\n')

        try:
            for i in self.open_data['Fire_Equipment']:
                name = i['name']
                url_api_1 = i['api_1']
                url_api_2 = i['api_2']
                print(f'{self.get_now_date()}  http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}' + f'  {Fore.RED}{Style.BRIGHT}{name}{Style.RESET_ALL}\n')
                url = f'http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}'
                pdfkit.from_url(url, os.path.join(self.File_folder, self.fire_folder, date, name) + '.pdf', options=htmltopdf.options, configuration=htmltopdf.config)
        except:
                print('請求失敗', url)

        # open folder
        start_directory = os.path.join(self.File_folder, self.fire_folder, date)
        self.startfile(start_directory)
    
    # 消防 指定搜尋單一類別
    def Fire_call_find(self, date:str, input:str):
        pass
        """
        消防設備
        輸入值 YYYY-DD
        """

        date = date  # ex : '2022-05'
        find = []
        #建立資料夾
        self.folder(os.path.join(self.File_folder,self.fire_folder))  #消防
        self.folder(os.path.join(self.File_folder,self.fire_folder,date))  # 月

        # with open("data.json", encoding="utf-8") as f: #Test usefile testdata.json
        #     p = json.load(f) # json data

        # 指定搜尋清單，例如：長青樓，正規 模糊搜尋要的清單 2/26 2023，已完成功能
        
        try:
            for i in self.open_data['Fire_Equipment']:
                if re.search(input ,i['name']):
                    find.append(i['name'])
                    
                
            # 使用清單 list 作為判斷，如果有資料 >0 就會執行統計
            if len(find) > 0:
                # 列出找到檔案數量
                print(f'{self.fire_folder} 共有 {len(find)} 個PDF\n')        

                    # 列出找到清單
                print(f'將下載以下報表 \n')
                for a in find:
                    print(f'{Fore.RED}{Style.BRIGHT}{a}{Style.RESET_ALL} \n')
            else:
                print(f'{Fore.RED}{Style.BRIGHT}找不到你輸入的：{input}{Style.RESET_ALL}')

        except:
            print(f'找不到你輸入的：{input}')





        # count file
        # count_file = []
        # for i in self.open_data['Fire_Equipment']:
        #     count_file.append(i['name'])

        # 使用清單 list 作為判斷，如果有資料 >0 就會執行爬蟲
        if len(find) > 0:
            try:
                for i in self.open_data['Fire_Equipment']:
                    if re.search(input ,i['name']):
        
                        print(f"下載報表： {i['name']} \n")

                        name = i['name']
                        url_api_1 = i['api_1']
                        url_api_2 = i['api_2']
                        print(f'{self.get_now_date()}  http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}' + f'  {Fore.RED}{Style.BRIGHT}{name}{Style.RESET_ALL}\n')
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
        電力
        輸入值 YYYY-DD

        api 結構
        /Report6BatchAll/237/2022-05-01/2022-05-31/237/248
        /Report6BatchAll/237/{2022-05}-01/{2022-05}-31/237/248

        /Report6BatchAll/237/{2022-05}-01/{2022-05}-{getmotn}/237/248

        使用批次檔案產生
        """

        #建立資料夾
        # self.folder(os.path.join(self.File_folder))  #水電消防報表
        self.folder(os.path.join(self.File_folder,self.electricity_folder))  #電力
        self.folder(os.path.join(self.File_folder,self.electricity_folder,date))  # 月



        # self.folder(os.path.join("test", date))  #建立資料夾

        date = date # ex : '2022-05'


        # file count
        self.count_file('electricity_every')

        # day
        try:
            for i in self.open_data['electricity_every_day']:
                name = i['name']
                url_api_1 = i['api_1']
                url_api_2 = i['api_2']
                print(f'{self.get_now_date()} http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}' + f'  {Fore.BLUE}{Style.BRIGHT}{name}{Style.RESET_ALL}\n')
                url = f'http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
                pdfkit.from_url(url, os.path.join(self.File_folder, self.electricity_folder, date, name) + '.pdf', options=htmltopdf.options, configuration=htmltopdf.config)
        except:
                print('請求失敗', url)


        #week
        try:
            for i in self.open_data['electricity_every_week']:
                name = i['name']
                url_api_1 = i['api_1']
                url_api_2 = i['api_2']
                print(f'{self.get_now_date()} http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}' + f'  {Fore.BLUE}{Style.BRIGHT}{name}{Style.RESET_ALL}\n')
                url = f'http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
                pdfkit.from_url(url, os.path.join(self.File_folder, self.electricity_folder, date, name) + '.pdf', options=htmltopdf.options, configuration=htmltopdf.config)
        except:
                print('請求失敗', url)


        #month
        try:
            for i in self.open_data['electricity_every_month']:
                if i['name'] == '中正樓24F停機坪照明設備巡檢紀錄':
                    name = i['name']
                    url_api_1 = i['api_1']
                    url_api_2 = i['api_2']
                    print(f'{self.get_now_date()} http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}' + f'  {Fore.BLUE}{Style.BRIGHT}{name}{Style.RESET_ALL}\n')
                    url = f'http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}'
                    pdfkit.from_url(url, os.path.join(self.File_folder, self.electricity_folder, date, name) + '.pdf', options=htmltopdf.options, configuration=htmltopdf.config)
                    continue

                name = i['name']
                url_api_1 = i['api_1']
                url_api_2 = i['api_2']
                print(f'{self.get_now_date()} http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}' + f'  {Fore.BLUE}{Style.BRIGHT}{name}{Style.RESET_ALL}\n')
                url = f'http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
                pdfkit.from_url(url, os.path.join(self.File_folder, self.electricity_folder, date, name) + '.pdf', options=htmltopdf.options, configuration=htmltopdf.config)
        except:
                print('請求失敗', url)


        # open folder
        start_directory = os.path.join(self.File_folder,self.electricity_folder,date) 
        self.startfile(start_directory)

    # 給排水設備
    def drain(self, date:str):
        """
        給排水設備
        輸入值 YYYY-DD

        api 結構
        /Report6BatchAll/237/2022-05-01/2022-05-31/237/248
        /Report6BatchAll/237/{2022-05}-01/{2022-05}-31/237/248
        /Report6BatchAll/237/{2022-05}-01/{2022-05}-{getmotn}/237/248
        使用批次檔案產生
        """

        #建立資料夾
        # self.folder(os.path.join(self.File_folder))  #水電消防報表
        self.folder(os.path.join(self.File_folder,self.drain_folder))  #給排水
        self.folder(os.path.join(self.File_folder,self.drain_folder,date))  # 月



        # self.folder(os.path.join("test", date))  #建立資料夾

        date = date # ex : '2022-05'
        
        # file count
        self.count_file('drain')

        # start_time = time.time() # START
        
        # day
        try:
            for i in self.open_data['drain_day']:
                name = i['name']
                url_api_1 = i['api_1']
                url_api_2 = i['api_2']
                print(f'{self.get_now_date()} http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}' + f'  {Fore.GREEN}{Style.BRIGHT}{name}{Style.RESET_ALL}\n')
                url = f'http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
                pdfkit.from_url(url, os.path.join(self.File_folder, self.drain_folder, date, name) + '.pdf', options=htmltopdf.options, configuration=htmltopdf.config)
        except:
                print('請求失敗', url)

        #week
        try:
            for i in self.open_data['drain_week']:
                name = i['name']
                url_api_1 = i['api_1']
                url_api_2 = i['api_2']
                print(f'{self.get_now_date()} http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}' + f'  {Fore.GREEN}{Style.BRIGHT}{name}{Style.RESET_ALL}\n')
                url = f'http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
                pdfkit.from_url(url, os.path.join(self.File_folder, self.drain_folder, date, name) + '.pdf', options=htmltopdf.options, configuration=htmltopdf.config)
        except:
                print('請求失敗', url)

        #month
        try:
            for i in self.open_data['drain_month']:
                name = i['name']
                url_api_1 = i['api_1']
                url_api_2 = i['api_2']
                print(f'{self.get_now_date()} http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}' + f'  {Fore.GREEN}{Style.BRIGHT}{name}{Style.RESET_ALL}\n')
                url = f'http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
                pdfkit.from_url(url, os.path.join(self.File_folder, self.drain_folder, date, name) + '.pdf', options=htmltopdf.options, configuration=htmltopdf.config)
        except:
                print('請求失敗', url)

        # open folder
        start_directory = os.path.join(self.File_folder,self.drain_folder,date) 
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
        # print(f'{date} 月份天數 {day[1]}')
        return day[1]
        

    # Remove the left zero ex:'08' to '8' ， Use in get month of day
    def delete_left_zero(sele,date:str):
        str = date
        del_zero = str.split('-')[1].lstrip('0')
        # print(del_zero)
        return del_zero
    
    def input_str(self):
        """
        call使用，限制格式，YYYY-MM
        """
        
        while True:
            text=input(f'輸入格式，範例： {self.get_date()} ：')
            if len(text)==7 and int(text[:4]) >= 2022 and text[4] == '-' and text[:4].isnumeric() \
                and len(text[:4]) == 4 and text[5:7].isnumeric()\
                and len(text[5:7]) == 2 and int(text[5:7]) >= 1 and int(text[5:7]) <= 12:
                break
            else:
                print('輸入錯誤，請輸入正確格式')
        print('Valid input')
        return text

    #Current date
    def get_date(self):
        """
        取得現在年、月、日，以電腦時間為準
        """
        if datetime.datetime.now().month == 1:
            get_year = datetime.datetime.now().year -1
        else:
            get_year = datetime.datetime.now().year

        get_month = datetime.datetime.now().month - 1 if datetime.datetime.now().month > 1 else 12 #下載上個月報表，減一個月
        get_day = datetime.datetime.now().day
        dt = datetime.datetime(get_year, get_month, get_day)
        # print(dt.strftime('%Y-%m-%d'))
        return dt.strftime('%Y-%m')
    
    
    # 資料API測試正確性
    def Data_verification(self):
        """
        消防 
            ['Fire_Equipment']
        電力 
            ['electricity_every_week']
            ['electricity_every_month']
            ['electricity_every_day']
        排水

        """

        for i in self.open_data['electricity_every_week']:
            name = i['name']
            url_api_1 = i['api_1']
            url_api_2 = i['api_2']

            if url_api_1 == 'input' or url_api_2 == 'input':
                continue
            else:
                print(url_api_2)
        
    def count_file(self,type_name:str):
        """
        電力、排水計算檔案統計使用
        type_name: type
            電力 electricity_every
            排水 drain
        """

        # for day,week,month in zip(self.open_data['drain_day'],self.open_data['drain_month'],self.open_data['drain_week']):
        #     print(day['name'])
        #     print(week['name'])
        #     print(month['name'])
        count_file = []
        month  = [i['name'] for i in self.open_data[f'{type_name}_month']]
        day  = [i['name'] for i in self.open_data[f'{type_name}_day']]
        week  = [i['name'] for i in self.open_data[f'{type_name}_week']]

        count_file += month + day + week

        name = {'electricity_every':'電力','drain':'排水',}

        print(f'{name[type_name]} 共有 {len(count_file)} 個PDF')

    def General_folder(self):
        """
        判斷建立共用資料夾 "水電消防報表"
        """
        self.folder(os.path.join(self.File_folder))  #水電消防報表

    # 取得現在日期
    def get_now_date(self):
        now = datetime.datetime.now()
        current_time = now.strftime(f"{Fore.BLUE}{Style.BRIGHT}%H:%M:%S{Style.RESET_ALL}")
        return current_time
    
    # 電子巡檢內建每月報表的PDF，搜尋特定標題並合併一個PDF檔案
    def pdf_report_merge(self, target_title:str):
        """
        每月報表產出 搜尋特定標題並合併一個PDF

        儲存在 報表合併pdf 資料夾
        """
        try:
            # 要搜尋的特定標題
            target_title = target_title  # 例如 '中正樓'

            # 指定PDF文件的路徑
            pdf_path_input = input(r'輸入pdf路徑：')


            if pdf_path_input[-3:] == 'pdf':
                pdf_path = pdf_path_input
            else:
                print('路徑輸入錯誤，檔案必須pdf')


            # 指定儲存檔案的資料夾路徑，資料夾名稱為當前時間的年月日
            now = datetime.datetime.now()
            folder_name = now.strftime("%Y%m%d")

            output_folder = os.path.join('報表合併pdf', folder_name)

            # 如果指定的資料夾不存在，則創建資料夾
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            # 讀取PDF文件，並搜索特定標題
            pdf_file = open(pdf_path, "rb")
            pdf_reader = PyPDF2.PdfFileReader(pdf_file)

            merged_pdf_writer = PyPDF2.PdfFileWriter()
            found = False
            print('執行中，請等待結果')
            for page_num in range(pdf_reader.numPages):
                page = pdf_reader.getPage(page_num)
                page_title = page.extractText().strip().split('\n')[0]
                if re.search(target_title, page_title, re.IGNORECASE):  # 使用re模組進行模糊搜尋，並忽略大小寫
                    merged_pdf_writer.addPage(page)
                    found = True

            # 如果搜尋到特定標題，則將合併後的PDF文件儲存為以特定標題為檔名的PDF文件
            if found:
                output_file_name = target_title + ".pdf"
                output_file_path = os.path.join(output_folder, output_file_name)
                output_file = open(output_file_path, "wb")
                merged_pdf_writer.write(output_file)
                output_file.close()
                print("已經儲存PDF文件到: " + output_file_path)
            else:
                print("未找到指定標題")
                
            pdf_file.close()
            
            # 開啟路徑資料夾
            self.startfile(output_folder)

        except:
            print('路徑輸入錯誤，檔案必須pdf')


    # 以下是 call 程式 main
    ## 重新設計 2024/1/28

    def seletct_input(self):
        
        text = f""" 
            {Fore.GREEN}{Style.BRIGHT}消防、電力、排水一起自動下載{Style.RESET_ALL}
            {Fore.WHITE}{Style.BRIGHT}輸入{Style.RESET_ALL} {Fore.RED}{Style.BRIGHT}A{Style.RESET_ALL} {Fore.WHITE}{Style.BRIGHT}全自動 ，當月執行程式會自動下載上個月報表{Style.RESET_ALL}
            {Fore.WHITE}{Style.BRIGHT}輸入{Style.RESET_ALL} {Fore.RED}{Style.BRIGHT}B{Style.RESET_ALL}  {Fore.WHITE}{Style.BRIGHT}手動輸入 年月，範例：2022-03{Style.RESET_ALL}
            ====================================================================
            {Fore.GREEN}{Style.BRIGHT}選擇單一種類別輸出{Style.RESET_ALL}
            {Fore.WHITE}{Style.BRIGHT}輸入{Style.RESET_ALL} {Fore.RED}{Style.BRIGHT}F{Style.RESET_ALL} {Fore.WHITE}{Style.BRIGHT}消防{Style.RESET_ALL}
            {Fore.WHITE}{Style.BRIGHT}輸入{Style.RESET_ALL} {Fore.RED}{Style.BRIGHT}E{Style.RESET_ALL} {Fore.WHITE}{Style.BRIGHT}消防，指定單一棟別報表下載{Style.RESET_ALL}
            {Fore.WHITE}{Style.BRIGHT}輸入{Style.RESET_ALL} {Fore.RED}{Style.BRIGHT}C{Style.RESET_ALL} {Fore.WHITE}{Style.BRIGHT}電力{Style.RESET_ALL}
            {Fore.WHITE}{Style.BRIGHT}輸入{Style.RESET_ALL} {Fore.RED}{Style.BRIGHT}D{Style.RESET_ALL} {Fore.WHITE}{Style.BRIGHT}排水{Style.RESET_ALL}
            ====================================================================
            {Fore.GREEN}{Style.BRIGHT}電子巡檢每月報表產出PDF，搜尋特定棟別名稱合併一個PDF{Style.RESET_ALL}
            {Fore.WHITE}{Style.BRIGHT}輸入{Style.RESET_ALL} {Fore.RED}{Style.BRIGHT}PDF{Style.RESET_ALL} {Fore.WHITE}{Style.BRIGHT}每月報表 搜尋特定棟別名稱合併一個PDF{Style.RESET_ALL}
            ====================================================================
            {Fore.WHITE}{Style.BRIGHT}輸入{Style.RESET_ALL} {Fore.YELLOW}{Style.BRIGHT}X{Style.RESET_ALL} {Fore.WHITE}{Style.BRIGHT}取消(關閉){Style.RESET_ALL}
            {Style.RESET_ALL}
        """
        # print(text)

        Manually_text = f"""
            {Fore.GREEN}{Style.BRIGHT}選擇下載模式{Style.RESET_ALL}
            {Fore.WHITE}{Style.BRIGHT}輸入{Style.RESET_ALL} {Fore.RED}{Style.BRIGHT}0{Style.RESET_ALL} {Fore.WHITE}{Style.BRIGHT}自動下載上個月報表{Style.RESET_ALL}
            {Fore.WHITE}{Style.BRIGHT}輸入{Style.RESET_ALL} {Fore.RED}{Style.BRIGHT}1{Style.RESET_ALL} {Fore.WHITE}{Style.BRIGHT}指定年月報表{Style.RESET_ALL}
            {Fore.WHITE}{Style.BRIGHT}輸入{Style.RESET_ALL} {Fore.YELLOW}{Style.BRIGHT}exit{Style.RESET_ALL} {Fore.WHITE}{Style.BRIGHT}回主選單{Style.RESET_ALL}
        """

        text_1 = f"""
            {Fore.WHITE}{Style.BRIGHT}【輸入數字】{Style.RESET_ALL}
            {Fore.RED}{Style.BRIGHT}0{Style.RESET_ALL}{Fore.WHITE}{Style.BRIGHT}=自動下載上個月報表{Style.RESET_ALL}
            {Fore.RED}{Style.BRIGHT}1{Style.RESET_ALL}{Fore.WHITE}{Style.BRIGHT}=手動指定日期報表{Style.RESET_ALL}
            """

        while True:
            print(text)
            print(f'{Fore.GREEN}{Style.BRIGHT}選擇模式{Style.RESET_ALL}{Fore.YELLOW}{Style.BRIGHT}(大小寫皆可){Style.RESET_ALL}') 
            seletct_input = input(f'：')

            if seletct_input == 'A' or seletct_input == 'a':

                """
                輸入 A 全自動
                """
                date = self.get_date()
                print(f'自動下載，下載報表日期 {date}')

                # 計算run時間
                start_time = time.time() # START

                # code
                # with ThreadPoolExecutor(max_workers=10) as executor: 
                #     executor.submit(my.Fire_call, date)
                #     executor.submit(my.electricity , date)
                #     executor.submit(my.drain , date)

                # _thread.start_new_thread(my.Fire_call, (date, ))
                # _thread.start_new_thread(my.electricity, (date, ))
                # _thread.start_new_thread(my.drain, (date, ))

                
                electricity_Thread = threading.Thread(target = my.electricity, args = (date,))
                drain_Thread = threading.Thread(target = my.drain, args = (date,))
            
                electricity_Thread.start()
                drain_Thread.start()

                my.Fire_call(date)
                
                electricity_Thread.join()
                drain_Thread.join()


                end_time = time.time() # END
                    
                # 時間處理
                Time_timing = end_time - start_time
                if int(Time_timing) >= 60 :
                    time_sum = int(Time_timing) / 60
                    print(f'已下載完成，總花費時間 {int(time_sum)} 分鐘')
                else:
                    print(f'{int(Time_timing)} 秒')            

                
                # break
            elif seletct_input == 'B' or seletct_input == 'b':
                Manually = self.input_str()
                """
                輸入 B 手動輸入
                """

                print(f'輸入的日期 {Manually}')
                
                # 計算run時間
                start_time = time.time() # START

                # code
                # with ThreadPoolExecutor(max_workers=10) as executor: 
                #     executor.submit(my.Fire_call, Manually)
                #     executor.submit(my.electricity , Manually)
                #     executor.submit(my.drain , Manually)

                # _thread.start_new_thread(my.Fire_call, (Manually, ))
                # _thread.start_new_thread(my.electricity, (Manually, ))
                # _thread.start_new_thread(my.drain, (Manually, ))
                # Fire_Thread = threading.Thread(target = my.Fire_call, args = (Manually,))
                electricity_Thread = threading.Thread(target = my.electricity, args = (Manually,))
                drain_Thread = threading.Thread(target = my.drain, args = (Manually,))
                
                # Fire_Thread.start()
                electricity_Thread.start()
                drain_Thread.start()

                my.Fire_call(Manually)

                # Fire_Thread.join()
                electricity_Thread.join()
                drain_Thread.join()

                end_time = time.time() # END
                    
                # 時間處理
                Time_timing = end_time - start_time
                if int(Time_timing) >= 60 :
                    time_sum = int(Time_timing) / 60
                    print(f'已下載完成，總花費時間 {int(time_sum)} 分鐘')
                else:
                    print(f'{int(Time_timing)} 秒')  

            elif seletct_input == 'F'or seletct_input == 'f':
                """
                輸入 F 消防
                自動輸入
                指定日期
                """
                print(f'{Fore.WHITE}{Style.BRIGHT}目前選擇{Style.RESET_ALL} {Fore.RED}{Style.BRIGHT}消防{Style.RESET_ALL}')
                print(Manually_text)
                print(text_1)
                seletct = input(f'：')

                if seletct == '0':
                    print(f'F {self.get_date()}')

                    # 計算run時間
                    start_time = time.time() # START

                    #code
                    self.Fire_call(self.get_date())

                    end_time = time.time() # END
                        
                    # 時間處理
                    Time_timing = end_time - start_time
                    if int(Time_timing) >= 60 :
                        time_sum = int(Time_timing) / 60
                        print(f'已下載完成，總花費時間 {int(time_sum)} 分鐘')
                    else:
                        print(f'{int(Time_timing)} 秒')

                    # break
                elif seletct == '1':
                    Manually = self.input_str()
                    print(f'F {Manually}')
                    
                    # 計算run時間
                    start_time = time.time() # START

                    #code
                    self.Fire_call(Manually)

                    end_time = time.time() # END
                        
                    # 時間處理
                    Time_timing = end_time - start_time
                    if int(Time_timing) >= 60 :
                        time_sum = int(Time_timing) / 60
                        print(f'已下載完成，總花費時間 {int(time_sum)} 分鐘')
                    else:
                        print(f'{int(Time_timing)} 秒')

                    # break
            elif seletct_input == 'C' or seletct_input =='c':
                """
                輸入 C 電力
                自動輸入
                指定日期
                """
                print(f'{Fore.WHITE}{Style.BRIGHT}目前選擇{Style.RESET_ALL} {Fore.RED}{Style.BRIGHT}電力{Style.RESET_ALL}')
                print(Manually_text)
                print(text_1)
                seletct = input(f'：')

                if seletct == '0':
                    print(f'F {self.get_date()}')
                    
                    # 計算run時間
                    start_time = time.time() # START

                    #code
                    self.electricity(self.get_date())

                    end_time = time.time() # END
                        
                    # 時間處理
                    Time_timing = end_time - start_time
                    if int(Time_timing) >= 60 :
                        time_sum = int(Time_timing) / 60
                        print(f'已下載完成，總花費時間 {int(time_sum)} 分鐘')
                    else:
                        print(f'{int(Time_timing)} 秒')                    
                    # break
                elif seletct == '1':
                    Manually = self.input_str()
                    print(f'C {Manually}')
                    

                    # 計算run時間
                    start_time = time.time() # START

                    #code
                    self.electricity(Manually)

                    end_time = time.time() # END
                        
                    # 時間處理
                    Time_timing = end_time - start_time
                    if int(Time_timing) >= 60 :
                        time_sum = int(Time_timing) / 60
                        print(f'已下載完成，總花費時間 {int(time_sum)} 分鐘')
                    else:
                        print(f'{int(Time_timing)} 秒')
                    
                    # break
            elif seletct_input == 'D' or seletct_input == 'd':
                """
                輸入 D 排水
                自動輸入
                指定日期
                """
                print(f'{Fore.WHITE}{Style.BRIGHT}目前選擇{Style.RESET_ALL} {Fore.RED}{Style.BRIGHT}排水{Style.RESET_ALL}')
                print(Manually_text)
                print(text_1)
                seletct = input(f'：')

                if seletct == '0':
                    print(f'F {self.get_date()}')
                   
                   # 計算run時間
                    start_time = time.time() # START

                    #code
                    self.drain(self.get_date())

                    end_time = time.time() # END
                        
                    # 時間處理
                    Time_timing = end_time - start_time
                    if int(Time_timing) >= 60 :
                        time_sum = int(Time_timing) / 60
                        print(f'已下載完成，總花費時間 {int(time_sum)} 分鐘')
                    else:
                        print(f'{int(Time_timing)} 秒')                    
                    # break
                elif seletct == '1':
                    Manually = self.input_str()
                    print(f'D {Manually}')
                    
                    # 計算run時間
                    start_time = time.time() # START

                    #code
                    self.drain(Manually)

                    end_time = time.time() # END
                        
                    # 時間處理
                    Time_timing = end_time - start_time
                    if int(Time_timing) >= 60 :
                        time_sum = int(Time_timing) / 60
                        print(f'已下載完成，總花費時間 {int(time_sum)} 分鐘')
                    else:
                        print(f'{int(Time_timing)} 秒')

                    # break
            elif seletct_input == 'E' or seletct_input == 'e':
                """指定下載單一棟別報表，限定消防"""
                
                print(f'{Fore.WHITE}{Style.BRIGHT}目前選擇{Style.RESET_ALL} {Fore.RED}{Style.BRIGHT}指定單一類別報表，消防{Style.RESET_ALL}')
                print(Manually_text)
                print(text_1)
                seletct = input(f'：')

                if seletct == '0':
                    print(f'F {self.get_date()}')
                    input_text = input("輸入要下載單一類別的報表，例如：中正樓： ")
                    print()
                    # 計算run時間
                    start_time = time.time() # START

                    #code
                    self.Fire_call_find(self.get_date(), input_text)
                    # self.Fire_call(self.get_date())

                    end_time = time.time() # END
                        
                    # 時間處理
                    Time_timing = end_time - start_time
                    if int(Time_timing) >= 60 :
                        time_sum = int(Time_timing) / 60
                        print(f'已下載完成，總花費時間 {int(time_sum)} 分鐘')
                    else:
                        print(f'{int(Time_timing)} 秒')

                    # break
                elif seletct == '1':

                    Manually = self.input_str()
                    print(f'F {Manually}')

                    input_text = input("輸入要下載單一棟別的報表，例如：中正樓： ")
                    
                    # 計算run時間
                    start_time = time.time() # START

                    #code
                    # self.Fire_call(Manually)
                    
                    self.Fire_call_find(Manually, input_text)

                    end_time = time.time() # END
                        
                    # 時間處理
                    Time_timing = end_time - start_time
                    if int(Time_timing) >= 60 :
                        time_sum = int(Time_timing) / 60
                        print(f'已下載完成，總花費時間 {int(time_sum)} 分鐘')
                    else:
                        print(f'{int(Time_timing)} 秒')

                    # break                
            elif seletct_input == 'PDF' or seletct_input == 'pdf':
                input_title = input('輸入要搜尋棟別的名稱 例如： 長青樓：')
                self.pdf_report_merge(input_title)

            elif seletct_input == 'X' or seletct_input == 'x':
                print('關閉視窗')
                os.exit()
            else:
                print(text)
                print('輸入錯誤，請依說明入正確的格式')

                

if __name__ == '__main__':
    my = htmltopdf()

    my.General_folder()

    print(
        f'{Fore.WHITE}{Style.BRIGHT}開源程式碼：https://github.com/kobojp/vghtpe_Electronic_inspection_information_webtopdf{Style.RESET_ALL}\n'
        f'{Fore.YELLOW}{Style.BRIGHT}自動下載水電消防日周月報表轉PDF檔{Style.RESET_ALL}'
        f'{Fore.YELLOW}{Style.BRIGHT}更新，2024.01{Style.RESET_ALL}'
    )

    my.seletct_input()
    
    # print(my.get_date())