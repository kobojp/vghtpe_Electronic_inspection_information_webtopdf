# from distutils import ccompiler
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
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog

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
        self.File_folder = File_folder
        self.fire_folder = fire
        self.electricity_folder = electricity_folder
        self.drain_folder = drain
        self.should_stop = False  # 新增停止標誌
        self.progress_callback = None  # 新增回調函數

        # 修改 wkhtmltopdf 路徑設定
        if getattr(sys, 'frozen', False):
            # 如果是打包後的執行檔
            application_path = sys._MEIPASS
        else:
            # 如果是直接執行 Python 腳本
            application_path = os.path.dirname(os.path.abspath(__file__))
            
        self.path_wkhtmltopdf = os.path.join(application_path, 'wkhtmltopdf.exe')
        self.config = pdfkit.configuration(wkhtmltopdf=self.path_wkhtmltopdf)

    colorama.init(autoreset=True)

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

        # count file
        count_file = []
        for i in self.open_data['Fire_Equipment']:
            count_file.append(i['name'])
        print(f'{self.fire_folder} 共有 {len(count_file)} 個PDF\n')

        success_count = 0
        for i in self.open_data['Fire_Equipment']:
            name = i['name']
            url_api_1 = i['api_1']
            url_api_2 = i['api_2']
            url = f'http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}'
            output_path = os.path.join(self.File_folder, self.fire_folder, date, name) + '.pdf'
            
            if self.download_report(url, output_path, name):
                success_count += 1

        print(f'{self.fire_folder}報表下載完成，成功 {success_count}/{len(count_file)} 個檔案')

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
        """
        #建立資料夾
        self.folder(os.path.join(self.File_folder,self.electricity_folder))  #電力
        self.folder(os.path.join(self.File_folder,self.electricity_folder,date))  # 月

        date = date # ex : '2022-05'
        
        # file count
        self.count_file('electricity_every')
        
        success_count = 0
        total_count = 0

        # day
        for i in self.open_data['electricity_every_day']:
            total_count += 1
            name = i['name']
            url_api_1 = i['api_1']
            url_api_2 = i['api_2']
            url = f'http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
            output_path = os.path.join(self.File_folder, self.electricity_folder, date, name) + '.pdf'
            
            if self.download_report(url, output_path, name):
                success_count += 1

        # week
        for i in self.open_data['electricity_every_week']:
            total_count += 1
            name = i['name']
            url_api_1 = i['api_1']
            url_api_2 = i['api_2']
            url = f'http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
            output_path = os.path.join(self.File_folder, self.electricity_folder, date, name) + '.pdf'
            
            if self.download_report(url, output_path, name):
                success_count += 1

        # month
        for i in self.open_data['electricity_every_month']:
            total_count += 1
            name = i['name']
            url_api_1 = i['api_1']
            url_api_2 = i['api_2']
            
            if name == '中正樓24F停機坪照明設備巡檢紀錄':
                url = f'http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}'
            else:
                url = f'http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
            
            output_path = os.path.join(self.File_folder, self.electricity_folder, date, name) + '.pdf'
            
            if self.download_report(url, output_path, name):
                success_count += 1

        print(f'{self.electricity_folder}報表下載完成，成功 {success_count}/{total_count} 個檔案')

        # open folder
        start_directory = os.path.join(self.File_folder,self.electricity_folder,date) 
        self.startfile(start_directory)

    # 給排水設備
    def drain(self, date:str):
        """
        給排水設備
        輸入值 YYYY-DD
        """
        #建立資料夾
        self.folder(os.path.join(self.File_folder,self.drain_folder))  #給排水
        self.folder(os.path.join(self.File_folder,self.drain_folder,date))  # 月

        date = date # ex : '2022-05'
        
        # file count
        self.count_file('drain')
        
        success_count = 0
        total_count = 0

        # day
        for i in self.open_data['drain_day']:
            total_count += 1
            name = i['name']
            url_api_1 = i['api_1']
            url_api_2 = i['api_2']
            url = f'http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
            output_path = os.path.join(self.File_folder, self.drain_folder, date, name) + '.pdf'
            
            if self.download_report(url, output_path, name):
                success_count += 1

        # week
        for i in self.open_data['drain_week']:
            total_count += 1
            name = i['name']
            url_api_1 = i['api_1']
            url_api_2 = i['api_2']
            url = f'http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
            output_path = os.path.join(self.File_folder, self.drain_folder, date, name) + '.pdf'
            
            if self.download_report(url, output_path, name):
                success_count += 1

        # month
        for i in self.open_data['drain_month']:
            total_count += 1
            name = i['name']
            url_api_1 = i['api_1']
            url_api_2 = i['api_2']
            url = f'http://210.61.217.104/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
            output_path = os.path.join(self.File_folder, self.drain_folder, date, name) + '.pdf'
            
            if self.download_report(url, output_path, name):
                success_count += 1

        print(f'{self.drain_folder}報表下載完成，成功 {success_count}/{total_count} 個檔案')

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
        """取得當前時間"""
        now = datetime.datetime.now()
        return now.strftime("%H:%M:%S")  # 只返回時間，不加入顏色代碼
    
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

    def set_progress_callback(self, callback):
        """設置進度回調函數"""
        self.progress_callback = callback
        
    def log_progress(self, message):
        """記錄進度"""
        if self.progress_callback:
            self.progress_callback(message)
        print(message)  # 同時保留控制台輸出

    def download_report(self, url, output_path, name, max_retries=3):
        """下載並儲存報表，包含重試機制"""
        for attempt in range(max_retries):
            if self.should_stop:
                raise Exception("使用者取消下載")
                
            try:
                progress_msg = f'{self.get_now_date()} 正在下載 {name}\n'
                self.log_progress(progress_msg)
                self.log_progress(f'URL: {url}\n')
                
                options = {
                    'no-background': None,
                    'disable-javascript': None,
                    'enable-local-file-access': None,
                    'javascript-delay': 2000,
                    'no-stop-slow-scripts': None,
                    'load-error-handling': 'ignore',
                    'load-media-error-handling': 'ignore',
                    'custom-header': [
                        ('Accept-Encoding', 'gzip')
                    ],
                    'encoding': 'utf-8',
                    'quiet': None
                }
                
                pdfkit.from_url(url, output_path,
                               options=options,
                               configuration=self.config)
                
                if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                    self.log_progress(f'{self.get_now_date()} {name} 下載完成\n')
                    return True
                else:
                    raise Exception("檔案建立失敗或大小為0")
                
            except Exception as e:
                if self.should_stop:
                    raise Exception("使用者取消下載")
                    
                if attempt < max_retries - 1:
                    wait_time = (attempt + 1) * 5
                    self.log_progress(f'下載失敗 (嘗試 {attempt + 1}/{max_retries}): {str(e)}\n')
                    self.log_progress(f'等待 {wait_time} 秒後重試...\n')
                    time.sleep(wait_time)
                else:
                    self.log_progress(f'下載失敗 {name}: {str(e)}\n')
                    return False

    def stop_all_downloads(self):
        """停止所有下載"""
        self.should_stop = True

    def run_all_reports(self, date):
        """並行執行所有報表下載任務"""
        start_time = time.time()
        
        # 建立資料夾結構
        self.folder(os.path.join(self.File_folder, self.fire_folder))
        self.folder(os.path.join(self.File_folder, self.electricity_folder))
        self.folder(os.path.join(self.File_folder, self.drain_folder))
        
        if date:
            self.folder(os.path.join(self.File_folder, self.fire_folder, date))
            self.folder(os.path.join(self.File_folder, self.electricity_folder, date))
            self.folder(os.path.join(self.File_folder, self.drain_folder, date))
        
        # 建立執行緒並使用事件來追蹤完成狀態
        threads = []
        completion_events = []
        
        # 使用中文名稱對應
        task_names = {
            self.electricity: '電力',
            self.drain: '排水',
            self.Fire_call: '消防'
        }
        
        tasks = [
            (self.electricity, '電力'),
            (self.drain, '排水'),
            (self.Fire_call, '消防')
        ]
        
        for task_func, task_name in tasks:
            if self.should_stop:
                break
                
            try:
                self.log_progress(f'開始下載{task_name}報表...\n')
                task_func(date)
                self.log_progress(f'{task_name}報表下載完成\n')
            except Exception as e:
                self.log_progress(f'{task_name}報表下載發生錯誤: {str(e)}\n')
                if self.should_stop:
                    break
        
        # 計算耗時
        duration = time.time() - start_time
        if duration >= 60:
            minutes = int(duration / 60)
            self.log_progress(f'已下載完成，總花費時間 {minutes} 分鐘\n')
        else:
            self.log_progress(f'已下載完成，總花費時間 {int(duration)} 秒\n')

# GUI
class HtmlToPdfGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("自動下載水電消防報表系統")
        self.geometry("800x550")
        
        # 定義按鈕樣式
        self.button_style = {
            'bg': '#0067B8',         # rgb(0, 103, 184) 轉換為十六進制
            'fg': 'white',           # 文字白色
            'font': ('微軟正黑體', 9),
            'relief': 'flat',
            'padx': 10,
            'pady': 5,
            'cursor': 'hand2',       # 滑鼠變成手指形狀
            'activebackground': '#005BA1',  # 按下時的背景色（稍微深一點）
            'activeforeground': 'white',    # 按下時的文字顏色
            'disabledforeground': 'black'   # 停用時的文字顏色改為黑色
        }
        
        # 建立 htmltopdf 實例
        self.pdf_handler = htmltopdf()
        self.pdf_handler.set_progress_callback(self.update_progress)
        
        # 控制下載狀態的變數
        self.is_downloading = False
        
        # 建立主要框架和所有 GUI 元件
        self.create_widgets()
        
        # 建立必要的資料夾 (移到最後)
        self.create_required_folders()

    def create_required_folders(self):
        """建立必要的資料夾"""
        required_folders = [
            '水電消防報表',
            os.path.join('水電消防報表', '消防'),
            os.path.join('水電消防報表', '電力'),
            os.path.join('水電消防報表', '排水'),
            '報表合併pdf'  # PDF合併功能需要的資料夾
        ]
        
        for folder in required_folders:
            if not os.path.exists(folder):
                try:
                    os.makedirs(folder)
                    self.update_progress(f"已建立資料夾：{folder}\n")
                except Exception as e:
                    self.update_progress(f"建立資料夾 {folder} 時發生錯誤：{str(e)}\n")

    def create_widgets(self):
        """建立所有 GUI 元件"""
        # 建立主框架
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 左側區域
        left_frame = ttk.Frame(main_frame)
        left_frame.grid(row=0, column=0, padx=5, sticky=(tk.N, tk.S))
        
        # 右側區域
        right_frame = ttk.Frame(main_frame)
        right_frame.grid(row=0, column=1, padx=5, sticky=(tk.N, tk.S))
        
        # === 右側元件 ===
        # 報表查詢區域
        report_frame = ttk.LabelFrame(right_frame, text="報表查詢", padding="10")
        report_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N))
        
        # 報表類型選擇
        type_frame = ttk.Frame(report_frame)
        type_frame.grid(row=0, column=0, padx=5)
        
        ttk.Label(
            type_frame,
            text="報表類型:"
        ).grid(row=0, column=0, padx=5, pady=5)
        
        self.report_type_var = tk.StringVar()
        report_types = [
            ("消防", "Fire_Equipment"),
            ("電力每日", "electricity_every_day"),
            ("電力每月", "electricity_every_month"),
            ("電力每周", "electricity_every_week"),
            ("排水每日", "drain_day"),
            ("排水每月", "drain_month"),
            ("排水每周", "drain_week")
        ]
        
        self.report_type_combo = ttk.Combobox(
            type_frame,
            textvariable=self.report_type_var,
            values=[t[0] for t in report_types],
            state="readonly",
            width=15
        )
        self.report_type_combo.grid(row=0, column=1, padx=5, pady=5)
        self.report_type_combo.bind('<<ComboboxSelected>>', self.update_report_names)
        
        # 報表名稱列表
        list_frame = ttk.Frame(report_frame)
        list_frame.grid(row=1, column=0, padx=5, pady=5)
        
        ttk.Label(
            list_frame,
            text="報表名稱:"
        ).grid(row=0, column=0, sticky=tk.W)
        
        # 建立報表名稱顯示區域和捲動條
        self.report_list = tk.Listbox(list_frame, width=35, height=20)
        self.report_list.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 加入垂直捲動條
        list_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.report_list.yview)
        list_scrollbar.grid(row=1, column=1, sticky="ns")
        self.report_list.configure(yscrollcommand=list_scrollbar.set)
        
        # === 左側元件 ===
        # 標題標籤
        title_label = ttk.Label(
            left_frame, 
            text="自動下載水電消防日周月報表轉PDF檔", 
            font=('微軟正黑體', 14, 'bold')
        )
        title_label.grid(row=0, column=0, columnspan=2, pady=10)
        
        # 自動下載區域
        auto_frame = ttk.LabelFrame(left_frame, text="自動下載", padding="10")
        auto_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=3)
        
        tk.Button(
            auto_frame, 
            text="自動下載上個月報表",
            command=self.auto_download,
            **self.button_style
        ).grid(row=0, column=0, padx=5)
        
        tk.Button(
            auto_frame, 
            text="手動輸入日期下載", 
            command=self.manual_download,
            **self.button_style
        ).grid(row=0, column=1, padx=5)
        
        # 單一類別下載區域
        single_frame = ttk.LabelFrame(left_frame, text="單一類別下載", padding="10")
        single_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=3)
        
        tk.Button(
            single_frame, 
            text="消防報表", 
            command=lambda: self.download_single_type("fire"),
            **self.button_style
        ).grid(row=0, column=0, padx=5)
        
        tk.Button(
            single_frame, 
            text="電力報表", 
            command=lambda: self.download_single_type("electricity"),
            **self.button_style
        ).grid(row=0, column=1, padx=5)
        
        tk.Button(
            single_frame, 
            text="排水報表", 
            command=lambda: self.download_single_type("drain"),
            **self.button_style
        ).grid(row=0, column=2, padx=5)
        
        # PDF 合併區域
        merge_frame = ttk.LabelFrame(left_frame, text="PDF合併功能", padding="10")
        merge_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=3)
        
        ttk.Label(
            merge_frame, 
            text="搜尋棟別:"
        ).grid(row=0, column=0, padx=5)
        
        self.building_var = tk.StringVar()
        ttk.Entry(
            merge_frame, 
            textvariable=self.building_var
        ).grid(row=0, column=1, padx=5)
        
        tk.Button(
            merge_frame, 
            text="合併PDF", 
            command=self.merge_pdf,
            **self.button_style
        ).grid(row=0, column=2, padx=5)
        
        # 進度顯示區域
        progress_frame = ttk.Frame(left_frame)
        progress_frame.grid(row=4, column=0, columnspan=2, pady=3)
        
        # 建立文字區域和捲動條
        text_frame = ttk.Frame(progress_frame)
        text_frame.grid(row=0, column=0, columnspan=2)
        
        self.progress_text = tk.Text(text_frame, height=8, width=50)
        self.progress_text.grid(row=0, column=0)
        
        # 加入垂直捲動條
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.progress_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.progress_text.configure(yscrollcommand=scrollbar.set)
        
        # 加入水平捲動條
        h_scrollbar = ttk.Scrollbar(text_frame, orient="horizontal", command=self.progress_text.xview)
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        self.progress_text.configure(xscrollcommand=h_scrollbar.set, wrap=tk.NONE)  # 設置不自動換行
        
        # 清除按鈕
        button_frame = ttk.Frame(progress_frame)
        button_frame.grid(row=1, column=0, columnspan=2, pady=5)
        
        tk.Button(
            button_frame,
            text="清除記錄",
            command=self.clear_progress,
            **self.button_style
        ).pack(side=tk.LEFT, padx=5)
        
        # 加入控制按鈕區域
        control_frame = ttk.Frame(left_frame)
        control_frame.grid(row=5, column=0, columnspan=2, pady=3)
        
        self.start_button = tk.Button(
            control_frame,
            text="開始下載",
            command=self.auto_download,
            **self.button_style
        )
        self.start_button.grid(row=0, column=0, padx=5)
        
        self.stop_button = tk.Button(
            control_frame,
            text="停止下載",
            command=self.stop_download,
            state=tk.DISABLED,
            **self.button_style
        )
        self.stop_button.grid(row=0, column=1, padx=5)
        
    def update_report_names(self, event=None):
        """更新報表名稱列表"""
        type_map = {
            "消防": "Fire_Equipment",
            "電力每日": "electricity_every_day",
            "電力每月": "electricity_every_month",
            "電力每周": "electricity_every_week",
            "排水每日": "drain_day",
            "排水每月": "drain_month",
            "排水每周": "drain_week"
        }
        
        selected_type = self.report_type_var.get()
        data_key = type_map.get(selected_type)
        
        # 清空列表
        self.report_list.delete(0, tk.END)
        
        if data_key and data_key in self.pdf_handler.open_data:
            names = [item["name"] for item in self.pdf_handler.open_data[data_key]]
            for name in names:
                self.report_list.insert(tk.END, name)
        
    def update_progress(self, message):
        """更新進度顯示"""
        self.progress_text.insert(tk.END, message)
        self.progress_text.see(tk.END)  # 自動捲動到最新位置
        self.update_idletasks()
        
    def auto_download(self):
        """自動下載上個月報表"""
        if self.is_downloading:
            return
            
        self.is_downloading = True
        self.pdf_handler.should_stop = False
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        
        date = self.pdf_handler.get_date()
        self.update_progress(f"開始下載 {date} 報表...\n")
        
        self.download_thread = threading.Thread(
            target=self._download_task,
            args=(date,)
        )
        self.download_thread.start()
        
    def _download_task(self, date):
        """執行下載任務的執行緒"""
        try:
            self.pdf_handler.run_all_reports(date)
            if self.is_downloading:  # 檢查是否被停止
                self.update_progress("下載完成！\n")
        except Exception as e:
            if self.is_downloading:  # 檢查是否被停止
                messagebox.showerror("錯誤", f"下載過程發生錯誤：{str(e)}")
        finally:
            self.is_downloading = False
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            
    def stop_download(self):
        """停止下載"""
        if self.is_downloading:
            self.pdf_handler.stop_all_downloads()  # 設置停止標誌
            self.update_progress("正在停止下載...\n")
            
            # 等待執行緒結束
            if hasattr(self, 'download_thread') and self.download_thread.is_alive():
                self.download_thread.join(timeout=5)  # 最多等待5秒
                
            self.is_downloading = False
            self.update_progress("下載已停止\n")
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
                
    def manual_download(self):
        """手動輸入日期下載"""
        dialog = DateInputDialog(self)
        if dialog.result:
            date = dialog.result
            self.update_progress(f"開始下載 {date} 報表...\n")
            
            try:
                self.pdf_handler.run_all_reports(date)
                self.update_progress("下載完成！\n")
            except Exception as e:
                messagebox.showerror("錯誤", f"下載過程發生錯誤：{str(e)}")
                
    def download_single_type(self, type_name):
        """下載單一類別報表"""
        if self.is_downloading:
            return
        
        dialog = DateInputDialog(self)
        if dialog.result:
            date = dialog.result
            
            # 轉換類型名稱為中文
            type_names = {
                "fire": "消防",
                "electricity": "電力",
                "drain": "排水"
            }
            display_name = type_names.get(type_name, type_name)
            
            self.is_downloading = True
            self.pdf_handler.should_stop = False
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            
            self.update_progress(f"開始下載{display_name}報表 {date}...\n")
            
            # 使用執行緒來執行下載
            self.download_thread = threading.Thread(
                target=self._download_single_task,
                args=(type_name, date)
            )
            self.download_thread.start()

    def _download_single_task(self, type_name, date):
        """執行單一類別下載任務的執行緒"""
        try:
            if type_name == "fire":
                self.pdf_handler.Fire_call(date)
            elif type_name == "electricity":
                self.pdf_handler.electricity(date)
            elif type_name == "drain":
                self.pdf_handler.drain(date)
            
            if self.is_downloading:  # 檢查是否被停止
                self.update_progress("下載完成！\n")
        except Exception as e:
            if self.is_downloading:  # 檢查是否被停止
                messagebox.showerror("錯誤", f"下載過程發生錯誤：{str(e)}")
        finally:
            self.is_downloading = False
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
                
    def merge_pdf(self):
        """合併PDF功能"""
        building = self.building_var.get()
        if not building:
            messagebox.showwarning("警告", "請輸入要搜尋的棟別名稱")
            return
            
        file_path = filedialog.askopenfilename(
            title="選擇PDF檔案",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if file_path:
            try:
                self.pdf_handler.pdf_report_merge(building)
                self.update_progress("PDF合併完成！\n")
            except Exception as e:
                messagebox.showerror("錯誤", f"PDF合併過程發生錯誤：{str(e)}")

    def clear_progress(self):
        """清除進度顯示區域的內容"""
        self.progress_text.delete('1.0', tk.END)
        self.update_idletasks()

class DateInputDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        
        self.title("輸入日期")
        self.result = None
        
        # 設定對話框大小和位置
        dialog_width = 300
        dialog_height = 100
        
        # 取得主視窗位置和大小
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        # 計算對話框應該出現的位置
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        # 設定對話框位置
        self.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        
        # 建立日期輸入框
        ttk.Label(
            self, 
            text="請輸入日期 (YYYY-MM):"
        ).grid(row=0, column=0, padx=5, pady=5)
        
        self.date_var = tk.StringVar()
        # 設置預設值為上個月
        last_month = self.get_last_month()
        self.date_var.set(last_month)
        
        ttk.Entry(
            self, 
            textvariable=self.date_var
        ).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Button(
            self, 
            text="確定", 
            command=self.confirm
        ).grid(row=1, column=0, columnspan=2, pady=10)
        
        self.transient(parent)
        self.grab_set()
        parent.wait_window(self)
    
    def get_last_month(self):
        """取得上個月的年月份"""
        today = datetime.datetime.now()
        first_day_this_month = today.replace(day=1)
        last_day_last_month = first_day_this_month - datetime.timedelta(days=1)
        return last_day_last_month.strftime("%Y-%m")
        
    def confirm(self):
        """確認日期輸入"""
        date = self.date_var.get()
        if self.validate_date(date):
            self.result = date
            self.destroy()
        else:
            messagebox.showwarning("警告", "請輸入正確的日期格式 (YYYY-MM)")
            
    def validate_date(self, date):
        """驗證日期格式"""
        try:
            year = int(date[:4])
            month = int(date[5:7])
            return (len(date) == 7 and 
                   year >= 2022 and
                   date[4] == '-' and
                   1 <= month <= 12)
        except:
            return False

if __name__ == '__main__':
    app = HtmlToPdfGUI()
    app.mainloop()