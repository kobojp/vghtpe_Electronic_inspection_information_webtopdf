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
from PIL import Image, ImageTk


REPORT_TYPE_MAP = {
    "消防": "Fire_Equipment",
    "電力每日": "electricity_every_day",
    "電力每月": "electricity_every_month",
    "電力每周": "electricity_every_week",
    "排水每日": "drain_day",
    "排水每月": "drain_month",
    "排水每周": "drain_week"
}

DEFAULT_SETTINGS = {
    "open_folder_after_completion": False
}


def validate_month(value):
    """驗證 YYYY-MM 格式，且年份不得早於 2022。"""
    try:
        parsed = datetime.datetime.strptime(value, "%Y-%m")
        return parsed.strftime("%Y-%m") == value and parsed.year >= 2022
    except (TypeError, ValueError):
        return False


def get_month_range(start_month, end_month):
    """取得包含起訖月份的 YYYY-MM 清單。"""
    if not validate_month(start_month) or not validate_month(end_month):
        raise ValueError("請輸入正確的日期格式 (YYYY-MM)，年份不得早於 2022")

    start = datetime.datetime.strptime(start_month, "%Y-%m")
    end = datetime.datetime.strptime(end_month, "%Y-%m")
    if start > end:
        raise ValueError("開始月份不得晚於結束月份")

    months = []
    current = start
    while current <= end:
        months.append(current.strftime("%Y-%m"))
        if current.month == 12:
            current = current.replace(year=current.year + 1, month=1)
        else:
            current = current.replace(month=current.month + 1)
    return months


def get_last_month():
    """取得上個月的 YYYY-MM。"""
    first_day_this_month = datetime.datetime.now().replace(day=1)
    return (first_day_this_month - datetime.timedelta(days=1)).strftime("%Y-%m")


def get_smooth_progress_value(current, target):
    """讓顯示進度平滑追趕目標值，避免一次跳動過大。"""
    current = max(0.0, min(100.0, float(current)))
    target = max(0.0, min(100.0, float(target)))
    difference = target - current
    if abs(difference) < 0.1:
        return target

    step = min(max(abs(difference) * 0.12, 0.35), 2.5)
    if difference > 0:
        return min(current + step, target)
    return max(current - step, target)


def build_report_url(report, report_type, month):
    """依報表類型建立單月份下載 URL。"""
    if not validate_month(month):
        raise ValueError("無效的報表月份")

    if report_type == "消防":
        return f'https://vghtpe-ue.httc.com.tw/Report6{report["api_1"]}{month}{report["api_2"]}'

    if report_type.startswith("電力") or report_type.startswith("排水"):
        year, month_number = (int(part) for part in month.split('-'))
        last_day = calendar.monthrange(year, month_number)[1]
        return (
            f'https://vghtpe-ue.httc.com.tw/Report6BatchAll{report["api_1"]}'
            f'{month}-01/{month}-{last_day}{report["api_2"]}'
        )

    raise ValueError("未知的報表類型")


def get_settings_path():
    """取得使用者設定檔路徑。"""
    appdata = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(appdata, "VghtpeReportDownloader", "settings.json")


def load_settings(settings_path=None):
    """讀取設定；檔案不存在或損壞時使用安全預設值。"""
    settings = DEFAULT_SETTINGS.copy()
    path = settings_path or get_settings_path()
    try:
        with open(path, encoding="utf-8") as settings_file:
            loaded = json.load(settings_file)
        if isinstance(loaded.get("open_folder_after_completion"), bool):
            settings["open_folder_after_completion"] = loaded["open_folder_after_completion"]
    except (OSError, ValueError, AttributeError):
        pass
    return settings


def save_settings(open_folder_after_completion, settings_path=None):
    """儲存使用者設定。"""
    path = settings_path or get_settings_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as settings_file:
        json.dump(
            {"open_folder_after_completion": bool(open_folder_after_completion)},
            settings_file,
            ensure_ascii=False,
            indent=2
        )

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
    def __init__(self, File_folder='水電消防報表', fire='消防', electricity_folder='電力', drain='排水', open_folder_after_completion=False):
        self.File_folder = File_folder
        self.fire_folder = fire
        self.electricity_folder = electricity_folder
        self.drain_folder = drain
        self.should_stop = False  # 新增停止標誌
        self.progress_callback = None  # 新增回調函數
        self.open_folder_after_completion = open_folder_after_completion

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
    def Fire_call(self, date:str, progress_callback=None):
        """消防設備下載"""
        try:
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
                if self.should_stop:
                    raise Exception("使用者取消下載")
                name = i['name']
                url_api_1 = i['api_1']
                url_api_2 = i['api_2']
                url = f'https://vghtpe-ue.httc.com.tw/Report6{url_api_1}{date}{url_api_2}'
                output_path = os.path.join(self.File_folder, self.fire_folder, date, name) + '.pdf'
                
                if self.download_report(url, output_path, name):
                    success_count += 1
                    if progress_callback:
                        progress_callback(True)
                else:
                    if progress_callback:
                        progress_callback(False)

            print(f'{self.fire_folder}報表下載完成，成功 {success_count}/{len(count_file)} 個檔案')

            # open folder
            start_directory = os.path.join(self.File_folder, self.fire_folder, date)
            self.startfile(start_directory)
        
        except Exception as e:
            if str(e) == "使用者取消下載":
                raise
            print(f'下載過程發生錯誤: {str(e)}')

    # 消防 指定搜尋單一類別
    def Fire_call_find(self, date:str, input:str):
        """
        消防設備
        輸入值 YYYY-DD
        """
        date = date  # ex : '2022-05'
        find = []
        #建立資料夾
        self.folder(os.path.join(self.File_folder,self.fire_folder))  #消防
        self.folder(os.path.join(self.File_folder,self.fire_folder,date))  # 月

        # 指定搜尋清單，例如：長青樓，正規 模糊搜尋要的清單
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

        # 使用清單 list 作為判斷，如果有資料 >0 就會執行下載
        success_count = 0
        if len(find) > 0:
            try:
                for i in self.open_data['Fire_Equipment']:
                    if re.search(input ,i['name']):
                        name = i['name']
                        url_api_1 = i['api_1']
                        url_api_2 = i['api_2']
                        url = f'https://vghtpe-ue.httc.com.tw/Report6{url_api_1}{date}{url_api_2}'
                        output_path = os.path.join(self.File_folder, self.fire_folder, date, name) + '.pdf'
                        
                        if self.download_report(url, output_path, name):
                            success_count += 1

                print(f'搜尋下載完成，成功 {success_count}/{len(find)} 個檔案')

                # open folder
                start_directory = os.path.join(self.File_folder, self.fire_folder, date)
                self.startfile(start_directory)
                    
            except Exception as e:
                print(f'下載過程發生錯誤: {str(e)}')

    # 電力設備
    def electricity(self, date:str, progress_callback=None):
        """電力設備下載"""
        try:
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
                if self.should_stop:
                    raise Exception("使用者取消下載")
                total_count += 1
                name = i['name']
                url_api_1 = i['api_1']
                url_api_2 = i['api_2']
                url = f'https://vghtpe-ue.httc.com.tw/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
                output_path = os.path.join(self.File_folder, self.electricity_folder, date, name) + '.pdf'
                
                if self.download_report(url, output_path, name):
                    success_count += 1
                    if progress_callback:
                        progress_callback(True)
                else:
                    if progress_callback:
                        progress_callback(False)

            # week
            for i in self.open_data['electricity_every_week']:
                if self.should_stop:
                    raise Exception("使用者取消下載")
                total_count += 1
                name = i['name']
                url_api_1 = i['api_1']
                url_api_2 = i['api_2']
                url = f'https://vghtpe-ue.httc.com.tw/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
                output_path = os.path.join(self.File_folder, self.electricity_folder, date, name) + '.pdf'
                
                if self.download_report(url, output_path, name):
                    success_count += 1
                    if progress_callback:
                        progress_callback(True)
                else:
                    if progress_callback:
                        progress_callback(False)

            # month
            for i in self.open_data['electricity_every_month']:
                if self.should_stop:
                    raise Exception("使用者取消下載")
                total_count += 1
                name = i['name']
                url_api_1 = i['api_1']
                url_api_2 = i['api_2']
                
                if name == '中正樓24F停機坪照明設備巡檢紀錄':
                    url = f'https://vghtpe-ue.httc.com.tw/Report6{url_api_1}{date}{url_api_2}'
                else:
                    url = f'https://vghtpe-ue.httc.com.tw/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
                
                output_path = os.path.join(self.File_folder, self.electricity_folder, date, name) + '.pdf'
                
                if self.download_report(url, output_path, name):
                    success_count += 1
                    if progress_callback:
                        progress_callback(True)
                else:
                    if progress_callback:
                        progress_callback(False)

            print(f'{self.electricity_folder}報表下載完成，成功 {success_count}/{total_count} 個檔案')

            # open folder
            start_directory = os.path.join(self.File_folder,self.electricity_folder,date) 
            self.startfile(start_directory)

        except Exception as e:
            if str(e) == "使用者取消下載":
                raise
            print(f'下載過程發生錯誤: {str(e)}')

    # 給排水設備
    def drain(self, date:str, progress_callback=None):
        """排水設備下載"""
        try:
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
                if self.should_stop:
                    raise Exception("使用者取消下載")
                total_count += 1
                name = i['name']
                url_api_1 = i['api_1']
                url_api_2 = i['api_2']
                url = f'https://vghtpe-ue.httc.com.tw/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
                output_path = os.path.join(self.File_folder, self.drain_folder, date, name) + '.pdf'
                
                if self.download_report(url, output_path, name):
                    success_count += 1
                    if progress_callback:
                        progress_callback(True)
                else:
                    if progress_callback:
                        progress_callback(False)

            # week
            for i in self.open_data['drain_week']:
                if self.should_stop:
                    raise Exception("使用者取消下載")
                total_count += 1
                name = i['name']
                url_api_1 = i['api_1']
                url_api_2 = i['api_2']
                url = f'https://vghtpe-ue.httc.com.tw/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
                output_path = os.path.join(self.File_folder, self.drain_folder, date, name) + '.pdf'
                
                if self.download_report(url, output_path, name):
                    success_count += 1
                    if progress_callback:
                        progress_callback(True)
                else:
                    if progress_callback:
                        progress_callback(False)

            # month
            for i in self.open_data['drain_month']:
                if self.should_stop:
                    raise Exception("使用者取消下載")
                total_count += 1
                name = i['name']
                url_api_1 = i['api_1']
                url_api_2 = i['api_2']
                url = f'https://vghtpe-ue.httc.com.tw/Report6BatchAll{url_api_1}{date}-01/{date}-{self.get_monthrange(date)}{url_api_2}'
                output_path = os.path.join(self.File_folder, self.drain_folder, date, name) + '.pdf'
                
                if self.download_report(url, output_path, name):
                    success_count += 1
                    if progress_callback:
                        progress_callback(True)
                else:
                    if progress_callback:
                        progress_callback(False)

            print(f'{self.drain_folder}報表下載完成，成功 {success_count}/{total_count} 個檔案')

            # open folder
            start_directory = os.path.join(self.File_folder,self.drain_folder,date) 
            self.startfile(start_directory)

        except Exception as e:
            if str(e) == "使用者取消下載":
                raise
            print(f'下載過程發生錯誤: {str(e)}')


    # Finish Open the folder
    def startfile(sele, filename):
        if not sele.open_folder_after_completion:
            return
        sele.open_folder(filename)

    def open_folder(sele, filename):
        """明確開啟指定資料夾，不受自動開啟設定影響。"""
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
            if len(text)==7 and int(text[:4]) >= 2022 and text[4] == '-' and text[:4].isnumeric() \…9437 tokens truncated…ame, date)
            )
            self.download_thread.start()

    def _download_single_task(self, type_name, date):
        """執行單一類別下載任務的執行緒"""
        try:
            self.reset_progress_bar()
            
            if type_name == "fire":
                # 計算消防報表總數
                total_files = len(self.pdf_handler.open_data['Fire_Equipment'])
                current_file = 0
                
                # 更新進度條顯示
                self.update_progress(f"開始下載消防報表，共 {total_files} 個檔案\n")
                self.update_progress_bar(0, total_files)
                
                def progress_callback(success):
                    nonlocal current_file
                    if self.should_stop:
                        raise Exception("使用者取消下載")
                    current_file += 1
                    self.update_progress_bar(current_file, total_files)
                    if success:
                        self.update_progress(f"已完成 {current_file}/{total_files} 個檔案\n")
                
                try:
                    if self.should_stop:
                        raise Exception("使用者取消下載")
                    self.pdf_handler.Fire_call(date, progress_callback=progress_callback)
                except Exception as e:
                    if str(e) == "使用者取消下載":
                        raise
                    self.update_progress(f"下載過程發生錯誤: {str(e)}\n")
                
            elif type_name == "electricity":
                # 計算電力報表總數
                total_files = (len(self.pdf_handler.open_data['electricity_every_day']) +
                             len(self.pdf_handler.open_data['electricity_every_week']) +
                             len(self.pdf_handler.open_data['electricity_every_month']))
                current_file = 0
                
                # 更新進度條顯示
                self.update_progress(f"開始下載電力報表，共 {total_files} 個檔案\n")
                self.update_progress_bar(0, total_files)
                
                def progress_callback(success):
                    nonlocal current_file
                    if self.should_stop:
                        raise Exception("使用者取消下載")
                    current_file += 1
                    self.update_progress_bar(current_file, total_files)
                    if success:
                        self.update_progress(f"已完成 {current_file}/{total_files} 個檔案\n")
                
                try:
                    if self.should_stop:
                        raise Exception("使用者取消下載")
                    self.pdf_handler.electricity(date, progress_callback=progress_callback)
                except Exception as e:
                    if str(e) == "使用者取消下載":
                        raise
                    self.update_progress(f"下載過程發生錯誤: {str(e)}\n")
                
            elif type_name == "drain":
                # 計算排水報表總數
                total_files = (len(self.pdf_handler.open_data['drain_day']) +
                             len(self.pdf_handler.open_data['drain_week']) +
                             len(self.pdf_handler.open_data['drain_month']))
                current_file = 0
                
                # 更新進度條顯示
                self.update_progress(f"開始下載排水報表，共 {total_files} 個檔案\n")
                self.update_progress_bar(0, total_files)
                
                def progress_callback(success):
                    nonlocal current_file
                    if self.should_stop:
                        raise Exception("使用者取消下載")
                    current_file += 1
                    self.update_progress_bar(current_file, total_files)
                    if success:
                        self.update_progress(f"已完成 {current_file}/{total_files} 個檔案\n")
                
                try:
                    if self.should_stop:
                        raise Exception("使用者取消下載")
                    self.pdf_handler.drain(date, progress_callback=progress_callback)
                except Exception as e:
                    if str(e) == "使用者取消下載":
                        raise
                    self.update_progress(f"下載過程發生錯誤: {str(e)}\n")
            
            if not self.should_stop:
                self.update_progress("下載完成！\n")
            
        except Exception as e:
            if str(e) == "使用者取消下載":
                self.update_progress("下載已被取消\n")
            else:
                messagebox.showerror("錯誤", f"下載過程發生錯誤：{str(e)}")
        finally:
            self.is_downloading = False
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            self.should_stop = False
            self.pdf_handler.should_stop = False
            self.reset_progress_bar()
                
    def merge_pdf(self):
        """合併PDF功能"""
        if hasattr(self, 'is_merging') and self.is_merging:
            return
        
        building = self.building_var.get()
        if not building:
            messagebox.showwarning("警告", "請輸入要搜尋的關鍵字")
            return
            
        # 使用執行緒來執行合併
        self.is_merging = True
        self.should_stop = False  # 重置停止標誌
        self.pdf_handler.should_stop = False  # 重置PDF處理器的停止標誌
        self.stop_merge_button.config(state=tk.NORMAL)
        
        self.merge_thread = threading.Thread(
            target=self._merge_pdf_task,
            args=(building,)
        )
        self.merge_thread.start()

    def _merge_pdf_task(self, building):
        """執行PDF合併任務的執行緒"""
        try:
            self.reset_progress_bar()
            
            def progress_callback(message, progress=None):
                if self.should_stop:  # 檢查是否需要停止
                    raise Exception("使用者取消合併")
                self.update_progress(message)
                if progress is not None:
                    self.update_progress_bar(progress, 100)
            
            def should_stop_callback():
                return self.should_stop
            
            # 執行合併
            try:
                # 重置停止標誌
                self.should_stop = False
                self.pdf_handler.should_stop = False
                
                self.pdf_handler.pdf_report_merge(
                    building,
                    progress_callback=progress_callback,
                    should_stop_callback=should_stop_callback
                )
            except Exception as e:
                if str(e) == "使用者取消合併":
                    self.update_progress("合併已被取消\n")
                else:
                    raise
            
        except Exception as e:
            if str(e) != "使用者取消合併":
                self.update_progress(f"合併過程發生錯誤: {str(e)}\n")
        finally:
            self.is_merging = False
            self.stop_merge_button.config(state=tk.DISABLED)
            self.reset_progress_bar()
            self.should_stop = False  # 重置停止標誌

    def stop_merge(self):
        """停止PDF合併"""
        if hasattr(self, 'is_merging') and self.is_merging:
            self.should_stop = True
            self.pdf_handler.should_stop = True  # 同時設置PDF處理器的停止標誌
            self.update_progress("正在停止合併...\n")
            
            # 等待合併執行緒結束
            if hasattr(self, 'merge_thread') and self.merge_thread.is_alive():
                self.merge_thread.join(timeout=5)
                
            self.is_merging = False
            self.update_progress("合併已停止\n")
            self.stop_merge_button.config(state=tk.DISABLED)

    def search_fire_report(self):
        """搜索并下载消防报表"""
        if self.is_downloading:
            return
        
        search_text = self.fire_search_var.get()
        if not search_text:
            messagebox.showwarning("警告", "請輸入要搜尋的消防設備名稱")
            return
        
        # 取得日期
        dialog = DateInputDialog(self)
        if dialog.result:
            date = dialog.result
            
            # 設置下載狀態
            self.is_downloading = True
            self.should_stop = False
            self.pdf_handler.should_stop = False
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            
            # 使用執行緒來執行下載
            self.download_thread = threading.Thread(
                target=self._search_fire_task,
                args=(date, search_text)
            )
            self.download_thread.start()

    def _search_fire_task(self, date, search_text):
        """執行消防搜尋下載任務的執行緒"""
        try:
            self.reset_progress_bar()  # 重置進度條
            self.update_progress(f"開始搜尋消防設備: {search_text}\n")
            
            # 搜尋匹配的報表
            find = []
            for i in self.pdf_handler.open_data['Fire_Equipment']:
                if self.should_stop:
                    raise Exception("使用者取消下載")
                if re.search(search_text, i['name']):
                    find.append(i)
            
            if len(find) > 0:
                self.update_progress(f"消防 共有 {len(find)} 個PDF\n")
                self.update_progress("將下載以下報表:\n")
                for item in find:
                    self.update_progress(f"{item['name']}\n")
                
                # 初始化 current_file
                current_file = 0
                
                # 在下載之前建立資料夾
                self.pdf_handler.folder(os.path.join(self.pdf_handler.File_folder, self.pdf_handler.fire_folder))
                self.pdf_handler.folder(os.path.join(self.pdf_handler.File_folder, self.pdf_handler.fire_folder, date))
                
                # 下載找到的報表
                success_count = 0
                for item in find:
                    if self.should_stop:
                        raise Exception("使用者取消下載")
                        
                    name = item['name']
                    url = f'https://vghtpe-ue.httc.com.tw/Report6{item["api_1"]}{date}{item["api_2"]}'
                    output_path = os.path.join(
                        self.pdf_handler.File_folder,
                        self.pdf_handler.fire_folder,
                        date,
                        f"{name}.pdf"
                    )
                    
                    if self.pdf_handler.download_report(url, output_path, name):
                        success_count += 1
                    
                    current_file += 1
                    self.update_progress_bar(current_file, len(find))  # 更新進度條
                
                self.update_progress(f"搜尋下載完成，成功 {success_count}/{len(find)} 個檔案\n")
                
                # 開啟資料夾
                if not self.should_stop:
                    start_directory = os.path.join(
                        self.pdf_handler.File_folder,
                        self.pdf_handler.fire_folder,
                        date
                    )
                    self.pdf_handler.startfile(start_directory)
            else:
                self.update_progress(f"找不到符合的報表：{search_text}\n")
                
        except Exception as e:
            if str(e) == "使用者取消下載":
                self.update_progress("下載已被取消\n")
            else:
                self.update_progress(f"下載過程發生錯誤: {str(e)}\n")
        finally:
            self.is_downloading = False
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            self.reset_progress_bar()  # 重置進度條
            
            # 重置停止標誌
            self.should_stop = False
            self.pdf_handler.should_stop = False

    def download_selected_report(self):
        """下載選定報表的月份區間。"""
        if self.is_downloading:
            return

        selection = self.report_list.curselection()
        if not selection:
            messagebox.showwarning("警告", "請先選擇要下載的報表")
            return
        
        selected_name = self.report_list.get(selection[0])
        selected_type = self.report_type_var.get()

        try:
            months = get_month_range(
                self.report_start_month_var.get().strip(),
                self.report_end_month_var.get().strip()
            )
        except ValueError as e:
            messagebox.showwarning("警告", str(e))
            return

        self.is_downloading = True
        self.should_stop = False
        self.pdf_handler.should_stop = False
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

        self.download_thread = threading.Thread(
            target=self._download_selected_report_task,
            args=(selected_name, selected_type, months)
        )
        self.download_thread.start()

    def _download_selected_report_task(self, selected_name, selected_type, months):
        """執行選定報表的跨月份下載任務。"""
        try:
            data_key = REPORT_TYPE_MAP.get(selected_type)
            reports = self.pdf_handler.open_data.get(data_key, [])
            report = next((item for item in reports if item['name'] == selected_name), None)
            if report is None:
                raise ValueError("找不到選取的報表資料")

            category_folder_name = (
                self.pdf_handler.fire_folder if selected_type == "消防"
                else self.pdf_handler.electricity_folder if selected_type.startswith("電力")
                else self.pdf_handler.drain_folder
            )
            category_folder = os.path.join(
                self.pdf_handler.File_folder,
                category_folder_name
            )
            self.pdf_handler.folder(category_folder)

            self.reset_progress_bar()
            self.update_progress(
                f"開始下載 {selected_name}，共 {len(months)} 個月份...\n"
            )

            success_count = 0
            failed_count = 0
            for index, month in enumerate(months, start=1):
                if self.should_stop:
                    raise Exception("使用者取消下載")

                output_folder = os.path.join(category_folder, month)
                self.pdf_handler.folder(output_folder)
                url = build_report_url(report, selected_type, month)
                output_path = os.path.join(output_folder, f"{selected_name}.pdf")

                if self.pdf_handler.download_report(url, output_path, selected_name):
                    success_count += 1
                else:
                    failed_count += 1
                self.update_progress_bar(index, len(months))

            self.update_progress(
                f"跨月份下載完成：成功 {success_count}、失敗 {failed_count}、"
                f"共 {len(months)} 個月份\n"
            )
            self.pdf_handler.startfile(category_folder)
        except Exception as e:
            if str(e) == "使用者取消下載":
                self.update_progress("下載已被取消\n")
            else:
                self.update_progress(f"下載過程發生錯誤：{str(e)}\n")
        finally:
            self.is_downloading = False
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            self.should_stop = False
            self.pdf_handler.should_stop = False

    def update_progress_bar(self, current, total):
        """設定進度目標，由主執行緒平滑更新顯示。"""
        if total <= 0:
            self.progress_target = 0.0
            return
        self.progress_target = max(0.0, min(100.0, (current / total) * 100))

    def _animate_progress_bar(self):
        """定時平滑更新進度條。"""
        self.progress_value = get_smooth_progress_value(
            self.progress_value,
            self.progress_target
        )
        self.progress_bar["value"] = self.progress_value
        if self.progress_value > 0 or self.progress_target > 0:
            self.progress_label.config(text=f"下載進度: {self.progress_value:.1f}%")
        else:
            self.progress_label.config(text="")
        self.after(30, self._animate_progress_bar)

    def reset_progress_bar(self):
        """重置進度條"""
        self.progress_value = 0.0
        self.progress_target = 0.0

    def clear_progress(self):
        """清除進度顯示區域的內容"""
        self.progress_text.delete('1.0', tk.END)  # 清除文字區域
        self.reset_progress_bar()  # 重置進度條
        self.update_idletasks()  # 更新界面

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
        return get_last_month()
        
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
        return validate_month(date)

if __name__ == '__main__':
    app = HtmlToPdfGUI()
    app.mainloop()

