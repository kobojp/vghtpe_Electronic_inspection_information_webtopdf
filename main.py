import json
import pdfkit
import os
import subprocess
from tkinter import *
import hashlib
import time
import main

# import requests

"""
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
    def __init__(self,init_window_name):
        self.init_window_name = init_window_name
    # json data    
    file_data = 'data.json' #Test usefile test_data.json
    # wkhtmltopdf Path setting
    path_wkhtmltopdf = os.path.join(os.getcwd(), 'wkhtmltox\\bin', 'wkhtmltopdf.exe')
    config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
    # PDF output settings
    options = {
    'no-background': None
    }

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

    def call(self, date:str):
        # url json https://www.delftstack.com/zh-tw/howto/python/python-get-json-from-url/
        
        self.folder(os.path.join("test", date))  #建立資料夾

        with open(self.file_data, encoding="utf-8") as f: #Test usefile test_data.json
            p = json.load(f) # json data

        date = date # ex : '2022-05'

        #在這裡寫判斷值 抓取當前月份 如果選擇大於月份就跳出錯誤訊息

        # 寫一個匿名funtion 計算多少筆檔案，抓name(使用len)
        try:
            for i in p['Fire_Equipment']:
                name = i['name']
                url_api_1 = i['api_1']
                url_api_2 = i['api_2']
                message = f'http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}' + f'  {name}'
                
                print(f'http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}' + f'  {name}')
                url = f'http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}'
                pdfkit.from_url(url, os.path.join("test", date, name) + '.pdf', options=htmltopdf.options, configuration=htmltopdf.config)
                self.log_data_Text.insert(END, message + '\n')
                self.log_data_Text.update() #動態更新訊息
                # 寫一個大於10筆就刪除訊息資料
                
                # return f'http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}' + f'  {name}'
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


# gui 設計區塊

    #设置窗口
    def set_init_window(self):
        self.init_window_name.title("下載消防報表_v1.0")           #窗口名
        #self.init_window_name.geometry('320x160+10+10')                         #290 160为窗口大小，+10 +10 定义窗口弹出时的默认展示位置
        self.init_window_name.geometry('800x600+10+10')
        #self.init_window_name["bg"] = "pink"                                    #窗口背景色，其他背景色见：blog.csdn.net/chl0000/article/details/7657887
        #self.init_window_name.attributes("-alpha",0.9)                          #虚化，值越小虚化程度越高
        #标签
        self.log_label = Label(self.init_window_name, text="日志")
        self.log_label.grid(row=2, column=5)
        self.select = Label(self.init_window_name, text="日期\n選擇")
        self.select.grid(row=3, column=1)


        #框
        self.log_data_Text = Text(self.init_window_name, width=60, height=30)  # 日志框
        self.log_data_Text.grid(row=3, column=10, columnspan=1)

        #按钮
        # self.str_trans_to_md5_button = Button(self.init_window_name, text="字符串转MD5", bg="lightblue", width=10,command=self.str_trans_to_md5)  # 调用内部方法  加()为直接调用
        # self.str_trans_to_md5_button.grid(row=1, column=11)
        
        #日期框
        self.select_option = Listbox(self.init_window_name,width=30, height=15)
        self.select_option.grid(row=3, column=2, columnspan=1)
        
        #教學 https://shengyu7697.github.io/python-tkinter-listbox/
        #取得選擇框選擇的值 get按鈕
        self.select_optionget = Button(self.init_window_name, text='get current selection', command=self.button_event)
        self.select_optionget.grid(row=20, column=0, columnspan=10)
        # 顯示清單
        # https://www.geeksforgeeks.org/how-to-get-selected-value-from-listbox-in-tkinter/
        # 取得選擇值寫法

        # li = ['C','python','php','html','SQL','java']
        year = str(time.localtime().tm_year)
        li = [year + '-01',year + '-02',year + '-03',year + '-04',year + '-05',year + '-06',
        year + '-07',year + '-08',year + '-09',year + '-10',year + '-11',year + '-12']

        """
        取當前年份
        共12月份
        
        +寫判斷取得當前月份如果選擇大於當前月份就要跳出錯誤

        """
        for item in li:            # 插入li，放入日期選項
            self.select_option.insert(0,item) 
        
            # 打印到text文字框，要寫一個def 執行程式就會打印print
            # self.init_data_Text.insert(END, item + '\n')

        # self.select_option.pack()

    # 回傳日期的值
    def button_event(self):
        # print(type(mylistbox.curselection()))
        print(self.select_option.curselection())
        # print(self.select_option.get())
        
        for i in self.select_option.curselection():
            print(self.select_option.get(i))
            self.call(self.select_option.get(i))
            # self.test_code()
            #這裡下call程式
        

    #參考功能函数
    def str_trans_to_md5(self):

        self.test_code()
        src = self.init_data_Text.get(1.0,END).strip().replace("\n","").encode()
        #print("src =",src)
        if src:
            try:
                myMd5 = hashlib.md5()
                myMd5.update(src)
                myMd5_Digest = myMd5.hexdigest()
                #print(myMd5_Digest)
                #输出到界面
                self.result_data_Text.delete(1.0,END)
                self.result_data_Text.insert(1.0,myMd5_Digest)
                self.write_log_to_Text("INFO:str_trans_to_md5 success")
                # self.result_data_Text.insert(1.0, 'TEST') 測試寫到TEXT
            except:
                self.result_data_Text.delete(1.0,END)
                self.result_data_Text.insert(1.0,"字符串转MD5失败")
        else:
            self.write_log_to_Text("ERROR:str_trans_to_md5 failed")


    #获取当前时间
    def get_current_time(self):
        current_time = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
        return current_time


    #日志动态打印
    def write_log_to_Text(self,logmsg):
        global LOG_LINE_NUM
        current_time = self.get_current_time()
        logmsg_in = str(current_time) +" " + str(logmsg) + "\n"      #换行
        if LOG_LINE_NUM <= 7:
            self.log_data_Text.insert(END, logmsg_in)
            LOG_LINE_NUM = LOG_LINE_NUM + 1
        else:
            self.log_data_Text.delete(1.0,2.0)
            self.log_data_Text.insert(END, logmsg_in)

    # 測試數據到text框
    def test_code(self):
        # date = date
        for i in range(1,100):
            self.log_data_Text.insert(END, str(i) + '\n')


def gui_start():
    init_window = Tk()              #实例化出一个父窗口
    ZMJ_PORTAL = htmltopdf(init_window)
    # 设置根窗口默认属性
    ZMJ_PORTAL.set_init_window()

    init_window.mainloop()          #父窗口进入事件循环，可以理解为保持窗口运行，否则界面不展示


if __name__ == '__main__':
    # my = htmltopdf()
    # my.call('2022-05')
    gui_start()
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