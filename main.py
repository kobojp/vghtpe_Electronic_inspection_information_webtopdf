import json
import pdfkit
import os
import subprocess
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

        with open("testdata.json", encoding="utf-8") as f: #Test usefile testdata.json
            p = json.load(f) # json data

        date = date # ex : '2022-05'

        # 寫一個匿名funtion 計算多少筆檔案，抓name(使用len)
        # 多現程 分配 3個線程資料分成三等份，計算處理筆數 
        for i in p['Fire_Equipment']:
            name = i['name']
            url_api_1 = i['api_1']
            url_api_2 = i['api_2']
            print(f'http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}' + f'  {name}')
            url = f'http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}'
            pdfkit.from_url(url, os.path.join("test", date, name) + '.pdf', options=htmltopdf.options, configuration=htmltopdf.config)
        
        # open folder
        start_directory = os.path.join("test", date) 
        self.startfile(start_directory)

    # Finish Open the folder
    def startfile(sele, filename):
        try:
            os.startfile(filename)
        except:
            subprocess.Popen(['xdg-open', filename])


if __name__ == '__main__':
    pass
    # call('2022-05')
    htmltopdf().call('2022-03')
    # pdfkit.from_url('http://210.61.217.104/Report6/82/2022-05/82/98', os.path.join("test", 'abc11') + '.pdf', options=options, configuration=config)
