import json
import pdfkit
import os

"""
目前先想功能就好，數據慢慢手動建立

打包後，要跳出彈出畫面 輸入年月份，只能一種格式，寫錯跳錯誤
輸出完成跳出資料夾

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
    # 使用相對路徑且資料夾都在根目錄 統一輸入一個值就可以解決
        folderpath = folderpath_name 
        # 檢查目錄是否存在 
        if os.path.isdir(folderpath):
            print('{} 資料夾存在。'.format(folderpath))
        else:
            print('資料夾不存在。建立{}資料夾'.format(folderpath))
            os.mkdir(folderpath)
            print('{}建立完成'.format(folderpath))

    def call(self, date:str):
        with open("testdata.json", encoding="utf-8") as f: #Test usefile testdata.json
            p = json.load(f) # json data

        date = date # ex : '2022-05'

        # 寫一個匿名funtion 計算多少筆檔案，抓name(使用len)

        for i in p['Fire_Equipment']:
            name = i['name']
            url_api_1 = i['api_1']
            url_api_2 = i['api_2']
            print(f'http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}' + f'  {name}')
            url = f'http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}'
            pdfkit.from_url(url, os.path.join("test", name) + '.pdf', options=htmltopdf.options, configuration=htmltopdf.config)

if __name__ == '__main__':
    pass
    # call('2022-05')
    htmltopdf().call('2022-05')
    # pdfkit.from_url('http://210.61.217.104/Report6/82/2022-05/82/98', os.path.join("test", 'abc11') + '.pdf', options=options, configuration=config)
