import json
import pdfkit

"""
新增功能
    當前路徑建立資料夾，判斷資料夾是否存在，以月分建檔名，5月消防月報表

colab 上處理，全部轉換後並一次上傳到drive，使用rclone工具

高階一點的寫法
    5個線程處理

wkhtmltopdf doc
 https://wkhtmltopdf.org/usage/wkhtmltopdf.txt

套件 
 https://pypi.org/project/pdfkit/
"""
options = {
    'no-background': None
}

def call(date:str):
    with open("data.json", encoding="utf-8") as f:
        p = json.load(f) # json data

    date = date # ex : '2022-05'
    for i in p['Fire_Equipment']:
        name = i['name']
        url_api_1 = i['api_1']
        url_api_2 = i['api_2']
        print(f'http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}' + f'  {name}')
        url = f'http://210.61.217.104/Report6{url_api_1}{date}{url_api_2}'
        pdfkit.from_url(url, name + '.pdf', options=options)

if __name__ == '__main__':

    call('2022-05')