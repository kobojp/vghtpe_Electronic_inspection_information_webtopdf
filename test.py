import json
import pdfkit
import sys
from PyQt5 import QtWidgets, QtWebEngineWidgets
from PyQt5.QtCore import QUrl
from PyQt5.QtGui import QPageLayout, QPageSize
from PyQt5.QtWidgets import QApplication

def topdf(url,file_name):
    app = QtWidgets.QApplication(sys.argv)
    loader = QtWebEngineWidgets.QWebEngineView()
    loader.setZoomFactor(1)
    layout = QPageLayout()
    layout.setPageSize(QPageSize(QPageSize.A4Extra))
    layout.setOrientation(QPageLayout.Portrait)
    loader.load(QUrl(url))
    loader.page().pdfPrintingFinished.connect(lambda *args: QApplication.exit())

    def emit_pdf(finished):
        loader.page().printToPdf(file_name +'.pdf', pageLayout=layout)

    loader.loadFinished.connect(emit_pdf)
    sys.exit(app.exec_())

    

with open("data.json", encoding="utf-8") as f:

    # 讀取 JSON 檔案
    p = json.load(f)
    # print(p)

    # 查看整個 JSON 資料解析後的結果
    # print("p =", p)
    # print("p =", p['members'][0]['name'])

date = '2022-05'
for i in p['Fire_Equipment']:
    s = i['name']
    b = i['api_1']
    c = i['api_2']
    print(f'http://210.61.217.104/Report6{b}{date}{c}' + f'  {s}')
    url = f'http://210.61.217.104/Report6{b}{date}{c}'

    pdfkit.from_url(url, s + '.pdf')
    # topdf(url,s)

# pdfkit.from_url('http://210.61.217.104/Report6/86/2022-05/86/102', 'outAA.pdf')
        # print(s,b,c)
    
    # # 取得 name 的值
    # print("name =", p["name"])

    # # 取得 skill 的值
    # print("skill =", p["skill"])