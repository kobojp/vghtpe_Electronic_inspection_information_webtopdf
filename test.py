# import os
# import subprocess
 
 
# # 利用subprocess
# def startfile(filename):
#     try:
#         os.startfile(filename)
#     except:
#         subprocess.Popen(['xdg-open', filename])
 
 
# start_directory = r'C:\my_code\vghtpe_Electronic_inspection_information_webtopdf\test'


# startfile(start_directory)

# from tkinter import ttk
# import tkinter as tk
 
# app = tk.Tk() 
# app.geometry('200x100')

# labelTop = tk.Label(app,
#                     text = "Choose your favourite month")
# labelTop.grid(column=0, row=0)

# comboExample = ttk.Combobox(app, 
#                             values=[
#                                     "January", 
#                                     "February",
#                                     "March",
#                                     "April"])
# print(dict(comboExample)) 
# comboExample.grid(column=0, row=1)
# comboExample.current(1)

# print(comboExample.current(), comboExample.get())

# app.mainloop()

# win = tk.Tk()
# win.title("Python GUI")    # 添加标题

# print(win.winfo_screenwidth()) #輸出螢幕寬度
# print(win.winfo_screenheight()) #輸出螢幕高度
# w=800  #width
# r=300  #height
# x=200  #與視窗左上x的距離
# y=300  #與視窗左上y的距離
# win.geometry('%dx%d+%d+%d' % (w,r,x,y))
 
# ttk.Label(win, text="Chooes a number").grid(column=1, row=0)    # 添加一个标签，并将其列设置为1，行设置为0
# ttk.Label(win, text="Enter a name:").grid(column=0, row=0)      # 设置其在界面中出现的位置  column代表列   row 代表行
 
# # button被点击之后会被执行
# def clickMe():   # 当acction被点击时,该函数则生效
#   action.configure(text='Hello ' + name.get()+ ' ' + numberChosen.get())     # 设置button显示的内容
#   action.configure(state='disabled')      # 将按钮设置为灰色状态，不可使用状态
 
# # 按钮
# action = ttk.Button(win, text="Click Me!", command=clickMe)     # 创建一个按钮, text：显示按钮上面显示的文字, command：当这个按钮被点击之后会调用command函数
# action.grid(column=2, row=1)    # 设置其在界面中出现的位置  column代表行   row 代表列
 
# # 文本框
# name = tk.StringVar()     # StringVar是Tk库内部定义的字符串变量类型，在这里用于管理部件上面的字符；不过一般用在按钮button上。改变StringVar，按钮上的文字也随之改变。
# nameEntered = ttk.Entry(win, width=12, textvariable=name)   # 创建一个文本框，定义长度为12个字符长度，并且将文本框中的内容绑定到上一句定义的name变量上，方便clickMe调用
# nameEntered.grid(column=0, row=1)       # 设置其在界面中出现的位置  column代表列   row 代表行
# nameEntered.focus()     # 当程序运行时,光标默认会出现在该文本框中
 
# # 创建一个下拉列表
# number = tk.StringVar()
# numberChosen = ttk.Combobox(win, width=12, textvariable=number)
# numberChosen['values'] = (1, 2, 4, 42, 100)     # 设置下拉列表的值
# numberChosen.grid(column=1, row=1)      # 设置其在界面中出现的位置行  column代表   row 代表列
# numberChosen.current(0)    # 设置下拉列表默认显示的值，0为 numberChosen['values'] 的下标值


# win.mainloop()      # 当调用mainloop()时,窗口才会显示出来


from tkinter import *

# print(到視窗)
# def check_expression():
#     #Your code that checks the expression
#     varContent = inputentry.get() # get what's written in the inputentry entry widget
#     outputtext.delete('0', 'end-1c') # clear the outputtext text widget
#     outputtext.insert(varContent)

# root = Tk()
# root.title("Post-fix solver")
# root.geometry("500x500")

# mainframe = Frame(root)
# mainframe.grid(column=0, row=0)

# inputentry = Entry(mainframe)
# inputentry.grid(column=1, row=1)

# executebutton = Button(mainframe, text="Run", command=check_expression)
# executebutton.grid(column=1, row=5)              

# outputtext = Text(mainframe)
# outputtext.grid(column=1, row=5)

# root.mainloop()

# from tkinter import *
# from tkinter import ttk
# root = Tk()
# frm = ttk.Frame(root, padding=10)
# frm.grid()
# ttk.Label(frm, text="Hello World!").grid(column=0, row=0)
# ttk.Button(frm, text="Quit", command=root.destroy).grid(column=1, row=0)
# root.mainloop()


# 以這個版本去修改，日期部分選擇值後取得選擇值，按下按鈕執行 爬取函式，將print結果丟到日誌窗

import hashlib
import time

LOG_LINE_NUM = 0

class MY_GUI():
    def __init__(self,init_window_name):
        self.init_window_name = init_window_name


    #设置窗口
    def set_init_window(self):
        self.init_window_name.title("下載消防報表_v1.0")           #窗口名
        #self.init_window_name.geometry('320x160+10+10')                         #290 160为窗口大小，+10 +10 定义窗口弹出时的默认展示位置
        self.init_window_name.geometry('800x600+10+10')
        #self.init_window_name["bg"] = "pink"                                    #窗口背景色，其他背景色见：blog.csdn.net/chl0000/article/details/7657887
        #self.init_window_name.attributes("-alpha",0.9)                          #虚化，值越小虚化程度越高
        #标签
        # self.init_data_label = Label(self.init_window_name, text="待处理数据")
        # self.init_data_label.grid(row=0, column=0)
        # self.result_data_label = Label(self.init_window_name, text="输出结果")
        # self.result_data_label.grid(row=0, column=12)
        self.log_label = Label(self.init_window_name, text="日志")
        self.log_label.grid(row=2, column=5)
        self.select = Label(self.init_window_name, text="選擇框")
        self.select.grid(row=2, column=1)


        #文本框
        # self.init_data_Text = Text(self.init_window_name, width=67, height=35)  #原始数据录入框
        # self.init_data_Text.grid(row=1, column=0, rowspan=10, columnspan=10)


        # self.result_data_Text = Text(self.init_window_name, width=70, height=49)  #处理结果展示
        # self.result_data_Text.grid(row=1, column=12, rowspan=15, columnspan=10)
        self.log_data_Text = Text(self.init_window_name, width=60, height=30)  # 日志框
        self.log_data_Text.grid(row=3, column=10, columnspan=1)

        #按钮
        # self.str_trans_to_md5_button = Button(self.init_window_name, text="字符串转MD5", bg="lightblue", width=10,command=self.str_trans_to_md5)  # 调用内部方法  加()为直接调用
        # self.str_trans_to_md5_button.grid(row=1, column=11)
        
        #選擇框
        self.select_option = Listbox(self.init_window_name,width=30, height=25)
        self.select_option.grid(row=3, column=2, columnspan=1)
        
        #教學 https://shengyu7697.github.io/python-tkinter-listbox/
        #取得選擇框選擇的值 
        self.select_optionget = Button(self.init_window_name, text='get current selection', command=self.button_event)
        self.select_optionget.grid(row=20, column=0, columnspan=10)
        # 顯示清單
        # https://www.geeksforgeeks.org/how-to-get-selected-value-from-listbox-in-tkinter/
        # 取得選擇值寫法

        li = ['C','python','php','html','SQL','java']
        for item in li:            # 第一个小部件插入数据
            self.select_option.insert(0,item) 
        
            # 打印到text文字框，要寫一個def 執行程式就會打印print
            # self.init_data_Text.insert(END, item + '\n')

        # self.select_option.pack()

    # 回傳選擇框的值
    def button_event(self):
        # print(type(mylistbox.curselection()))
        print(self.select_option.curselection())
        # print(self.select_option.get())
        for i in self.select_option.curselection():
            print(self.select_option.get(i))
            test_1 = self.select_option.get(i)
            #這裡下call程式
        

    #功能函数
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
    ZMJ_PORTAL = MY_GUI(init_window)
    # 设置根窗口默认属性
    ZMJ_PORTAL.set_init_window()

    init_window.mainloop()          #父窗口进入事件循环，可以理解为保持窗口运行，否则界面不展示


gui_start()


