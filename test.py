from tkinter import *
import hashlib
import time
import main_UI

"""
將此檔案改為ui，並繼承到main再去call

"""

# 以這個版本去修改，日期部分選擇值後取得選擇值，按下按鈕執行 爬取函式，將print結果丟到日誌窗

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
            test_1 = self.select_option.get(i)
            main_UI.htmltopdf().call(self.select_option.get(i))
            self.test_code()
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


