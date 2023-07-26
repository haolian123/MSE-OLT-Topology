import tkinter as tk
from tkinter import scrolledtext
from tkinter import filedialog
import pandas as pd  # 引入 pandas
import os
from datetime import datetime

from TopGraph import TopologyDiagram as TD


class GUI:
    def __init__(self):
        # 创建一个窗口对象
        self.window = tk.Tk()

        # 设置窗口标题
        self.window.title('MSE-OLT')

        # 设置窗口大小和位置
        self.window.geometry('550x400+500+200')

        #设置窗口logo
        self.window.iconphoto(False, tk.PhotoImage(file='logo.png'))


        # 用于存储文件路径
        self.datefile_path = tk.StringVar()
        self.MSE_col_names=tk.StringVar()
        self.OLT_col_names=tk.StringVar()
        self.sheet_name=tk.StringVar()


        self.MSE_col_names.set("Z端设备中文名称")
        self.OLT_col_names.set("A端设备中文名称")
        self.sheet_name.set("3")
        # 创建一个标签，显示"数据文件："
        self.datafile_label = tk.Label(self.window, text="数据文件：", font=('Arial', 13))
        self.datafile_label.place(x=5, y=8)

        # 创建一个文本框，用于显示数据文件路径，并设置为只读状态
        self.datafile_entry = tk.Entry(self.window, textvariable=self.datefile_path, width=30, state='readonly')
        self.datafile_entry.place(x=100, y=8)

        # 创建一个按钮，点击后调用select_file函数，参数为1，用于选择数据文件
        # self.datafile_button = tk.Button(self.window, command=lambda: self.emailManager.select_file(1,self.datefile_path,self.show_text), text="选择数据文件")
        self.datafile_button = tk.Button(self.window, command=lambda: self.select_file(self.datefile_path,self.show_text), text="选择数据文件")
        self.datafile_button.place(x=350, y=8)


         #创建文本框,sheet_name序列号
        self.sheet_name_label = tk.Label(self.window, text="子表序号[1...n]", font=('Arial', 10))
        self.sheet_name_label.place(x=5, y=40)
        # 创建一个文本框
        self.sheet_name_entry = tk.Entry(self.window, textvariable=self.sheet_name, width=10, state='normal')
        self.sheet_name_entry.place(x=100, y=40)


        #创建文本框,MSE字段名
        self.MSE_label = tk.Label(self.window, text="MSE字段名：", font=('Arial', 10))
        self.MSE_label.place(x=5, y=65)
        # 创建一个文本框
        self.MSE_entry = tk.Entry(self.window, textvariable=self.MSE_col_names, width=30, state='normal')
        self.MSE_entry.place(x=100, y=65)

        #创建文本框,OLT字段名
        self.OLT_label = tk.Label(self.window, text="OLT字段名：", font=('Arial', 10))
        self.OLT_label.place(x=5, y=85)
        # 创建一个文本框
        self.OLT_entry = tk.Entry(self.window, textvariable=self.OLT_col_names, width=30, state='normal')
        self.OLT_entry.place(x=100, y=85)

        # 创建一个按钮，点击后调用thread_run函数，用于运行程序
        # self.exec_button = tk.Button(self.window, text='运行', command=self.thread_run, font=('Arial', 13))
        self.exec_button = tk.Button(self.window, command=self.run, text='运行', font=('Arial', 13))
        self.exec_button.place(x=460, y=8)


        # 创建一个文本框，用于显示程序运行结果，设置为不可编辑状态

        self.show_text = tk.Text(self.window, height=20, width=77, fg="#06EB00", bg='black', state='disabled', font=('', 10))
        self.show_text.place(x=2, y=120)


        # 创建一个滚动条
        self.scrollbar = tk.Scrollbar(self.window)

        # 设置滚动条与文本框的关联
        self.scrollbar.config(command=self.show_text.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.show_text.config(yscrollcommand=self.scrollbar.set)


        # 滚动到文本框末尾
        self.show_text.see('end')
    def show_mainmenu(self):
        welcome_string =''' 
注意:
    子表序号：待转化的数据子表在数据文件中的序号，编号从1开始。
    MSE字段名：数据子表中MSE设备的字段名(列名)。
    OLT字段名：数据子表中OLT设备的字段名(列名)。
[*] 请选择需要处理的数据文件和子表在数据文件里的序号，并运行。
'''
        # self.show_text.config(state='normal')
        self.insert_show_text(show_text=self.show_text,insert_text=welcome_string)
        # 进入窗口的主循环
        self.window.mainloop()

    #向文本框插入文本
    def insert_show_text(self,show_text,insert_text=''):
        show_text.config(state='normal')
        if insert_text[-1]!='\n':
            insert_text+='\n'
        show_text.insert('end',insert_text)
        show_text.config(state='disable')

    #运行按钮
    def run(self):
        self.insert_show_text(self.show_text,self.sheet_name.get())
        self.insert_show_text(self.show_text,self.MSE_col_names.get())
        self.insert_show_text(self.show_text,self.OLT_col_names.get())
        self.insert_show_text(self.show_text,self.datefile_path.get())
        TD.MSE_to_OLT(self.datefile_path.get(),sheet_name=int(self.sheet_name.get())-1,MSE_name=self.MSE_col_names.get(),OLT_name=self.OLT_col_names.get())




    #选择数据文件
    def select_file(self,filePath,show_text):
        #选择数据文件
        
        file_path = tk.filedialog.askopenfilename(title='请选择要处理的文件')
        filePath.set(file_path) 
        # show_text.config(state='normal')
        # show_text.insert('end', '[*] 选择的数据文件为：' + file_path + '\n')
        # show_text.config(state='disable')
        self.insert_show_text(self.show_text,'[*] 选择的数据文件为：' + file_path + '\n')
   
        