from tkinter import *
from tkinter.ttk import *
import pymssql
import pandas as pd
# 定义数据库服务器连接

servername='10.132.60.188'
username='sa'
mypassword='zpf13062990859'
mydatabase='hhgzhdc'
conn = pymssql.connect(server=servername, user=username, password=mypassword, database=mydatabase,charset="utf8") 
cursor=conn.cursor
# 定义时间范围
yearitems = ("2022", "2023", "2024", "2025")
monthitems = ("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")

# 退出子程序
def tuichu():
    conn.close() 
    root.destroy()


def shengchen():
    if conn:
        sql='select * from 合并上报表' + ' where year=' +yearchoice +' and month='+monthchoice
        cursor.execute(sql)
        fields = [field[0] for field in cursor.description]
        all_data = cursor.fetchall()

    
# 生成窗口
root = Tk()
root.title("住户地方点数据汇总系统")
root.geometry("1024x768")
root.resizable(TRUE, TRUE)
    
# 生成选择时间按钮
lable1 = Label(root, text='请选择时间')
lable1.grid(row=0, column=0)

yearchoice = Combobox(root)
yearchoice.grid(row=0, column=1)
yearchoice['values'] = yearitems

lable2 = Label(root, text='年')
lable2.grid(row=0, column=2)

monthchoice = Combobox(root)
monthchoice.grid(row=0, column=3)
monthchoice['values'] = monthitems

lable3 = Label(root, text='月')
lable3.grid(row=0, column=4)

sc_botton1=Button(root,text='生成excel表格',command=shengchen)
sc_botton1.grid(row=4, column=0)
sc_botton2=Button(root,text='退出',command=tuichu)
sc_botton2.grid(row=4, column=1)

# 定义生成子程序
root.mainloop()