import tkinter as tk
from tkinter import ttk
import pyodbc
import json
import pandas as pd

# 读取mysql.json文件中的参数
with open("mysql.json", "r") as f:
    mysql_params = json.load(f)

# 连接到SQL Server数据库
conn = pyodbc.connect("DRIVER={ODBC Driver 18 for SQL Server};SERVER=%s;DATABASE=%s;ENCRYPT=yes;TrustServerCertificate=yes;UID=%s;PWD=%s"
                      % (mysql_params["server"], 
                         mysql_params["database"], mysql_params["user"], mysql_params["password"]))

# 创建GUI窗口
root = tk.Tk()
root.title("畜禽月季报处理程序")

# 月份下拉选择框
month_var = tk.StringVar(value="01")
month_label = ttk.Label(root, text="月份:")
month_label.grid(row=0, column=0, padx=5, pady=5)
month_combobox = ttk.Combobox(root, textvariable=month_var, values=["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"])
month_combobox.grid(row=0, column=1, padx=5, pady=5)

# “初始化本月数据库”按钮
def init_db():
    month = month_var.get()
    symonth=str(int(month)-1)
    if len(symonth) <2:
        symonth='0'+symonth
        
    cursor = conn.cursor()
    csh1='''insert into [dbo].[畜禽调查表]
SELECT  [户编码] as [代码]
      ,[名称]
      ,0 as [指标1]
      ,0 as [指标2]
      ,0 as [指标3]
      ,0 as [指标4]
      ,0 as [指标5]
      ,0 as [指标6]
      ,0 as [指标7]
      ,0 as [指标8]
      ,0 as [指标9]
      ,0 as [指标10]
      ,0 as [指标11]
      ,0 as [指标12]
      ,0 as [指标13]
      ,0 as [指标14]
      ,0 as [指标15]
      ,0 as [指标16]
      ,0 as [指标17]
      ,0 as [指标18]
      ,0 as [指标19]
      ,0 as [指标20]
      ,0 as [指标21]
      ,0 as [指标22]
      ,0 as [指标23]
      ,'''+'\'' +month+'\''+''' as [yf]
      , [养殖类型] as [yzlx]
  FROM [xqdc].[dbo].[国家名录库] 
  where [养殖类型]='1\''''
    csh2='''UPDATE  畜禽调查表
SET         指标19 = 畜禽调查表_1.指标1
FROM      畜禽调查表 INNER JOIN
                畜禽调查表 AS 畜禽调查表_1 ON 畜禽调查表.代码 = 畜禽调查表_1.代码
WHERE   (畜禽调查表.yf = '''+'\'' +month+'\''+''') AND (畜禽调查表_1.yf = '''+'\'' +symonth+'\''+''')
    '''
    csh3='''CREATE VIEW 审核公式视图 AS
SELECT t1.代码, t1.名称, t1.yzlx,
CASE
WHEN t1.指标1 <> t1.指标2 + t1.指标3 + t1.指标5 THEN '期末存栏<>仔猪+待育肥猪+种猪'
WHEN t1.指标3 < t1.指标4 THEN '待育肥猪应大于等于50公斤以上'
WHEN t1.指标5 < t1.指标6 THEN '种猪应大于等于能繁母猪'
WHEN t1.指标7 < t1.指标8 + t1.指标9 THEN '期内增加应大于等于自繁加购进'
WHEN t1.指标10 <> t1.指标11 + t1.指标12 + t1.指标15 THEN '期内减少应等于自宰加出售头数加其他原因减少'
WHEN t1.指标15 < t1.指标16 THEN '其他原因减少应大于等于出售仔猪数量'
WHEN (t1.指标14 / NULLIF(t1.指标12, 0) > 200 OR t1.指标14 / NULLIF(t1.指标12, 0) < 70) AND t1.指标12 <> 0 THEN '检查肥猪重量是否在合理范围'
WHEN (t1.指标13 / NULLIF(t1.指标14, 0) > 20 OR t1.指标13 / NULLIF(t1.指标14, 0) < 8) AND t1.指标14 <> 0 THEN '检查出售单价是否在合理范围'
WHEN t1.指标1 <> t2.指标1 + t1.指标7 - t1.指标10 THEN '期末存栏不等于上期存栏加本期增加减本期减少'
WHEN (t1.指标12 + t1.指标11 > t2.指标4) THEN '本期自宰加出栏不应大于上期50公斤以上待育肥猪'
END AS 错误类型
FROM 畜禽调查表 t1
LEFT JOIN 畜禽调查表 t2 ON t1.代码 = t2.代码
AND t1.yf = '''+'\'' +month+'\''+''' AND t1.yzlx = N'1'
AND t2.yf = '''+'\'' +symonth+'\''+''' AND t2.yzlx = N'1'
WHERE t1.yf = '''+'\'' +month+'\''+''' AND t1.yzlx = N'1'
    '''
    csh4='drop view 审核公式视图'
    print(csh1)
    print(csh2)
    print(csh3)
    cursor.execute(csh1)
    cursor.execute(csh2)
    cursor.execute(csh4)
    cursor.execute(csh3)
    cursor.commit()
    cursor.close()

init_btn = ttk.Button(root, text="初始化本月数据库", command=init_db)
init_btn.grid(row=1, column=0, padx=5, pady=5)

# “查询错误信息”按钮
def query_error():
    month = month_var.get()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM mytable WHERE month=?", month)
    rows = cursor.fetchall()
    cursor.close()
    if len(rows) > 0:
        df = pd.DataFrame(rows, columns=["id", "month"])
        df.to_excel("错误信息%s.xlsx" % month, index=False)
        tk.messagebox.showinfo("查询结果", "查询成功，结果已保存到文件中。")
    else:
        tk.messagebox.showwarning("查询结果", "未找到相关记录。")

error_btn = ttk.Button(root, text="查询错误信息", command=query_error)
error_btn.grid(row=1, column=1, padx=5, pady=5)

# “输出过录表”按钮
def gen_report():
    month = month_var.get()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM mytable WHERE month=?", month)
    rows = cursor.fetchall()
    cursor.close()
    if len(rows) > 0:
        df = pd.DataFrame(rows, columns=["id", "month"])
        df.to_excel("过录表%s.xlsx" % month, index=False)
        tk.messagebox.showinfo("输出结果", "输出成功，结果已保存到文件中。")
    else:
        tk.messagebox.showwarning("输出结果", "未找到相关记录。")

report_btn = ttk.Button(root, text="输出过录表", command=gen_report)
report_btn.grid(row=1, column=2, padx=5, pady=5)

root.mainloop()
