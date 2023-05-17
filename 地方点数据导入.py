import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import pymysql
import pyodbc
import xlwt


def gui_window():
    window = tk.Tk()
    window.title("GUI程序")

    # 创建年份和月份选项数组
    yearList = []
    for i in range(2023, 2029):
        yearList.append(str(i))
    monthList = []
    for i in range(1, 13):
        if len(str(i)) == 1:
            monthList.append('0' + str(i))
        else:
            monthList.append(str(i))

    # 创建年份和月份选择框
    yearVar = tk.StringVar(window)
    monthVar = tk.StringVar(window)

    yearLabel = tk.Label(window, text="年份：")
    yearLabel.grid(column=0, row=0, padx=5, pady=5)
    yearOption = tk.OptionMenu(window, yearVar, *yearList)
    yearOption.grid(column=1, row=0, padx=5, pady=5)

    monthLabel = tk.Label(window, text="月份：")
    monthLabel.grid(column=2, row=0, padx=5, pady=5)
    monthOption = tk.OptionMenu(window, monthVar, *monthList)
    monthOption.grid(column=3, row=0, padx=5, pady=5)

    # 创建导入Excel、数据清理和导出汇总表按钮
    def import_excel():
        # TODO: 实现Excel导入功能
        pass

    def clean_data():
        # TODO: 实现数据清理功能
        pass

    def export_summary_table():
        # TODO: 实现导出汇总表功能
        pass

    importButton = tk.Button(window, text="导入Excel", command=import_excel)
    importButton.grid(column=0, row=1, padx=5, pady=5)

    cleanButton = tk.Button(window, text="数据清理", command=clean_data)
    cleanButton.grid(column=1, row=1, padx=5, pady=5)

    exportButton = tk.Button(window, text="导出汇总表", command=export_summary_table)
    exportButton.grid(column=2, row=1, padx=5, pady=5)

    window.mainloop()
