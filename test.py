import datetime
from functools import partial
import tkinter as tk
from tkinter import ttk
from tkintertable import TableModel, TableCanvas
from DataProcess import *
from Button import *


default_excel_filepath = "合同表格 copy.xlsx"
default_exceldict = excel_to_dict(default_excel_filepath)
# print(default_table)
default_model = TableModel()
default_model.importDict(default_exceldict)


def choose_excel(tablecanvas: TableCanvas, dict_variable: tk.Variable, default_excel_filepath="合同表格 copy.xlsx"):
    """
    open a dialog to let the user choose the excel file they want to open

    tablecanvas: the canvas to show the result
    dict_variable: a variable to store the data dict
    """
    excel_file = filedialog.askopenfilename()
    if excel_file == "":
        excel_file = default_excel_filepath
    excel_dict = excel_to_dict(excel_file)
    dict_variable.set(excel_dict)
    excel_model = TableModel()
    excel_model.importDict(excel_dict)
    tablecanvas.setModel(excel_model)


def query(excel_dict: dict, month: str, pay_way: str):
    print(excel_dict.keys())
    print(month)
    print(pay_way)
    print("ask for query")


root = tk.Tk()
root.title("Main Window")
root.geometry("800x600")

# 创建表格框架
table_frame = tk.Frame(root)
table_frame.pack()

# 创建表格并显示
datadict = {}
default_model = TableModel()
default_model.importDict(default_exceldict)
table = TableCanvas(table_frame, model=default_model)
table.show()

# 创建选择文件按钮
dict_variable = tk.Variable()
dict_variable.set(default_exceldict)
open_file_button = tk.Button(root, text="选择文件", command=partial(
    choose_excel, table, dict_variable))
open_file_button.pack()

# 创建功能框架
function_frame = tk.Frame(root)
function_frame.pack()

# 创建月份选择组合框
month_variable = tk.StringVar()
current_date = datetime.date.today()
month_variable.set(current_date.month)
months = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
month_label = tk.Label(function_frame, text="月份选择")
month_label.grid(row=0, column=0, sticky="w")
month_combo = ttk.Combobox(
    function_frame, textvariable=month_variable, values=months, width=10)
month_combo.grid(row=0, column=1, sticky="w")

# 创建付款方式选择组合框
pay_way_variable = tk.StringVar()
pay_way_variable.set("季度")
pay_ways = ["季度", "半年", "年度"]
pay_way_label = tk.Label(function_frame, text="付款方式")
pay_way_label.grid(row=0, column=2, sticky="w")
pay_way_combo = ttk.Combobox(
    function_frame, textvariable=pay_way_variable, values=pay_ways, width=10)
pay_way_combo.grid(row=0, column=3, sticky="w")

# 创建查询按钮
query_button = tk.Button(function_frame, text="查询", command=partial(
    query, dict_variable.get(), month_variable.get(), pay_way_variable.get()))
query_button.grid(row=0, column=4, sticky="w")

root.mainloop()
