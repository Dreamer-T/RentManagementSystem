
import datetime
from functools import partial
import tkinter as tk
from tkinter import ttk
from tkintertable import TableModel, TableCanvas
from DataProcess import *
from Button import *


default_excel_filepath = "合同表格copy.xlsx"
default_exceldict = excel_to_dict(default_excel_filepath)
# print(default_table)
default_model = TableModel()
default_model.importDict(default_exceldict)

initial_button(default_exceldict)


root = tk.Tk()
root.title("Main Window")
root.geometry("800x600")
table_frame = tk.Frame(root)
table_frame.pack()
datadict = {}
table = TableCanvas(table_frame, model=default_model)
table.show()


# anthor section
function_frame = tk.Frame(root)
function_frame.pack()

open_file_button = tk.Button(
    function_frame, text="选择文件",  command=partial(choose_excel, table))
open_file_button.grid(row=0, column=2)
space_label0 = tk.Label(function_frame, width=10)
space_label0.grid(row=1, column=3)

# default query month will be today's mont
month_variable = tk.StringVar()
current_date = datetime.date.today()
month_variable.set(current_date.month)
months = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
# print(months)
space_label1 = tk.Label(function_frame, width=10)
space_label1.grid(row=3, column=2)

month_label = tk.Label(function_frame, text="月份选择")
month_label.grid(row=4, column=1)
month_combo = ttk.Combobox(
    function_frame, textvariable=month_variable, values=months, width=10)
month_combo.grid(row=4, column=2)

space_label2 = tk.Label(function_frame, width=10)
space_label2.grid(row=5, column=2)

pay_way_variable = tk.StringVar()
pay_way_variable.set("季度")

pay_ways = ["季度", "半年", "年度", "所有"]
pay_way_label = tk.Label(function_frame, text="付款方式")
pay_way_label.grid(row=6, column=1)
pay_way_combo = ttk.Combobox(
    function_frame, textvariable=pay_way_variable, values=pay_ways, width=10)
pay_way_combo.grid(row=6, column=2, sticky="NSEW")
query_button = tk.Button(function_frame, text="查询", command=partial(
    query, month_combo, pay_way_combo))
query_button.grid(row=9, column=2)

space_label3 = tk.Label(function_frame, width=10)
space_label3.grid(row=7, column=3)
change_button = tk.Button(function_frame, text="修改",
                          command=modify)
change_button.grid(row=8, column=2)

root.mainloop()
