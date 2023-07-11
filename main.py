import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter import filedialog
from functools import partial
import pandas as pd
from tkintertable import TableModel, TableCanvas
import datetime

# file path
excel_file = "合同表格 copy.xlsx"
data = pd.read_excel(excel_file)


def choose_excel(table: TableCanvas, datadict: dict):
    """
    This function will show the table on GUI, which is chosen from the user
    Note that the global variable data and excel_file will change

    table: where to show the result

    datadict: a dictionary which stores the data
    """
    global excel_file
    excel_file = filedialog.askopenfilename()
    global data
    data = read_data()
    if data.shape[1] == 12:
        tk.Message(text="Error")
    datadict = renew_data(data)
    table_model = TableModel()
    table_model.importDict(datadict)
    table.setModel(table_model)
    # print(data)


def read_data():
    global excel_file
    excel = pd.read_excel(excel_file)
    return excel


def renew_data(data: pd.DataFrame):
    datadict = {}
    # get the number of the line in the table
    linecount = data[data.keys()[0]].shape[0]
    for i in range(linecount):
        eachline = {}
        for j, col in enumerate(data.keys()):
            # print(col, data.iloc[i, j])
            if pd.isna(data.iloc[i, j]):
                eachline[col] = ""
            else:
                if j == 0 or j == 1 or j == 10:
                    eachline[col] = str(data.iloc[i, j])
                if j == 2 or j == 6 or j == 7 or j == 8 or j == 9:
                    eachline[col] = datetime.datetime.strftime(pd.to_datetime(data.iloc[i, j]),"%Y-%m-%d")
                if j == 3:
                    eachline[col] = int(data.iloc[i, j])
                if j == 4 or j == 5:
                    eachline[col] = float(data.iloc[i, j])
        datadict[str(i)] = eachline
    return datadict


def get_latest(time_list: list):
    if time_list.count == 0:
        return ""
    latest_time = time_list[0]
    for i in time_list:
        if i.datetime.year > latest_time.datetime.year:
            latest_time = i
        elif i.datetime.year == latest_time.datetime.year:
            if i.datetime.month > latest_time.datetime.month:
                latest_time = i
            elif i.datetime.month == latest_time.datetime.month:
                if i.datetime.day > latest_time.datetime.day:
                    latest_time = i
    return latest_time


def find_latest(datadict: dict):
    time_list = []
    for line in datadict:
        if datadict[line]["第一次缴费"] != "":
            time_list.append(datetime.datetime.strptime(datadict[line]['第一次缴费'],"%Y-%m-%d"))
        if datadict[line]['第二次缴费'] != "":
            time_list.append(datetime.datetime.strptime(datadict[line]['第二次缴费'],"%Y-%m-%d"))
        if datadict[line]['第三次缴费'] != "":
            time_list.append(datetime.datetime.strptime(datadict[line]['第三次缴费'],"%Y-%m-%d"))
        if datadict[line]['第四次缴费'] != "":
            time_list.append(datetime.datetime.strptime(datadict[line]['第四次缴费'],"%Y-%m-%d"))
    return get_latest(time_list)



def season_pay(datadict: dict,current_time:datetime.date):
    tempData = {}
    for line in datadict:
        if datadict[line]["付款方式"] == 0:
            tempData[line] = datadict[line]
    match_data = {}
    for data in tempData:
        latest_time = find_latest(tempData[data])
        if latest_time.datetime.year == current_time.year:
            if current_time.month - latest_time.datetime.month >= 3:
                print(tempData[data])
                tempData.pop(data)
        tempData.pop(data)
    return tempData


def half_pay(datadict: dict):
    pass


def year_pay(datadict: dict):
    pass


def query(month_variable: tk.StringVar, payway_variable: tk.StringVar, result_frame, data):
    result_model = TableModel()

    datadict = renew_data(data)
    month = month_variable.get()
    payway = payway_variable.get()
    result_dict = {}
    
    current_date = datetime.date.today()
    if payway == "0":
        result_dict = season_pay(datadict,current_date)
    if payway == "1":
        result_dict = half_pay(datadict)
    if payway == "2":
        result_dict = year_pay(datadict)
    result_model.importDict(result_dict)
    result_table = TableCanvas(result_frame, model=result_model)
    result_table.setModel(result_model)
    result_table.show()


def create_query_window(root):
    # print("query function")
    # print(somthing)
    query_window = Toplevel(root)
    query_window.title("Query Window")
    # canvas = TableCanvas(query_window, read_only=True)
    # canvas.show()
    current_date = datetime.date.today()
    print(current_date)
    months = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
    month_variable = tk.StringVar()
    month_variable.set(current_date.month-1)
    selection_label = tk.Label(query_window, text="月份选择")
    selection_label.pack()
    month_combo = ttk.Combobox(
        query_window, textvariable=month_variable, values=months, width=10)
    month_combo.pack()
    # set default month as current month
    month_combo.current(month_variable.get())

    payway_label = tk.Label(query_window, text="付钱方式")
    payway_label.pack()

    payway_variable = tk.StringVar()
    payway_variable.set(3)
    # 0: every 3 months; 1: every 6 months; 2: every year; 3: all
    payways = [0, 1, 2, 3]
    payway_combo = ttk.Combobox(
        query_window, textvariable=payway_variable, values=payways, width=10)
    payway_combo.pack()
    payway_combo.current(3)
    #     print(excel[col])
    result_frame = tk.Frame(query_window)
    query_button = tk.Button(query_window, text="query button", command=partial(
        query, month_variable, payway_variable, result_frame, data))
    query_button.pack()
    print(month_variable.get())
    # all the data
    global excel_file
    excel = pd.read_excel(excel_file)
    # payway is a DataFrame type
    payway = excel.iloc[:, [3]]
    if payway.loc[3].values[0] == 0:
        print(payway.loc[3])
    # print(payway.loc[3])
    # excel.keys() is the head name of each column
    # for col in excel.keys():
    result_frame.pack()
    result_label_rental = tk.Label()
    query_window.mainloop()


root = tk.Tk()
root.title("Main Window")
root.geometry("800x600")
frame = tk.Frame(root)
datadict = renew_data(data)
table_model = TableModel()
table_model.importDict(datadict)
whole_table = TableCanvas(frame, model=table_model)
whole_table.show()
open_file_button = tk.Button(
    root, text="选择文件", command=partial(choose_excel, whole_table, datadict))
open_file_button.pack()
frame.pack()
query_button = tk.Button(root, text="query button", command=partial(
    create_query_window, root))
query_button.pack(side=BOTTOM)
root.mainloop()
