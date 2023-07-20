import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
from os import path
from tkintertable import TableModel


def excel_to_dict(excel_file):
    if not path.exists(excel_file):
        return {0: {'房号': '', '公司名称': '', '合同日期': '', '付款方式': '', '房租': '', '物管费': '', '第一次缴费': '', '第二次缴费': '', '第三次缴费': '', '第四次缴费': '', '保证金': '', '备注': ''}}
    excel_df = pd.read_excel(excel_file)
    excel_df = excel_df.astype(str).where(pd.notnull(excel_df), '')
    excel_df = excel_df.fillna('')
    excel_list = excel_df.to_dict('records')
    excel_dict = {}
    for i, data in enumerate(excel_list):
        excel_dict[i] = data
    return excel_dict


def find_latest_date(data: dict):
    # print("find latest")
    pay_dates = list()
    if data["第一次缴费"] != "":
        pay_dates.append(data["第一次缴费"])
    if data["第二次缴费"] != "":
        pay_dates.append(data["第二次缴费"])
    if data["第三次缴费"] != "":
        pay_dates.append(data["第三次缴费"])
    if data["第四次缴费"] != "":
        pay_dates.append(data["第四次缴费"])
    latest_date = ""
    if pay_dates.__len__() > 0:
        latest_date = datetime.datetime.strptime(pay_dates[0], '%Y-%m-%d')
        for pay_date in pay_dates:
            temp_date = datetime.datetime.strptime(pay_date, '%Y-%m-%d')
            if temp_date-latest_date > datetime.timedelta(0):
                latest_date = temp_date
    return datetime.datetime.strftime(latest_date, '%Y-%m-%d')


def season_pay(month: str, data: dict):
    month_int = int(month)
    result_data = {}
    for line in data:
        contract_date = datetime.datetime.strptime(
            data[line]["合同日期"], '%Y-%m-%d')
        if (month_int + 12 - contract_date.month) % 3 == 0 and data[line]["付款方式"] == "季度":
            result_data[line] = data[line]
    # print(result_data)
    return result_data


def halfyear_pay(month: str, data: dict):
    month_int = int(month)
    result_data = {}
    for line in data:
        contract_date = datetime.datetime.strptime(
            data[line]["合同日期"], '%Y-%m-%d')
        if (month_int + 12 - contract_date.month) % 6 == 0 and data[line]["付款方式"] == "半年":
            result_data[line] = data[line]
    # print(result_data)
    return result_data


def year_pay(month: str, data: dict):
    month_int = int(month)
    result_data = {}
    for line in data:
        contract_date = datetime.datetime.strptime(
            data[line]["合同日期"], '%Y-%m-%d')
        if (month_int + 12 - contract_date.month) % 12 == 0 and data[line]["付款方式"] == "年度":
            result_data[line] = data[line]
    # print(result_data)
    return result_data


def sum_rent(data: dict):
    rent = 0
    for line in data:
        if data[line]["付款方式"] == "季度":
            rent = float(data[line]["房租"]) * 3 + rent
        if data[line]["付款方式"] == "半年":
            rent = float(data[line]["房租"]) * 6 + rent
        if data[line]["付款方式"] == "年度":
            rent = float(data[line]["房租"]) * 12 + rent
    return rent


def sum_management(data: dict):
    management = 0
    for line in data:
        if data[line]["付款方式"] == "季度":
            management = float(data[line]["物管费"]) * 3 + management
        if data[line]["付款方式"] == "半年":
            management = float(data[line]["物管费"]) * 6 + management
        if data[line]["付款方式"] == "年度":
            management = float(data[line]["物管费"]) * 12 + management
    return management


def get_begin_end(date: datetime.datetime, month: int, mode: int):
    date = datetime.datetime(
        datetime.datetime.today().year, date.month, date.day)

    if mode == 1:
        begin_date = datetime.datetime(date.year, month, date.day)
        end_date = begin_date+relativedelta(months=+3, days=-1)
        string = datetime.datetime.strftime(
            begin_date, "%Y-%m-%d") + "至" + datetime.datetime.strftime(end_date, "%Y-%m-%d")
        return string
    if mode == 2:
        begin_date = datetime.datetime(date.year, month, date.day)
        end_date = begin_date+relativedelta(months=+6, days=-1)
        string = datetime.datetime.strftime(
            begin_date, "%Y-%m-%d") + "至" + datetime.datetime.strftime(end_date, "%Y-%m-%d")
        return string
    if mode == 3:
        begin_date = datetime.datetime(date.year, month, date.day)
        end_date = begin_date+relativedelta(months=+12, days=-1)
        string = datetime.datetime.strftime(
            begin_date, "%Y-%m-%d") + "至" + datetime.datetime.strftime(end_date, "%Y-%m-%d")
        return string
    else:
        return ""

def get_longest_string(data:TableModel):
    pass