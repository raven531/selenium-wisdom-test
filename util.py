import re
import pandas as pd

ptn = r"-(.+[\d])"
file_path = "C:\\Users\\bingjiunchen\\Downloads\\鈊象盤點表20200311.xlsx"

table = "盤點表"


def parse_list_name(list_name):
    result = re.search(ptn, list_name)
    return result.group(1) + table


def read_xlsx(sheet: str):
    arr = []
    row_data = pd.read_excel(file_path, sheet_name=sheet, engine="openpyxl")
    # questions = row_data["原始問句"].array
    #
    # for i in range(len(questions)):
    #     if type(questions[i]) == float:
    #         continue
    #     arr.append(questions[i])
    # return arr
    return row_data["原始問句"].array
