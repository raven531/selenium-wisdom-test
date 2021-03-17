import re
import pandas as pd
import xlwings

file_path = "C:\\Users\\bingjiunchen\\Desktop\\明星3缺1_20200311_展開版.xlsx"


# ptn = r"-(.+[\d])"
# table = "盤點表展開"
#
#
# def parse_list_name(list_name):
#     result = re.search(ptn, list_name)
#     return result.group(1) + table


def read_xlsx(sheet: str, row_name):
    row_data = pd.read_excel(file_path, sheet_name=sheet, engine="openpyxl")
    row_arr = row_data[row_name].array
    return [i for i in row_arr]


def write_xlsx(row_list, sheet_name, col_name, col_pos):
    workbook = xlwings.Book(file_path)
    worksheet = workbook.sheets[sheet_name]
    worksheet[col_pos].options(pd.DataFrame, index=False).value = pd.DataFrame(
        {col_name: row_list})


def compare_response_with_answer(require, sheet_name, row_name):
    query = read_xlsx(sheet=sheet_name, row_name=row_name)
    return query.count(require) > 0
