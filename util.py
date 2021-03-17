import re
import pandas as pd
import xlwings
from difflib import SequenceMatcher

file_path = "C:\\Users\\bingjiunchen\\Desktop\\金好運_20200311_展開版.xlsx"


def read_xlsx(sheet: str, row_name):
    row_data = pd.read_excel(file_path, sheet_name=sheet, engine="openpyxl")
    row_arr = row_data[row_name].array
    return [i for i in row_arr]


def write_xlsx(row_list, sheet_name, col_name, col_pos):
    workbook = xlwings.Book(file_path)
    worksheet = workbook.sheets[sheet_name]
    worksheet[col_pos].options(pd.DataFrame, index=False).value = pd.DataFrame(
        {col_name: row_list})


def compare_response_with_answer(require, sheet_name, row_name, row_index):
    query = read_xlsx(sheet=sheet_name, row_name=row_name)
    fetch = query.pop(row_index)
    require = re.sub(r'[^\w]', '', require)
    fetch = re.sub(r'[^\w]', '', fetch)
    return SequenceMatcher(None, require, fetch).ratio() == 1.0
