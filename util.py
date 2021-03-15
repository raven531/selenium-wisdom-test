import re
import pandas as pd

ptn = r"-(.+[\d])"
file_path = "C:\\Users\\bingjiunchen\\Desktop\\明星3缺1_20200311_展開版.xlsm"

table = "盤點表展開"


def parse_list_name(list_name):
    result = re.search(ptn, list_name)
    return result.group(1) + table


def read_xlsx(sheet: str):
    row_data = pd.read_excel(file_path, sheet_name=sheet, engine="openpyxl")
    row_arr = row_data["品檢問題填寫處"].array
    return [i for i in row_arr if type(i) != float]


# TODO
def write_xlsx(row_list: list, sheet_name: str):
    df = pd.DataFrame(row_list, columns=["answer"])
    df.to_excel(file_path, sheet_name=sheet_name, index=False)


# TODO
def compare_response_with_answer():
    return
