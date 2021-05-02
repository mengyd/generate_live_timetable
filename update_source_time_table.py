import os
import pandas as pd
from datetime import datetime

def headerToDate(header_string):
    date_string = header_string.strip().split(' ')[0]
    date = datetime.strptime(date_string, '%Y/%m/%d')
    return date

def read_source(source_file):
    source_table = pd.read_excel(source_file)
    headers = source_table.columns
    dates = []
    for header in 


def read_target(target_file):
    pass


if __name__ == '__main__':
    while True:        
        source_file = input("输入要导入的源排期表路径：")
        if source_file and os.path.exists(source_file) :
            break

    while True:        
        target_file = input("输入要更新的排期表路径：")
        if target_file and os.path.exists(target_file) :
            break

    