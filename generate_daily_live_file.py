import pandas as pd
import os, math
from datetime import datetime
from Live import Live

__product_file__ = 'C:\\Workspace\\generate_live_timetable\\UGC产品总表.xlsx'

def read_timetable(source_file, date):
    source_table = pd.read_excel(source_file)
    second_table = pd.read_excel(source_file, 1)
    product_table = pd.read_excel(__product_file__)
    product_table = product_table.fillna(method='pad')
    print(product_table)
    datas = source_table.loc[source_table['日期'] == date]
    influencers = datas['主播']
    for influencer in influencers:
        print(influencer)
        account_line = second_table.loc[second_table['主播'] == influencer.strip()]
        accounts = account_line['账号']
        print(accounts.to_numpy()[0])
        products = product_table.loc[product_table['Présentateur'] == influencer.strip()]
        print(products)
        

if __name__ == '__main__':
    workpath = os.path.abspath(os.path.join(os.getcwd(), ""))
    while True:
        source_file = input("输入要导入的源排期表路径：")
        if source_file and os.path.exists(source_file) :
            break
    while True:
        date_text = input("输入要生成的文件对应日期：")
        if not date_text:
            date_text = datetime.today().strftime('%Y%m%d')

        try:
            date = datetime.strptime(date_text, '%Y%m%d')
            print(date)
            break
        except Exception as e:
            pass

    read_timetable(source_file, date)
