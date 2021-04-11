import pandas as pd
import os, math
from datetime import datetime
from LiveInfo import LiveInfo
import generatePDF

__product_file__ = 'C:\\Workspace\\generate_live_timetable\\UGC产品总表.xlsx'
__image_timetable__ = 'C:\\Workspace\\generate_live_timetable\\t.png'
__image_producttable__ = 'C:\\Workspace\\generate_live_timetable\\p.png'

def read_timetable(source_file, date):
    source_table = pd.read_excel(source_file)
    second_table = pd.read_excel(source_file, 1)
    product_table = pd.read_excel(__product_file__)
    product_table['Présentateur'] = product_table['Présentateur'].fillna(method='pad')
    datas = source_table.loc[source_table['日期'] == date]
    influencers = datas['主播']
    liveinfos = []
    for influencer in influencers:
        print(influencer)
        products_dict = {}
        codes = {}
        account_line = second_table.loc[second_table['主播'] == influencer.strip()]
        accounts = account_line['账号']
        account = accounts.to_numpy()[0]
        # print(account)
        products = product_table.loc[product_table['Présentateur'] == influencer.strip()].to_numpy()
        for product in products:
            products_dict[product[3]] = product[2].strip()
            codes[product[3]] = product[4]
        liveinfo = LiveInfo(influencer, account, date, products_dict, codes)
        liveinfos.append(liveinfo)

    return liveinfos


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

    liveinfos = read_timetable(source_file, date)
    generatePDF.generate_PDF(liveinfos, __image_timetable__, __image_producttable__)
