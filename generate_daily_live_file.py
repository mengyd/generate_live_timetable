import pandas as pd
import os, math
from datetime import datetime
from LiveInfo import LiveInfo
import generatePDF

__default_sourcefile__ = 'C:\\Users\\yidon\\OneDrive\\桌面\\05月排期细表.xlsx'
__product_file__ = 'C:\\Users\\yidon\\OneDrive\\桌面\\5月UGC产品总表.xlsx'
__image_timetable__ = 'C:\\Users\\yidon\\OneDrive\\桌面\\t.png'
__image_producttable__ = 'C:\\Users\\yidon\\OneDrive\\桌面\\p.png'
__normal_live_types__ = ["UGC","IP-UGC","运动","华为首发"]

def read_timetable(source_file, date):
    source_table = pd.read_excel(source_file)
    second_table = pd.read_excel(source_file, 1)
    product_table = pd.read_excel(__product_file__)
    product_table['主播'] = product_table['主播']\
        .fillna(method='pad')
    product_table['主播'] = product_table['主播'].str.strip()
    datas = source_table.loc[source_table['日期'] == date]
    influencers = datas['主播']
    liveinfos = []
    for influencer in influencers:
        print(influencer)
        products_dict = {}
        products_dict2 = {}
        codes = {}
        liveline = datas.loc[datas['主播'] == influencer.strip()]
        livetypes = liveline['类型'].to_numpy()
        for livetype in livetypes:
            if livetype in __normal_live_types__:
                account_line = second_table.loc[
                    second_table['主播'] == influencer.strip()]
                account = account_line['账号'].to_numpy()[0]
                products = product_table.loc[
                    product_table['主播'] == influencer.strip()
                    ].to_numpy()
                for product in products:
                    # 产品名
                    if isinstance(product[2], str):
                        # 产品名对应产品链接
                        # print(product[5], product[2])
                        products_dict[product[5]] = product[2].strip()
                    # 产品ID
                    if isinstance(product[4], str):
                        # 产品名对应产品ID
                        products_dict2[product[5]] = str(int(product[3])).strip()
                    # 产品Code
                    if isinstance(product[6], str):
                        # 产品名对应产品Code
                        codes[product[5]] = product[6]
                liveinfo = LiveInfo(influencer, account, date, products_dict, products_dict2, codes)
                liveinfos.append(liveinfo)

    return liveinfos


if __name__ == '__main__':
    workpath = os.path.abspath(os.path.join(os.getcwd(), ""))
    while True:
        source_file = input("输入要导入的源排期表路径：")
        if not source_file : 
            source_file = __default_sourcefile__
            print(source_file)
        if os.path.exists(source_file):
            break
    while True:
        date_text = input("输入要生成的文件对应日期(yyyymmdd)：")
        if not date_text:
            date_text = datetime.today().strftime('%Y%m%d')

        try:
            date = datetime.strptime(date_text, '%Y%m%d')
            print(date)
            break
        except Exception as e:
            pass
    while True:
        h_time_string = input("输入时间表高度（default 150）：")
        if not h_time_string:
            h_time = 150
        else:
            h_time = int(h_time_string)

        if h_time > 0:
            print('时间表高度: ', h_time)
            break
    while True:
        h_prod_string = input("输入产品表高度（default 450）：")
        if not h_prod_string:
            h_prod = 450
        else:
            h_prod = int(h_prod_string)

        if h_prod > 0:
            print('产品表高度: ', h_prod)
            break

    live_infos = read_timetable(source_file, date)
    generatePDF.generate_PDF(live_infos, __image_timetable__, \
        __image_producttable__, h_time, h_prod)
