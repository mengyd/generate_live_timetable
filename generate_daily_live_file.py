import pandas as pd
import os, math
from datetime import datetime
from LiveInfo import LiveInfo
from Product import Product
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
        products = []
        liveline = datas.loc[datas['主播'] == influencer.strip()]
        livetypes = liveline['类型'].to_numpy()
        for livetype in livetypes:
            if livetype in __normal_live_types__:
                account_line = second_table.loc[
                    second_table['主播'] == influencer.strip()]
                account = account_line['账号'].to_numpy()[0]
                products_list = product_table.loc[
                    product_table['主播'] == influencer.strip()
                    ].to_numpy()
                for prod in products_list:
                    # 如果产品链接非空
                    if isinstance(prod[2], str):
                        # 新Product对象，参数产品ID
                        product = Product(str(int(prod[3])).strip())
                        product.codes = []
                        product.name = prod[4]
                        product.alias = prod[5]
                        product.link = prod[2].strip()
                        # 如果产品Code非空
                        if isinstance(prod[6], str):
                            # 添加产品Code
                            product.codes.append(prod[6])
                        print(product.codes)
                        products.append(product)
                liveinfo = LiveInfo(influencer, account, date, products)
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
