import pandas as pd
import os, json
from datetime import datetime, timedelta
from LiveInfo import LiveInfo
from Product import Product
import generatePDF

# Load configurations
def loadConfig(config_path):
    f = open(config_path,'r', encoding='UTF-8')
    config_data = json.load(f)
    return config_data

workpath = os.path.abspath(os.path.join(os.getcwd(), ""))
config = loadConfig(workpath+'/config.json')

def read_timetable(source_file, date):
    source_table = pd.read_excel(source_file)
    account_table = pd.read_excel(source_file, 1)
    product_table = pd.read_excel(config['product_file'])

    prod_influencer_col = config["producttable_influencer_column_title"]
    prod_link_col = config["producttable_link_column_num"]
    prod_id_col = config["producttable_id_column_num"]
    prod_name_col = config["producttable_name_column_num"]
    prod_alias_col = config["producttable_alias_column_num"]
    prod_code_col = config["producttable_code_column_num"]

    src_influencer_col = config["sourcetable_influencer_column_title"]
    src_date_col = config["sourcetable_date_column_title"]
    src_livetype_col = config["sourcetable_livetype_column_title"]

    account_influencer_col = config["accounttable_influencer_column_title"]
    account_accountmail_col = config["accounttable_account_column_title"]

    product_table[prod_influencer_col] = product_table[prod_influencer_col]\
        .fillna(method='pad')
    product_table[prod_influencer_col] = product_table[prod_influencer_col].str.strip()
    datas = source_table.loc[source_table[src_date_col] == date]
    influencers = datas[src_influencer_col]
    liveinfos = []
    for influencer in influencers:
        print(influencer)
        products = []
        liveline = datas.loc[datas[src_influencer_col] == influencer.strip()]
        livetypes = liveline[src_livetype_col].to_numpy()
        for livetype in livetypes:
            if livetype in config['single_live_types']:
                account_line = account_table.loc[
                    account_table[account_influencer_col] == influencer.strip()]
                account = account_line[account_accountmail_col].to_numpy()[0]
                products_list = product_table.loc[
                    product_table[prod_influencer_col] == influencer.strip()
                    ].to_numpy()
                if livetype in config['normal_live_types']:
                    for prod in products_list:
                        # 如果产品链接非空
                        if isinstance(prod[prod_link_col], str):
                            # 新Product对象，参数产品ID
                            product = Product(str(int(prod[prod_id_col])).strip())
                            product.codes = []
                            product.name = prod[prod_name_col]
                            product.alias = prod[prod_alias_col]
                            product.link = prod[prod_link_col].strip()
                            # 如果产品Code非空
                            if isinstance(prod[prod_code_col], str):
                                # 添加产品Code
                                product.codes.append(prod[prod_code_col])
                            products.append(product)
                liveinfo = LiveInfo(influencer, account, date, products)
                liveinfos.append(liveinfo)

    return liveinfos


if __name__ == '__main__':
    while True:
        source_file = input("输入要导入的源排期表路径：")
        if not source_file : 
            source_file = config['default_sourcefile']
            print(source_file)
        if os.path.exists(source_file):
            break
    while True:
        today = datetime.today()
        tomorrow = today + timedelta(1)
        date_text = input("输入要生成的文件对应日期(yyyymmdd, "\
            "'+' for tomorrow)：")
        if not date_text:
            date_text = today.strftime('%Y%m%d')
        if date_text == '+':
            date_text = tomorrow.strftime('%Y%m%d')

        try:
            date = datetime.strptime(date_text, '%Y%m%d')
            print(date)
            break
        except Exception as e:
            pass
    while True:
        h_time_string = input("输入时间表高度（default " \
            + str(config['timetable_height']) + "）：")
        if not h_time_string:
            h_time = config['timetable_height']
        else:
            h_time = int(h_time_string)

        if h_time > 0:
            print('时间表高度: ', h_time)
            break
    while True:
        h_prod_string = input("输入产品表高度（default " \
            + str(config['producttable_height']) + "）：")
        if not h_prod_string:
            h_prod = config['producttable_height']
        else:
            h_prod = int(h_prod_string)

        if h_prod > 0:
            print('产品表高度: ', h_prod)
            break

    live_infos = read_timetable(source_file, date)
    generatePDF.generate_PDF(live_infos, config['image_timetable'], \
        config['image_producttable'], h_time, h_prod)
