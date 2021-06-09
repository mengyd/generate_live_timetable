import pandas as pd
import os, json
import calendar
from datetime import datetime, timedelta
from Live import Live

pd.set_option('display.max_rows', None)
# Load configurations
def loadConfig(config_path):
    f = open(config_path,'r', encoding='UTF-8')
    config_data = json.load(f)
    return config_data

def loadAllias(allias_path):
    f = open(allias_path,'r', encoding='UTF-8')
    dict_allias = json.load(f)
    return dict_allias

workpath = os.path.abspath(os.path.join(os.getcwd(), ""))
config = loadConfig(workpath+'/config.json')
allias = loadAllias(workpath+'/allias.json')

def read_source_file(source_file):
    source_df = pd.read_excel(source_file)
    source_df2 = pd.read_excel(source_file, 1)
    source_file_lives = []
    for row in source_df.itertuples():
        live = Live(influencer=row[1], start_time=row[2], end_time=row[3],
                    live_type=row[4], date=row[5], weekday=row[6],
                    scene=row[7], comment=row[8])
        source_file_lives.append(live)
    return source_file_lives, source_df2

def read_update_source(update_source):
    update_source_df = pd.read_excel(update_source)
    update_source_file_lives = []
    for row in update_source_df.itertuples():
        live = Live(influencer=row[1], start_time=row[2], end_time=row[3],
                    live_type=row[4], date=row[5], weekday=row[6],
                    scene=row[7], comment=row[8])
        update_source_file_lives.append(live)
    return update_source_file_lives

def merge_sources(source_file_lives, update_source_file_lives, begin, end):
    updated_lives = []
    for live in source_file_lives:
        if datetime.strptime(datetime.strftime(live.date, "%Y%m%d"), '%Y%m%d') < begin:
            updated_lives.append(live)
    for live in update_source_file_lives:
        if datetime.strptime(live.date, '%Y-%m-%d') >= begin and datetime.strptime(live.date, '%Y-%m-%d') <= end:
            live.date = datetime.strptime(live.date, '%Y-%m-%d')
            updated_lives.append(live)
    return updated_lives

def write_data(updated_lives, account_sheet_df):
    influencers = []
    dates = []
    start_times = []
    end_times = []
    live_types = []
    weekdays = []
    scenes = []
    comments = []

    for live in updated_lives:
        influencers.append(live.influencer)
        dates.append(datetime.strftime(live.date, '%Y-%m-%d'))
        start_times.append(live.start_time)
        end_times.append(live.end_time)
        live_types.append(live.live_type)
        weekdays.append(live.weekday)
        scenes.append(live.scene)
        comments.append(live.comment)

    datas = {'主播':influencers, '开始':start_times,
        '结束':end_times, '类型':live_types,'日期':dates, 
        '星期':weekdays, '景':scenes, '备注':comments}

    lives_df = pd.DataFrame(data=datas)
    today = datetime.today()
    thismonth = today.strftime('%m')
    writer = pd.ExcelWriter('./' + str(thismonth) + '月排期细表-updated.xlsx')
    lives_df.to_excel(writer, sheet_name='sheet1', index=None)
    account_sheet_df.to_excel(writer, sheet_name='sheet2', index=None)
    writer.save()


if __name__ == '__main__':
    while True:
        source_file = input("输入要更新的排期表路径：")
        if not source_file : 
            source_file = config['default_sourcefile']
            print(source_file)
        if os.path.exists(source_file):
            break
    while True:
        update_source = input("输入要导入的源排期表路径：")
        if not update_source :
            update_source = config['update_source_file']
            print(update_source)
        if os.path.exists(update_source):
            break
    while True:
        today = datetime.today()
        tomorrow = today + timedelta(1)
        update_scope = input("输入更新范围，1全部，2今天及之后，3今后, 4手动输开始入日期(yyyymmdd)：")
        if not update_scope:
            update_scope = '3'
        if update_scope == '1':
            begin_text = datetime(today.year,today.month,1).strftime('%Y%m%d')
            end_text = datetime(today.year,today.month,calendar.monthrange(today.year,today.month)[1]).strftime('%Y%m%d')
        if update_scope == '2':
            begin_text = today.strftime('%Y%m%d')
            end_text = datetime(today.year,today.month,calendar.monthrange(today.year,today.month)[1]).strftime('%Y%m%d')
        if update_scope == '3':
            begin_text = tomorrow.strftime('%Y%m%d')
            end_text = datetime(today.year,today.month,calendar.monthrange(today.year,today.month)[1]).strftime('%Y%m%d')
        if update_scope == '4':
            begin_text = input("开始日期：")
            end_text = datetime(today.year,today.month,calendar.monthrange(today.year,today.month)[1]).strftime('%Y%m%d')
        try:
            begin = datetime.strptime(begin_text, '%Y%m%d')
            end = datetime.strptime(end_text, '%Y%m%d')
            print(begin, end)
            break
        except Exception as e:
            pass
    source_lives, account_sheet_df = read_source_file(source_file)
    new_lives = read_update_source(update_source)
    updated_lives = merge_sources(source_lives, new_lives, begin, end)
    write_data(updated_lives, account_sheet_df)
