import pandas as pd
import os
import calendar
from datetime import datetime, timedelta
from Live import Live
from loadconfig import loadConfig

# Print all rows
pd.set_option('display.max_rows', None)
# Load configurations
config = loadConfig('params')

def read_source_file(source_file):
    source_df = pd.read_excel(source_file)
    source_file_lives = []
    for row in source_df.itertuples():
        live = Live(influencer=row[1], start_time=row[2], end_time=row[3],
                    live_type=row[4], date=row[5], weekday=row[6],
                    scene=row[7], comment=row[8])
        source_file_lives.append(live)
    return source_file_lives

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
        live.date = datetime.strptime(datetime.strftime(live.date, "%Y%m%d"), '%Y%m%d')
        if live.date < begin:
            updated_lives.append(live)
    for live in update_source_file_lives:
        live.date = datetime.strptime(live.date, '%Y-%m-%d')
        live.start_time = datetime.strptime(live.start_time, '%H:%M:%S').time()
        live.end_time = datetime.strptime(live.end_time, '%H:%M:%S').time()
        if live.date >= begin and live.date <= end:
            updated_lives.append(live)
    return updated_lives

def write_data(updated_lives):
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
        dates.append(live.date.date())
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
        update_scope = input("输入更新范围，1本月全部，2本月今天及之后，3本月今后, 4手动输开始入日期(yyyymmdd)：")
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
            year = int(begin_text[:4])
            month = int(begin_text[5:6])
            lastday = calendar.monthrange(year,month)[1]
            end_date = datetime(year, month, lastday)
            end_text = end_date.strftime('%Y%m%d')
        try:
            begin = datetime.strptime(begin_text, '%Y%m%d')
            end = datetime.strptime(end_text, '%Y%m%d')
            print(begin, end)
            break
        except Exception as e:
            pass
    source_lives = read_source_file(source_file)
    new_lives = read_update_source(update_source)
    updated_lives = merge_sources(source_lives, new_lives, begin, end)
    write_data(updated_lives)
