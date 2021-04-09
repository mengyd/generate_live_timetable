import pandas as pd
import os, math
from datetime import datetime
from Live import Live


def hasNumbers(source_string):
    for char in source_string:
        if char.isdigit():
            return True
    return False

def headerToDate(header_string):
    date_string = header_string.strip().split(' ')[0]
    date = datetime.strptime(date_string, '%Y/%m/%d')
    return date

def formatTimeString(time_string):
    if time_string.endswith('H') :
        time_string = time_string.replace('H', ':00')
    time_string = time_string.replace('H', ':')
    return time_string
    
def extractTimes(time_string):
    time_string = time_string.split('/')[0]
    time_string = time_string.split(' ')[0]
    time_string = time_string.split('\n')[0]
    time_string = time_string.split('\r')[0]
    time_string = time_string.split('+')[0]

    strings = time_string.split('-')
    start_string = formatTimeString(strings[0])
    end_string = formatTimeString(strings[1])

    start_time = datetime.strptime(start_string, '%H:%M')
    end_time = datetime.strptime(end_string, '%H:%M')
    return start_time, end_time

def readData(source_file):
    source_table = pd.read_excel(source_file)
    headers = source_table.columns
    rows = source_table.itertuples()
    lives = []
    for row in rows :
        col_index = 0
        influencer, live_type, date, weekday, start_time, end_time = '', '', None, None, None, None
        isUGC = False
        # print(row)
        for content in row :
            isLive = False
            if col_index != 0 :
                col_name = headers[col_index-1]
                if col_name == 'No.' and hasNumbers(str(content)) :
                    isUGC = True
                    print('No.:' + str(content))
                if isUGC :
                    if col_name == '主播定位' :
                        live_type = content
                        print(live_type)
                    if col_name == '主播' :
                        influencer = content
                        print(influencer)
                    if hasNumbers(str(content)) and hasNumbers(str(col_name)) :
                        date = headerToDate(col_name)
                        weekday = date.isoweekday() + 1
                        lines = content.splitlines()
                        for line in lines :
                            if hasNumbers(line) :
                                start_time, end_time = extractTimes(line.strip())
                                # print(datetime.strftime(start_time, '%H:%M') + '-' + datetime.strftime(end_time, '%H:%M'))
                            isLive = True
            col_index += 1
            live = Live(influencer, date, start_time, end_time, weekday, live_type=live_type)
            if isLive and isUGC:
                lives.append(live)

    for live in lives:
        print(live.influencer, ' : ', live.date, ' : ', live.weekday)

    return lives

def write_data(lives):
    influencers = []
    dates = []
    start_times = []
    end_times = []
    live_types = []
    weekdays = []

    i = 0
    for live in lives:
        influencers.append(live.influencer)
        start_times.append(live.start_time)
        end_times.append(live.end_time)
        live_types.append(live.live_type)
        dates.append(live.date)
        weekdays.append(live.weekday)
        i += 1
    datas = {'主播':influencers, '开始':start_times,
            '结束':end_times, '类型':live_types,
            '日期':dates, '星期':weekdays}
    df = pd.DataFrame(data=datas)
    print(df)
    today = datetime.today()
    thismonth = today.strftime('%m')
    print(thismonth)
    df.to_excel('./' + str(thismonth) + '月排期细表.xlsx')



if __name__ == '__main__':
    workpath = os.path.abspath(os.path.join(os.getcwd(), ""))
    while True:
        source_file = input("输入要导入的源排期表路径：")
        if source_file and os.path.exists(source_file) :
            break
    
    lives = readData(source_file)
    write_data(lives)
