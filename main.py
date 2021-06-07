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
    time_string = time_string + ':00'
    return time_string
   
def extractTimes(time_string):
    time_string = time_string.split('/')[0]
    time_string = time_string.split(' ')[0]
    time_string = time_string.split('\n')[0]
    time_string = time_string.split('\r')[0]
    time_string = time_string.split('+')[0]

    strings = time_string.split('-')
    start_time = formatTimeString(strings[0])
    end_time = formatTimeString(strings[1])

    return start_time, end_time

def readData(source_file):
    source_table = pd.read_excel(source_file)
    source_table['主播'] = source_table['主播'].str.strip()
    headers = source_table.columns
    rows = source_table.itertuples()
    lives = []
    for row in rows :
        col_index = 0
        influencer, live_type, date, start_time, end_time = '', '', None, None, None
        isUGC = False
        # print(row)
        for content in row :
            isLive = False
            if col_index != 0 :
                col_name = headers[col_index-1]
                if col_name == 'No.' and hasNumbers(str(content)) :
                    isUGC = True
                if isUGC :
                    live_type = 'UGC'
                    if col_name == '主播' :
                        influencer = content
                    if hasNumbers(str(content)) and hasNumbers(str(col_name)) :
                        date = headerToDate(col_name)
                        lines = content.splitlines()
                        for line in lines :
                            if hasNumbers(line) :
                                start_time, end_time = extractTimes(line.strip())
                            isLive = True
            col_index += 1
            live = Live(influencer, date, start_time, end_time, live_type)
            if isLive and isUGC:
                lives.append(live)

    return lives

def write_data(lives):
    influencers = []
    dates = []
    start_times = []
    end_times = []
    live_types = []
    weekdays = []
    weekdayChinese = {
        'Monday':'周一',
        'Tuesday':'周二',
        'Wednesday':'周三',
        'Thursday':'周四',
        'Friday':'周五',
        'Saturday':'周六',
        'Sunday':'周日',
    }

    i = 0
    for live in lives:
        influencers.append(live.influencer)
        start_times.append(live.start_time)
        end_times.append(live.end_time)
        live_types.append(live.live_type)
        dates.append(datetime.strftime(live.date, '%Y-%m-%d'))
        weekdays.append(weekdayChinese[datetime.strftime(live.date, '%A')])
        i += 1
    datas = {'主播':influencers, '开始':start_times,
            '结束':end_times, '类型':live_types,
            '日期':dates, '星期':weekdays, '景':None, '备注':None}
    df = pd.DataFrame(data=datas)
    print(df)
    today = datetime.today()
    thismonth = today.strftime('%m')
    df.to_excel('./' + str(thismonth) + '月排期细表.xlsx', index=None)



if __name__ == '__main__':
    workpath = os.path.abspath(os.path.join(os.getcwd(), ""))
    while True:
        source_file = input("输入要导入的源排期表路径：")
        if source_file and os.path.exists(source_file) :
            break
    
    lives = readData(source_file)
    write_data(lives)
