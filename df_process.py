import re

import pandas as pd

from url_to_df import url_to_df

cn_to_int = {'一': 0, '二': 1, '三': 2, '四': 3, '五': 4, '六': 5, '日': 6}
class_map = {'0': 0, '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '6': 6, '7': 7, '8': 8, '9': 9, '10': 10, 'A': 11, 'B': 12,
             'C': 13, 'D': 14}


def get_time(df):
    for i, raw in enumerate(df['時間教室']):
        m = re.findall(r'(\S*?)\((.*?)\)', raw)
        for group in m:
            time_raw = group[0]
            df.at[i, 'classroom'] = group[1]
            time_raw = re.match(r'週?(.)(.*)', time_raw)
            weekday = cn_to_int[time_raw.group(1)]
            periods = time_raw.group(2).split(',')
            for period in periods:
                period = class_map[period]
                df.at[i, 'time_' + str(weekday) + '_' + str(period)] = 1
    return df


def get_field(df):
    for i, raw in enumerate(df['備註']):
        m = re.search(r'A(\d*)', raw)
        if m:
            for f in m.group(1):
                df.at[i, 'A' + f] = 1
    return df


def get_df(mode='系所', dpt='9010', sem='110-1', startrec=0):
    df = url_to_df(mode, dpt, sem, startrec)
    df = get_time(df)
    df = get_field(df)
    return df


def save_csv(mode='系所', dpt='9010', sem='110-1'):
    if mode == '系所':
        name = dpt
    else:
        name = mode

    if mode == '體育':
        df = get_df(mode, dpt, sem, 0)
        for i in range(15, 210, 15):
            df = df.append(get_df(mode, dpt, sem, i), ignore_index=True, sort=False)
    else:
        df = get_df(mode, dpt, sem)
    df.to_csv('./save_csv/' + name + '.csv', encoding='utf-8-sig', index=0)


def prefered_cols(df, cols=None):
    if cols is None:
        cols = ['流水號', '課程名稱', '學分', '全/半年', '授課教師', '加選方式', '時間教室', '總人數', '選課限制條件', '備註']
    return df[cols]


if __name__ == "__main__":
    pd.set_option('display.max_columns', None)
    df = get_df(mode='系所')
    print(df.head())
    save_csv(mode='系所', dpt='9010')
    save_csv(mode='系所', dpt='9020')
    save_csv(mode='通識')
    save_csv(mode='體育')
