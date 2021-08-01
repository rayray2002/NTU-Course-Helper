from os import path, makedirs

import pandas as pd
from styleframe import StyleFrame

from df_process import *

StyleFrame.A_FACTOR = 13
StyleFrame.P_FACTOR = 1.5


def input_schedule():
    name = input()
    fields = input()
    schedule = [[0] * 15 for i in range(7)]
    for i in range(7):
        raw = input()
        if len(raw) > 1:
            periods = raw[1:].split(',')
            for period in periods:
                schedule[i][class_map[period]] = 1
    raw = input()
    dpts = []
    while len(raw) > 0:
        dpts.append(raw)
        raw = input()
    print(name + ' imported')
    return name, fields, schedule, dpts


def general_filter(name, writer):
    if not path.exists('save_csv/通識.csv'):
        save_csv(mode='通識')
    df = pd.read_csv('save_csv/通識.csv')

    for i in range(7):
        for j in range(15):
            if schedule[i][j]:
                df.drop(df[df['time_' + str(i) + '_' + str(j)] == 1].index, inplace=True)
    df.drop(df[~df[['A' + f for f in fields]].any(axis=1)].index, inplace=True)
    if name == 'Ray':
        df.drop(df[df['備註'].str.contains('限非電資|限電資學院以外', regex=1)].index, inplace=True)
    df = prefered_cols(df)
    sf = StyleFrame(df)
    sf.to_excel(writer, sheet_name='通識', index=0, row_to_add_filters=0,
                best_fit=['流水號', '課程名稱', '學分', '全/半年', '授課教師', '加選方式', '時間教室', '總人數', '選課限制條件', '備註'])
    print(name + ' 通識 export success')
    return df


def department_filter(name, dpt, writer):
    if type(dpt) == int:
        dpt = str(dpt)
    if not path.exists('save_csv/' + dpt + '.csv'):
        save_csv(mode='系所', dpt=dpt)
    df = pd.read_csv('save_csv/' + dpt + '.csv')

    for i in range(7):
        for j in range(15):
            if schedule[i][j]:
                df.drop(df[df['time_' + str(i) + '_' + str(j)] == 1].index, inplace=True)
    df.drop(df[df['課程名稱'].str.contains('服務學習|專題', regex=1)].index, inplace=True)

    df = prefered_cols(df)
    sf = StyleFrame(df)
    sf.to_excel(writer, sheet_name=dpt, index=0, row_to_add_filters=0,
                best_fit=['流水號', '課程名稱', '學分', '全/半年', '授課教師', '加選方式', '時間教室', '總人數', '選課限制條件', '備註'])
    print(name + ' ' + dpt + ' export success')
    return df


def language_filter(name, writer):
    if not path.exists('save_csv/共同.csv'):
        save_csv(mode='共同')
    df1 = pd.read_csv('save_csv/共同.csv')

    if not path.exists('save_csv/外文.csv'):
        save_csv(mode='外文')
    df2 = pd.read_csv('save_csv/外文.csv')

    df = df1.append(df2, ignore_index=True, sort=False)
    # print(df)
    for i in range(7):
        for j in range(15):
            if schedule[i][j]:
                df.drop(df[df['time_' + str(i) + '_' + str(j)] == 1].index, inplace=True)
    df.drop(df[df['課程名稱'].str.contains('國際|僑生', regex=1)].index, inplace=True)
    df.drop(df[df['選課限制條件'].str.contains('限生農學院學生|限僑生、國際學生|限醫學院學生|限文學院學生|限學士班一年級|限國際學生', regex=1)].index, inplace=True)

    df = prefered_cols(df)
    sf = StyleFrame(df)
    sf.to_excel(writer, sheet_name='共同', index=0, row_to_add_filters=0,
                best_fit=['流水號', '課程名稱', '學分', '全/半年', '授課教師', '加選方式', '時間教室', '總人數', '選課限制條件', '備註'])
    print(name + ' 語文 export success')
    return df


def pe_filter(name, writer):
    if not path.exists('save_csv/體育.csv'):
        save_csv(mode='體育')
    df = pd.read_csv('save_csv/體育.csv')

    for i in range(7):
        for j in range(15):
            if schedule[i][j]:
                df.drop(df[df['time_' + str(i) + '_' + str(j)] == 1].index, inplace=True)
    # df.drop(df[df['課程名稱'].str.contains('國際|僑生', regex=1)].index, inplace=True)
    # df.drop(df[df['選課限制條件'].str.contains('限生農學院學生|限僑生、國際學生|限醫學院學生|限文學院學生|限學士班一年級|限國際學生', regex=1)].index, inplace=True)

    df = prefered_cols(df)
    sf = StyleFrame(df)
    sf.to_excel(writer, sheet_name='體育', index=0, row_to_add_filters=0,
                best_fit=['流水號', '課程名稱', '學分', '全/半年', '授課教師', '加選方式', '時間教室', '總人數', '選課限制條件', '備註'])
    print(name + ' 體育 export success')
    return df


if __name__ == "__main__":
    if not path.exists('./out'):
        makedirs('./out')
    if not path.exists('./save_csv'):
        makedirs('./save_csv')
    while True:
        try:
            name, fields, schedule, dpts = input_schedule()
            if not path.exists('./out/' + name + '.xlsx'):
                open('./out/' + name + '.xlsx', 'w')
            writer = pd.ExcelWriter('./out/' + name + '.xlsx')

            general_df = general_filter(name, writer)
            language_filter(name, writer)
            pe_filter(name, writer)
            for dpt in dpts:
                department_filter(name, dpt, writer)

            writer.save()
            print(name + '.xlsx save success')

        except EOFError:
            print('Input end')
            break
        except Exception as e:
            print(e)
