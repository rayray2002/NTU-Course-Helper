from os import makedirs
from url_to_df import *


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

            general_df = general_filter(name, writer, schedule, fields)
            language_filter(name, writer, schedule)
            pe_filter(name, writer, schedule)
            for dpt in dpts:
                department_filter(name, dpt, writer, schedule)

            writer.save()
            print(name + '.xlsx save success')

        except EOFError:
            print('Input end')
            break
        except Exception as e:
            print(e)
