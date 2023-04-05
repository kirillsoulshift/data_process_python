import pandas as pd
import os
import re

"""
- удалить дубликаты по перв столбцу
- удалить Source==Target
- в англоязычных каталогах проверить, что в правильном
варианте каталожного номера стоят только английские буквы
- замена происходит буквами в том же ригистре, что и буква в оригинале"""

d_low = {'к': 'k', 'е': 'e', 'н': 'h', 'х': 'x', 'в': 'b', 'а': 'a', 'р': 'p', 'о': 'o', 'с': 'c', 'м': 'm', 'т': 't'}
d_up = {'К': 'K', 'Е': 'E', 'Н': 'H', 'Х': 'X', 'В': 'B', 'А': 'A', 'Р': 'P', 'О': 'O', 'С': 'C', 'М': 'M', 'Т': 'T'}


def find_cyrrilic(s):
    if re.search(r'[А-Яа-я]', s):
        return True
    else:
        return False


def find_latin(s):
    if re.search(r'[A-Za-z]', s):
        return True
    else:
        return False


def replace_cyr(s):
    if find_cyrrilic(s):
        for k in d_low.keys():
            if re.search(k, s):
                s = re.sub(k, d_low[k], s)
        for k in d_up.keys():
            if re.search(k, s):
                s = re.sub(k, d_up[k], s)
        return s
    else:
        return s


def replace_lat(s):
    l = s.split()
    english_position = False
    for w in l:
        if re.search(r'[A-Za-z]{3,}', w):
            english_position = True
    if english_position:
        return s
    if find_latin(s):
        for k in d_low.keys():
            if re.search(d_low[k], s):
                s = re.sub(d_low[k], k, s)
        for k in d_up.keys():
            if re.search(d_up[k], s):
                s = re.sub(d_up[k], k, s)
        return s
    else:
        return s


curr_path = os.getcwd()
filenames = os.listdir(curr_path)
files = []
for file in filenames:
    if file.endswith('.xlsx') and file.count('~') == 0:
        files.append(file)
if len(files) == 0:
    raise Exception(f'Нет файлов для обработки в пути {curr_path}')
# english_partnumbers = True
# resp = ''
#
# while not re.search(r'eng|rus', resp, flags=re.I):
#     resp = input('Принадлежность букв в партномерах:\n'
#                  '    английские - eng\n'
#                  '    русские - rus')
#
# if re.search(r'eng', resp, flags=re.I):
#     english_partnumbers = True
# elif re.search(r'rus', resp, flags=re.I):
#     english_partnumbers = False

for file in files:
    df = pd.read_excel(os.path.join(curr_path, file), header=None, dtype=str)
    english_partnumbers = True
    if re.search(r'mash|belaz', file, flags=re.I):
        english_partnumbers = False

    if re.search(r'epiroc', file, flags=re.I):
        df.columns = ['Source', 'Target', 'Comments']

        df.drop_duplicates(subset='Source', inplace=True, ignore_index=True)

        for n in range(len(df)):
            if df.loc[n, 'Comments'] == 'В LO КН 8и значные, без "26", в ТК с "26" ЕСТЬ' or \
                    df.loc[n, 'Comments'] == 'В LO КН 8и значные, без "26", в ТК с "26" нет':
                df.at[n, 'Target'] = '26' + df.loc[n, 'Target']
            elif df.loc[n, 'Comments'] == 'В LO КН без "0" спереди, в ТК с "0" есть':
                df.at[n, 'Target'] = '0' + df.loc[n, 'Target']

        check_df_1 = df[df['Source'] == df['Target']]
        check_df_2 = df[df['Target'].isna() == True]
        check_df_3 = df[df['Source'].isna() == True]
        check_list = list(set(check_df_1.index.tolist() +
                              check_df_2.index.tolist() +
                              check_df_3.index.tolist()))
        df.drop(check_list, inplace=True)
        df.reset_index(inplace=True, drop=True)

        df.loc[:, 'Target'] = df['Target'].apply(replace_cyr)

        df.to_excel(file[:-5] + '_partnumbers_checked.xlsx', index=False, header=False)

    elif re.search(r'rudgormash', file, flags=re.I):
        df.columns = ['Source', 'Target']
        df.drop_duplicates(subset='Source', inplace=True, ignore_index=True)
        check_df_1 = df[df['Source'] == df['Target']]
        check_df_2 = df[df['Target'].isna() == True]
        check_df_3 = df[df['Source'].isna() == True]
        check_list = list(set(check_df_1.index.tolist() +
                              check_df_2.index.tolist() +
                              check_df_3.index.tolist()))
        df.drop(check_list, inplace=True)


        df.to_excel(file[:-5] + '_partnumbers_checked.xlsx', index=False, header=False)

    elif re.search(r'TZJT', file, flags=re.I):
        df.columns = ['Source', 'Target', 'Desc']

        df.drop_duplicates(inplace=True, ignore_index=True)
        check_df_1 = df[df['Source'] == df['Target']]
        check_df_2 = df[df['Target'].isna() == True]
        check_df_3 = df[df['Source'].isna() == True]
        check_list = list(set(check_df_1.index.tolist() +
                              check_df_2.index.tolist() +
                              check_df_3.index.tolist()))
        df.drop(check_list, inplace=True)
        df.reset_index(inplace=True, drop=True)

        df.loc[:, 'Target'] = df['Target'].apply(replace_cyr)
        df.to_excel(file[:-5] + '_partnumbers_checked.xlsx', index=False, header=False)

    else:
        df.columns = ['Source', 'Target']
        df.drop_duplicates(subset='Source', inplace=True, ignore_index=True)
        check_df_1 = df[df['Source'] == df['Target']]
        check_df_2 = df[df['Target'].isna() == True]
        check_df_3 = df[df['Source'].isna() == True]
        check_list = list(set(check_df_1.index.tolist() +
                              check_df_2.index.tolist() +
                              check_df_3.index.tolist()))
        df.drop(check_list, inplace=True)
        df.reset_index(inplace=True, drop=True)

        if english_partnumbers:
            df.loc[:, 'Target'] = df['Target'].apply(replace_cyr)
        else:
            df.loc[:, 'Target'] = df['Target'].apply(replace_lat)
        df.to_excel(file[:-5] + '_partnumbers_checked.xlsx', index=False, header=False)