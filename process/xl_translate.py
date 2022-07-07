import pandas as pd
import numpy as np
import re
import sys
import os


"""
ПРИМЕР ИСХОДНОГО .xlsx (словарь)
________________________________________________________
1  |- AIR MOTOR ASSY  | - ПНЕВМОДВИГАТЕЛЬ В СБОРЕ       |
___|__________________|_________________________________|
2  |- ASSY,PISTON     | - ПОРШЕНЬ В СБОРЕ               |
___|__________________|_________________________________|
3  |BACK UP           | КОЛЬЦО УПОРНОЕ                  |
___|__________________|_________________________________|
4  |BASE,CLEVIS       | ОПОРА СКОБЫ                     |
___|__________________|_________________________________|
   |                  |                                 |
                     ...
             <любое кол-во строк>              
                     ...
=====================================================
=====================================================
=====================================================
ПРИМЕР ИСХОДНОГО .xlsx (переводимый файл)
ПЕРВАЯ СТРОКА = НАИМЕНОВАНИЯ КОЛОНОК
________________________________________________________
1  |REF.  | PART NO.  |  QTY.  |  UNIT  |  DESCRIPTION  |  # первая строка наименовая столбцов 
___|______|___________|________|________|_______________|
2  |      |           |        |        |               |
___|______|___________|________|________|_______________|
3  |577932|           |        |        |               |  
___|______|___________|________|________|_______________|
4  |BODY,P|           |        |        |               |  
___|______|___________|________|________|_______________|
5  |REF.  | PART NO.  |  QTY.  |  UNIT  |  DESCRIPTION  |  # может содержать лишние ненужные наименовая столбцов
___|______|___________|________|________|_______________|
6  |REF.  | PART NO.  |  QTY.  |  UNIT  |  DESCRIPTION  |
___|______|___________|________|________|_______________|
7  |001   | 57793200  |  1     |  EA    |HOUSE,MACHINERY|
___|______|___________|________|________|_______________|
   |      |           |        |        |               |
                     ...
             <любое кол-во строк>              
                     ...


взаимодействует с файлами .xlsx 
содержащимися в директории скрипта.

в качестае словаря использует файл .xlsx
с совпадением 'transl' в названии

переводит файл .xlsx с содержанием
'draft' в названии путем создания
колонки перевода в нем

переводится столбец с совпадением 'desc'
переводится столбец с совпадением 'comment'
переводится первый столбец содержащий заголовки страниц

переводимый файл .xlsx является выгрузкой таблиц и
наименований из OCR проекта.
ПЕРВАЯ СТРОКА draft .xlsx = НАИМЕНОВАНИЯ КОЛОНОК

скрипт формирует новый файл .xlsx в своей директории
"""


def isNaN(num):
    return num != num


def str_to_NaN(s):
    if s == s:
        if re.search(r'^nan$', s, flags=re.IGNORECASE):
            s = np.nan
            return s
        else:
            return s
    else:
        return s



translated_name = 'default'
curr_path = os.path.abspath(os.getcwd())

filenames = os.listdir(curr_path)
for file in filenames:

    if re.search(r'draft', file, flags=re.IGNORECASE) and file.endswith('.xlsx') and file.count('~') == 0:
        translated_name = file
        draft = pd.read_excel(os.path.join(curr_path, file), header=0, dtype=str)
        columns_name = draft.columns.to_list()
        print('+++это драфт+++\n', draft.head())
        for num, col in enumerate(columns_name):
            if re.search(r'desc', col, flags=re.IGNORECASE):
                col_num = num

    elif re.search(r'transl', file, flags=re.IGNORECASE) and file.endswith('.xlsx') and file.count('~') == 0:
        dictionary = pd.read_excel(os.path.join(curr_path, file), header=None, dtype=str)
        dictionary.columns = ['source', 'target']
        dictionary.dropna(inplace=True)

        # проверка наличия дубликатов в словаре
        check_df_1 = dictionary[dictionary.duplicated(subset=['source'])]

        if len(check_df_1) != 0:
            print('Дубликаты в столбце оригинала:\n', check_df_1)
            resp = ''
            while not re.search(r'[yn]', resp, flags=re.IGNORECASE):
                resp = input('Удалить дубликаты в словаре и продолжить перевод (y/n)?')
            if re.search(r'y', resp, flags=re.IGNORECASE):
                dictionary.drop_duplicates(subset = ['source'], inplace=True)
                print('Дубликаты словаря удалены')

            else:
                sys.exit("Словарь содержит дубликаты")
        print('+++это словарь+++\n', dictionary.head())


merged = draft.merge(dictionary, how='left', left_on=draft[columns_name[col_num]], right_on='source')
merged.drop(columns='source', inplace=True)

# проверка наличия непереведенных позиций в description
description = merged[columns_name[col_num]].to_list()
translation = merged['target'].to_list()
not_translated = []
for item, translated in zip(description, translation):
    if not isNaN(item) and \
     isNaN(translated) and \
     item not in not_translated and \
     item.lower() != columns_name[col_num].lower():
        not_translated.append(item)
    elif item == translated:
        not_translated.append(item)

# перевод столбца заголовков, первого столбца
headers = merged[merged.columns[0]].tolist()
translated_headers = []
untranslated_headers = []
for item in headers:
    if dictionary['source'].isin([item]).any():
        translated_headers.append(dictionary[dictionary['source'] == item]['target'].values[0])
    else:
        if item == item and\
          re.search(r'[a-z]{3,}', item, flags=re.IGNORECASE) and \
          item.lower() != columns_name[col_num].lower():
            untranslated_headers.append(item)
        translated_headers.append(item)
translated_headers = list(map(str_to_NaN, translated_headers))
merged[merged.columns[0]] = translated_headers

# перевод столбца 'Comment'
for num, col in enumerate(merged.columns.to_list()):  # определение столбца коментариев
    if re.search(r'comment', col, flags=re.IGNORECASE):
        comment_num = num
if 'comment_num' in globals():
    comments = merged[merged.columns[comment_num]].to_list()
    translated_comments = []
    for item in comments:
        if dictionary['source'].isin([item]).any():
            translated_comments.append(dictionary[dictionary['source'] == item]['target'].values[0])
        else:
            translated_comments.append(item)
    translated_comments = list(map(str_to_NaN, translated_comments))
    merged['Comment'] = translated_comments

# вывод непереведенных заголовков (если есть)
if len(untranslated_headers) != 0 or len(not_translated) != 0:
    untranslated_headers.extend(not_translated)
    untranslated_headers = list(set(untranslated_headers))
    untranslated_headers = list(map(str_to_NaN, untranslated_headers))

    untranslated = pd.DataFrame({'untrnsl':untranslated_headers})
    untranslated.dropna(inplace=True)
    untranslated.to_excel(os.path.join(curr_path, translated_name[:-5] + '_untranslated.xlsx'),
                          header=False, index=False, na_rep='')

merged.to_excel(os.path.join(curr_path, translated_name[:-5] + '_translated.xlsx'), index=False, na_rep='')
print('====это непереведенные позиции====\n', set(untranslated_headers))
