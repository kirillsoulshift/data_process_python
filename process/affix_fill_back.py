import pandas as pd
import re
import os
import sys


"""
слияние файлов .xlsx в директории скрипта (с одинаковым кол-вом позиций)
с совпадением 'to_translate' - файл с 2 колонками с дефисами и без
с совпадением 'translated' - файл словарь с 2 колонками

подстановка к колонке перевода дефисов в
тех ячейках где это нужно

формирование нового файла .xlsx в директории скрипта:
 - с исходными позициями каталога (с дефисами);
 - с переводом позиций (с дефисами).
"""


def what_prefix(df, n):
    prefix = ''
    for let in df.iloc[n, 0]:
        if re.search(r'-|\s', let):
            prefix = prefix + let
        else:
            break
    return prefix


path = os.getcwd()
files = os.listdir(path)
translated_files = []
for file in files:
    if re.search('translated', file, flags=re.IGNORECASE) and \
            file.endswith('.xlsx') and \
            file.count('~') == 0:
        translated_files.append(file)
if len(translated_files) != 1:
    print(f'Файлы: {files}')
    sys.exit('Файл для обработки не определен')

process_files = []
for file in files:
    if re.search('to_translate', file, flags=re.IGNORECASE) and \
            file.endswith('.xlsx') and \
            file.count('~') == 0:
        process_files.append(file)
if len(process_files) != 1:
    print(f'Файлы: {files}')
    sys.exit('Файл для обработки не определен')

translation = pd.read_excel(os.path.join(path, translated_files[0]), header=None, dtype=str)
source = pd.read_excel(os.path.join(path, process_files[0]), header=None, dtype=str)
source.columns = ['source', 'afxless']
source['target'] = translation.iloc[:, 1]
discrepancy = []
for i in source.index:
    if source.iloc[i, 0] != source.iloc[i, 1]:
        discrepancy.append(i)
for i in discrepancy:
    if source.iloc[i, 2] == source.iloc[i, 2]:
        source.iloc[i, 2] = what_prefix(source, i) + source.iloc[i, 2]

source.drop(['afxless'], axis=1, inplace=True)

source.to_excel(os.path.join(path, process_files[0][:-5] + 'DICT.xlsx'), header=False, index=False)
