import pandas as pd
import re
import os
import sys


"""
создает файлы для перевода
файл заголовков
файл позиций спецификации
"""


def remove_index(ser):
    """
    удалить индексы позиций и партномера
    :param ser: Series
    :return: Series
    """
    ser.reset_index(drop=True, inplace=True)
    idx_to_remove = []
    for n, item in enumerate(ser):
        if re.search(r'^\w?[\d-]+\w?$', item):
            idx_to_remove.append(n)
    ser.drop(idx_to_remove, inplace=True)

    return ser


curr_path = os.getcwd()
filenames = os.listdir(curr_path)
files = []
for file in filenames:
    if file.endswith('.xlsx') and file.count('~') == 0:
        files.append(file)
if len(files) != 1:
    print(f'Файлы: {files}')
    sys.exit(f'Файл для обработки в пути {curr_path} не определен')

file = files[0]
df = pd.read_excel(os.path.join(curr_path, file), header=0, dtype=str)
columns_name = df.columns
test_check = ''.join(columns_name)
if not re.search(r'desc', test_check, flags=re.IGNORECASE):
    print('\033[91m' + 'Столбцы исходного файла должны иметь названия схожие с:' + '\033[0m')
    print('\033[91m' + 'Столбец описания = Description' + '\033[0m')
    print('\033[91m' + 'Первый столбец в таблице = Заголовки и номера позиций' + '\033[0m')
    sys.exit()

headers = df[columns_name[0]]
headers.dropna(inplace=True)
headers.drop_duplicates(inplace=True)
headers = remove_index(headers)

desc_col = ''
for col in columns_name:
    if re.search(r'desc', col, flags=re.IGNORECASE):
        desc_col = col

if desc_col:
    rows = df[desc_col]
    rows.dropna(inplace=True)
    rows.drop_duplicates(inplace=True)
else:
    print('\033[91m' + 'Столбец описания не найден' + '\033[0m')
    sys.exit()

headers.to_excel(file[:-5] + '_headers.xlsx', index=False)
rows.to_excel(file[:-5] + '_rows.xlsx', index=False)
headers.append(rows).to_excel(file + '_all.xlsx', index=False)
