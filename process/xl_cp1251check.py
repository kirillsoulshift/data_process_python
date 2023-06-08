import pandas as pd
import os
import sys


"""
проверка кодируемости строк кодеком cp1251 во второй колонке
"""

def check_encode(item):
    if item == item:
        try:
            item.encode('cp1251')
            return float('nan')
        except UnicodeEncodeError:
            return item
    else:
        return item


curr_path = os.getcwd()
filenames = os.listdir(curr_path)
files = []
for file in filenames:
    if file.endswith('.xlsx') and file.count('~') == 0:
        files.append(file)
if len(files) != 1:
    print(f'Файлы: {files}')
    sys.exit('Файл для обработки не определен')

file = files[0]
df = pd.read_excel(os.path.join(curr_path, file), header=None, dtype=str)

df[df.columns[1]] = df[df.columns[1]].apply(check_encode)
df.dropna(inplace=True)
if len(df) != 0:
    print(df)
else:
    print('Некодируемых позиций не найдено')
