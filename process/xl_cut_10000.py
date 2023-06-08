import pandas as pd
# import re
import os
import sys

"""
Разделить таблицу .xlsx на множество
таблиц с количеством позиций в
переменной 'quantifier'
"""

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

positions = len(df)
quantifier = 6501
for n in range((positions // quantifier) + 1):
    start = n*quantifier
    end = (n + 1)*quantifier
    if end < positions:
        df_to_write = df.iloc[start:end].copy()
        df_to_write.to_excel(os.path.join(curr_path, file[:-5] + f'_{n}.xlsx'), header=False, index=False)

    else:
        end = positions
        df_to_write = df.iloc[start:end].copy()
        df_to_write.to_excel(os.path.join(curr_path, file[:-5] + f'_{n}.xlsx'), header=False, index=False)
