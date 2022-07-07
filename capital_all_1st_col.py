import pandas as pd
import os
import sys

"""
Все строки в первой колонке сделать заглавными
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

df[df.columns[0]] = [i.upper() for i in df[df.columns[0]]]
df.to_excel(file[:-5] + '_allcap.xlsx', index=False, header=False)