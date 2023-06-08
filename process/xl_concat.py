import pandas as pd
import os

"""
Объединение всех .xlsx файлов в директории скрипта
скрипт формирует объединенный файл .xlsx в своей директории
"""

curr_path = os.getcwd()
filenames = os.listdir(curr_path)
files = []
for file in filenames:
    if file.endswith('.xlsx') and file.count('~') == 0:
        files.append(file)

for n, file in enumerate(files):
    if n == 0:
        df = pd.read_excel(os.path.join(curr_path, file), header=None, dtype=str)
    else:
        df = df.append(pd.read_excel(os.path.join(curr_path, file), header=None, dtype=str))

df.to_excel(os.path.join(curr_path, 'concated_xl.xlsx'), header=False, index=False)
