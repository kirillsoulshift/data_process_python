import pandas as pd
import re
import os
import sys


"""
разделяет файл .xlsx на два файла .xlsx по содержанию
строк на русском и английском языках
проверка по первой колонке
скрипт формирует два новых файла в своей директории
"""

def is_english(s):
    if s != s:
        return True
    else:
        if re.search(r'[а-я]', s, flags=re.IGNORECASE):
            return False
        else:
            return True


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
rus = []
eng = []
for item in df[df.columns[0]]:
    if is_english(item):
        eng.append(item)
    else:
        rus.append(item)
rus_df = pd.DataFrame(rus)
eng_df = pd.DataFrame(eng)
rus_df.to_excel(file[:-5] + '_rus.xlsx', header=False, index=False)
eng_df.to_excel(file[:-5] + '_eng.xlsx', header=False, index=False)
