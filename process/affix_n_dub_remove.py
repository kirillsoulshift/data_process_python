import pandas as pd
import re
import os
import sys

"""
обрабатывает файл .xlsx в директории скрипта
с совпадением 'to_translate'

удаление пустых строк, дубликатов
создание второго столбца с терминами
без дефисов в начале строк (если нужно)
"""

path = os.getcwd()
files = os.listdir(path)
process_files = []
for file in files:
    if re.search('to_translate', file, flags=re.IGNORECASE) and \
            file.endswith('.xlsx') and \
            file.count('~') == 0:
        process_files.append(file)
if len(process_files) != 1:
    print(f'Файлы: {files}')
    sys.exit('Файл для обработки не определен')

df = pd.read_excel(os.path.join(path, process_files[0]), header=None, dtype=str)
df.columns = ['source']
df.dropna(inplace=True)
df.drop_duplicates(df.columns[0], keep='last', inplace=True)
afx_count = 0

for item in df['source']:
    if re.search(r'^[\s-]', item):
        afx_count += 1

if afx_count != 0:
    df['afx_less'] = df['source'].apply(lambda x: x.strip(' -'))

df.to_excel(os.path.join(path, process_files[0]), header=False, index=False)