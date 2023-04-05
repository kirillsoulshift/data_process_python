import pandas as pd
import re
import os
import sys

"""
обрабатывает файл .xlsx в директории скрипта


создание второго столбца с терминами
без дефисов в начале строк (если нужно)
"""

path = os.getcwd()
files = os.listdir(path)
process_files = []
for file in files:
    if file.endswith('.xlsx') and file.count('~') == 0:
        process_files.append(file)
if len(process_files) != 1:
    print(f'Файлы: {files}')
    sys.exit('Файл для обработки не определен')

df = pd.read_excel(os.path.join(path, process_files[0]), header=None, dtype=str)

if len(df.columns) == 1:
    df.columns = ['source']

    afx_count = 0

    for item in df['source']:
        if re.search(r'^[\s-]', item):
            afx_count += 1
            break

    if afx_count != 0:
        df['afx_less'] = df['source'].apply(lambda x: x.strip(' -'))

    df.to_excel('afxless.xlsx', header=False, index=False)
