import pandas as pd
import os
import sys

"""
слияние двух указанных колонок
таблицы
"""
COL1 = 'serial'
COL2 = 'remarks'

def ammend_cols(df, col1, col2):

    res = []
    for item1, item2 in zip(df[col1], df[col2]):
        if item1 == item1 and item2 == item2:
            res.append('[' + item1 + ']' + ' ' + item2)
        elif item1 != item1 and item2 == item2:
            res.append(item2)
        elif item1 == item1 and item2 != item2:
            res.append('[' + item1 + ']')
        elif item1 != item1 and item2 != item2:
            res.append(float('nan'))
    df[col2] = res


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

ammend_cols(df, COL1, COL2)

df.to_excel(file[:-5] + '_1.xlsx', index=False)
