import pandas as pd
import os
import sys
import re

"""
Все строки в первой (второй) колонке сделать заглавными
удалить китайские иероглифы
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

RE = re.compile(u'[⺀-⺙⺛-⻳⼀-⿕々〇〡-〩〸-〺〻㐀-䶵一-鿃豈-鶴侮-頻並-龎]', re.UNICODE)
df[df.columns[1]] = [RE.sub('', i).strip() for i in df[df.columns[1]]] # удаление во второй колонке

# df[df.columns[1]] = df[df.columns[1]].apply(lambda x: x.upper())

df.to_excel(file[:-5] + '_nonchinese.xlsx', index=False, header=False)