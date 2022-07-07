import pandas as pd
import os
import sys
import re

"""
Формирует .xlsx из .tmx находящегося в директории скрипта
результирующий файл содержит колонки source | target | date
"""

curr_path = os.getcwd()
filenames = os.listdir(curr_path)
files = []
for file in filenames:
    if file.endswith('.tmx') and file.count('~') == 0:
        files.append(file)
if len(files) != 1:
    print(f'Файлы: {files}')
    sys.exit('Файл tmx для обработки не определен')

filename = files[0]
all = []
source = []
target = []
dates =[]
with open(filename, 'rb') as file:
    while line := file.readline():
        pr_string = line.decode().strip()
        if pr_string.count('<tu creationdate'):
            m = re.search(r'changedate=".+"', pr_string)
            date = re.sub(r'[a-zA-Z="]+', '', m.group(0))[:8]  # changedate="20200205T090240Z" -> 20200205
            dates.append(date)

        elif pr_string.count('<seg>', 0, 5):
            pr_string = pr_string[5:-6]
            all.append(pr_string)

for n, item in enumerate(all):
    if n % 2 == 0:
        source.append(item)
    else:
        target.append(item)


df = pd.DataFrame({'source' : source, 'target' : target, 'date' : dates})
df.astype('str')

df.to_excel('tmx_load.xlsx', header=False, index=False)
