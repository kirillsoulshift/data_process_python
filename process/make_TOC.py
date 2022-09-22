import pandas as pd
import os
import sys

"""
входной файл: файл словаря (2 колонки)

вывод: создание директорий по количеству строк входного файла
                      <0000a NA PAGE_NAME>
"""


def make_folder_num(num):

    if len(str(num)) == 1:
        num = ''.join(['000', str(num)])
    elif len(str(num)) == 2:
        num = ''.join(['00', str(num)])
    elif len(str(num)) == 3:
        num = ''.join(['0', str(num)])
    elif len(str(num)) == 4:
        num = str(num)
    return num


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

for n, name in enumerate(df[df.columns[1]]):
    dest_folder = make_folder_num(n + 1) + 'a' + ' NA ' + name[:10].strip()
    dest_path = os.path.join(curr_path, dest_folder)
    os.mkdir(dest_path)
    with open(os.path.join(dest_path, 'name.txt'), 'w') as f:
        f.write(name)  # полное наименование страницы записывается в текстовый файл в папке
