import pandas as pd
import re
import os
import shutil

"""список бук кодов принимается в файле эксель в директории каталога 
поиск ЛО книг с буккодами из списка в директории search_dir
перемещение папок с ЛО книгами в move_dir
"""

curr_path = os.getcwd()
filenames = os.listdir(curr_path)
files = []
for file in filenames:
    if file.endswith('.xlsx') and file.count('~') == 0:
        files.append(file)
if len(files) != 1:
    raise Exception(f'Файл для обработки в пути {curr_path} не определен: Файлы: {files}')

# book codes for move
book_codes = files[0]
df = pd.read_excel(os.path.join(curr_path, book_codes), header=None, dtype=str)
if len(df.columns) != 1:
    raise Exception(f'file {book_codes} contains more than 1 column')

# abs paths please
search_dir = r'C:\Users\nekrasovkn\Desktop\Python38-32\PageRedactor\_mass_replace'
move_dir = r'C:\LinkOne\ready\Polyus\CAT'
search_dir = os.path.join(search_dir)
move_dir = os.path.join(move_dir)

cats_to_move_count = len(df)
moved_folders = []
for item in df[df.columns[0]]:
    # item = item.replace('.', r'\.')
    for folder in os.listdir(search_dir):
        if item == folder:   # re.search(item, folder, flags=re.I):
            shutil.move(os.path.join(search_dir, item), move_dir)
            moved_folders.append(item)
moved_folders_count = len(moved_folders)

if cats_to_move_count != moved_folders_count:
    print(f'folders to move : {cats_to_move_count}, moved: {moved_folders_count}')
    print(f'folders missed: {[i for i in df[df.columns[0]] if i not in moved_folders]}')
else:
    print(f'folders to move : {cats_to_move_count}, moved: {moved_folders_count}')
