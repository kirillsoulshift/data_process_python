from data.crushDecompressor import CrushDecompressor
from data.bbiredactor import Redactor
import pandas as pd

PATH = os.getcwd()
decompressor = CrushDecompressor()
bbiredactor = Redactor()

def folder_surf(path, dest_path):
    content = os.listdir(path)
    for item in content:
        if os.path.isdir(os.path.join(path, item)):
            folder_surf(os.path.join(path, item), dest_path)


def get_lo_path(cat_path):
    files = []
    for file in os.listdir(cat_path):
        if (file.endswith('.bbi') or file.endswith('.BBI')) and file.count('~') == 0:
            file = os.path.realpath(file).replace('\\', '/')
            files.append(file)
    if len(files) == 0:

        raise Exception(f'Путь {cat_path} не содержит .bbi')
    return files

folders = []
for item in os.listdir(PATH):
    if os.path.isdir(item) and item.count('test'):
        folders.append(item)
if len(folders) != 1:

    raise Exception(f'Путь {PATH} не содержит директории test')

for folder in folders:
    pixes = get_lo_path(folder)
    print(pixes)
    for path in pixes:

        temp_bbi_path, temp_size = decompressor.decompress_bbi(path)

bbiredactor.read_file(temp_bbi_path)
structure = bbiredactor.get_struct()
book_struct = bbiredactor.get_book_bbi(structure)

for k, v in book_struct:
    print(f'====={k}=====')
    print(v)





# with open(new_file_path, 'rb') as f:
#     file = f.read()
#     print(file.decode())



