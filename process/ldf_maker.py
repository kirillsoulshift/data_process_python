import pandas as pd
import re
import os
import sys
import shutil
from PIL import Image
import time


"""
ПРИМЕР директории в которую помещается скрипт
_____________________________________________________________________________
0000 PV4354 PV-275
|_0001 57793226 КОРПУС ЗАЩИТНЫЙ С ГЕНЕРАТОРНОЙ УСТАНОВКОЙ
| |_КОРПУС ЗАЩИТНЫЙ С ГЕНЕРАТОРНОЙ УСТАНОВКОЙ.xlsx
|_0002 57730574 УСТАНОВКА ПЛАТФОРМЫ
| |_1.jpg
| |_2.jpg
| |_3.jpg
| |_УСТАНОВКА ПЛАТФОРМЫ.xlsx
| |_0003 57739377 ПЛАТФОРМА С ПЕРИЛАМИ, ПРОХОД К КАБИНЕ
|   |_ПЛАТФОРМА С ПЕРИЛАМИ, ПРОХОД К КАБИНЕ.xlsx
|   |_3.jpg
| |_0004 57739468 ПЛАТФОРМА С ПЕРИЛАМИ
|   |_ПЛАТФОРМА С ПЕРИЛАМИ.xlsx
|   |_4.jpg
| |_0005 57739526 ПЛАТФОРМА С ПЕРИЛАМИ
|   |_ПЛАТФОРМА С ПЕРИЛАМИ.xlsx
|   |_5.png
|   |_5-1.png
|_0008 57501678 ЛЕСТНИЦА
| |_ЛЕСТНИЦА.xlsx
| |_8.png
|_<скрипт>

первая строка файлов .xlsx = наименования столбцов
порядок столбцов файлов .xlsx:

 ================================================================================================================
           >>>'Item ID', 'Part Number', 'Quantity', 'Unit', 'Description', 'Translation', 'Comment'<<<  
 ================================================================================================================
 
в директориях находятся по одному файлу .xlsx
и файлы изображений, такие как: .jpeg, .jpg, .png, .bmp, .gif

создает .ldf файл для каждой директории
прикрепляет все изображения, номера папок в качестве ссылок
к созданному .ldf

помещает созданные .ldf файлы с копиями
изображений в новую директорию 'compiled_ldf',
которая создается в родительской директории скрипта

папка 'compiled_ldf' НЕ будет содержать метаданных для построения
книги в Book Builder (это файлы .inf, .fmt, .sdf)
"""


def count_and_trim(el):
    """
    Обрезает строку наименования ldf файла (title)
    т.к. ее длина не может превышать 128 Б
    :param el: str
    :return: str
    """
    letters = re.sub(r'[^А-Я]', '', el)
    misc = re.sub(r'[А-Я]', '', el)
    if len(letters) * 2 + len(misc) > 128:
        if el.count('('):
            return count_and_trim(el.rpartition('(')[0].strip())
        elif el.count(' ') and not el.count('('):
            return count_and_trim(el.rpartition(' ')[0].strip())
        else:
            return el[:len(el)//2].strip()
    else:
        return el.strip()


def get_name(path):
    """
    поиск файла txt с названием в текущей папке
    возврат текста отредактированного по длине названия страницы
    :param: путь к текущей директории
    :return: str
    """
    try:
        with open(os.path.join(path, 'name.txt')) as f:
            first_line = f.readline()
        return count_and_trim(first_line)
    except FileNotFoundError:
        sys.exit(f'В пути {path} не обнаружен текстовый файл <name.txt> с названием страницы')


def make_catnum(s):
    if s == 'NA':
        return ''
    else:
        return s


def entry(s):
    if s.count('\n'):
        s = s.replace('\n', ' ')
    if re.search(r'\s{2,}', s):
        s = re.sub(r'\s{2,}', ' ', s)
    if s.count('"'):
        s = s.replace('"', "''")
    return '"' + s + '"'


def get_pics(path):

    pics = []
    files = os.listdir(path)

    for file in files:
        if (file.endswith('.jpeg') or
          file.endswith('.jpg') or
          file.endswith('.png') or
          file.endswith('.bmp') or
          file.endswith('.gif')):

            ends = file[file.rfind('.'):]

            definitions = []

            try:
                im = Image.open(os.path.join(path, file).replace('\\', '/'))
                definitions = [im.width, im.height]
            except Exception as e:
                definitions = [1000, 1000]

            pic = [os.path.join(path, file).replace('\\', '/'),
                   ends,
                   definitions]
            pics.append(pic)
    return pics


def entity_define(path):
    """
    определяет количество файлов
    и папок в директории

    :param path: pathlike object
    :return: dict
    """

    content = os.listdir(path)
    ent_count = {'files': 0, 'folders': 0}
    for entity in content:
        if entity.endswith('.xlsx') and entity.count('~') == 0:
            ent_count['files'] += 1
        elif os.path.isdir(os.path.join(path, entity)):
            ent_count['folders'] += 1
    if ent_count['files'] > 1:
        print(path)
        sys.exit('Более одного файла .xlsx в папке')
    return ent_count


def get_ref(folder_path, cat_num:str):
    """
    возвращает номер ссылки по совпадению каталожного номера
    :param folder_path: pathlike object
    :param cat_num: str
    :return: str
    """
    content = os.listdir(folder_path)
    reference_list = []
    for item in content:
        if re.search(cat_num, item):
            reference_list.append(item.split(' ')[0])  # номер вначале "0007 57711103 УСТАНОВКА ПЕРИЛ"
    if len(reference_list) == 0:
        return ''
    elif len(reference_list) == 1:
        return reference_list[0]
    else:
        print(folder_path)
        sys.exit("Дублирование каталожных номеров в папке")


def check_units(row):
    """
    проверка наличия десятичных дробей
    и букв обозначений единиц измерения в графе
    количество
    :param row: list
    :return: list
    """

    if re.search(r'[.,]', row[2]): # если содержатся десятичные дроби
        new_row = []
        qty = row[2]
        unit = row[3]

        for n, item in enumerate(row):
            if n == 2:
                new_row.append('1')
            elif n == 3:
                new_row.append('EA')
            else:
                new_row.append(item)
        if len(row) == 6:
            new_row.append(qty + ' ' + unit)
            return new_row
        elif len(row) == 7:
            comment = ' '.join([row[-1], qty, unit])
            new_row = new_row[:-1]
            new_row.append(comment)
            return new_row
        else:  # len(row) == 5
            new_row.append('')
            new_row.append(qty + ' ' + unit)
            return new_row

    elif re.search(r'[a-zа-я]', row[2], flags=re.IGNORECASE): # если содержатся сокращения единиц измерения
        new_row = []
        qty = re.sub(r'[^0-9]', '', row[2])
        unit = re.sub(r'[0-9]', '', row[2])

        for n, item in enumerate(row):
            if n == 2:
                new_row.append('1')
            else:
                new_row.append(item)
        if len(row) == 6:
            new_row.append(qty + ' ' + unit)
            return new_row
        elif len(row) == 7:
            comment = ' '.join([row[-1], qty, unit]).strip()
            new_row = new_row[:-1]
            new_row.append(comment)
            return new_row
        else:  # len(row) == 5
            new_row.append('')
            new_row.append(qty + ' ' + unit)
            return new_row
    else:
        return row


def write_ldf(path, dest_path, content_dict):
    """
    формирование, заполнение ldf на основании path
    помещение ссылок по номерам папок находящихся в path
    помещение ссылок на все изображения находящихся в path

    :param path: pathlike object это папка
    :param dest_path: pathlike object это путь помещения сформированного ldf
    :param content_dict: dict это словарь с количеством файлов и папок в директории
    :return: None
    """
    ldf_name = path.split('\\')[-1].split(' ')[0]  # "0007 57711103 УСТАНОВКА ПЕРИЛ" ==> "0007"
    page_name = get_name(path)  # извлечение названия из файла txt из текущей папки
    if content_dict['files'] != 0:  # если ldf файл делается из .xlsx
        content = os.listdir(path)
        xlsx_list = []
        for item in content:
            if item.endswith('.xlsx') and item.count('~') == 0:
                xlsx_list.append(item)
        if len(xlsx_list) != 1:
            print(path)
            sys.exit("Неверное количество файлов .xlsx")

        raw_file = pd.read_excel(os.path.join(path, xlsx_list[0]), header=0, dtype=str)
        raw_file.fillna('', inplace=True)

    pics = get_pics(path)
    ldf_path = os.path.join(dest_path, ''.join([ldf_name, '.ldf']))
    folders_content = [i for i in os.listdir(path) if os.path.isdir(os.path.join(path, i))]  # только папки

    with open(ldf_path, 'wb') as ldf:

        ldf.write(b'VERSION,05.100\n')  # первые две строки ldf одинаковые для всех файлов
        ldf.write(b'LIST,')
        ldf.write(ldf_name.encode('cp1251'))
        ldf.write(b',')
        ldf.write(ldf_name.encode('cp1251'))
        ldf.write(b',')
        ldf.write(entry(page_name).encode('cp1251'))
        ldf.write(b',')
        ldf.write(str(len(pics)).encode('cp1251'))
        ldf.write(b',')
        ldf.write(entry(page_name).encode('cp1251'))
        ldf.write(b',')
        ldf.write(b',\n')

        if content_dict['files'] != 0:  # если ldf файл делается из .xlsx
            reference_count = 0

            #     0            1            2          3          4              5            6
            # 'Item ID', 'Part Number', 'Quantity', 'Unit', 'Description', 'Translation', 'Comment'
            for n, row in enumerate(raw_file.values.tolist()):
                row = check_units(row)
                row = list(map(entry, row))

                ldf.write(b'ENTRY,')
                ldf.write(row[0].encode('cp1251'))
                ldf.write(b',')
                ldf.write(make_catnum(row[1]).encode('cp1251'))
                ldf.write(b',')
                ldf.write(row[4].encode('cp1251'))
                ldf.write(b',')
                ldf.write(row[2].encode('cp1251'))
                ldf.write(b',,,')
                ldf.write(str(n + 1).encode('cp1251'))                        # столбец №7 в ldf = порядк. номер строки
                ldf.write(b',,,,')                                            # что используется для ссылок на строку %n
                if content_dict['folders'] != 0:                              # если директория содержит папки
                    ldf.write(get_ref(path, row[1].strip(' "')).encode('cp1251'))  # в ldf прописывается ссылка
                    reference_count += 1
                ldf.write(b',')
                ldf.write(row[3].encode('cp1251'))
                ldf.write(b',\n')
                if len(row) >= 6 and len(row[5]) != 0:
                    ldf.write(b'FIELD,Translate,')
                    ldf.write(row[5].encode('cp1251'))
                    ldf.write(b'\n')
                if len(row) == 7 and len(row[6]) != 0:
                    ldf.write(b'FIELD,Comment,')
                    ldf.write(row[6].encode('cp1251', 'ignore'))  # попадание в строку Comment символов типа 0xc4
                    ldf.write(b'\n')

            # здесь дописываются ссылки, на которые не нашлись совпадения по партномерам
            if content_dict['folders'] != reference_count:
                partn_in_folder = list(map(lambda x: x.split(' ')[1], folders_content))
                partn_in_file = raw_file[raw_file.columns[1]].tolist()
                partn_to_write = []
                refs_to_write = []
                for folder in partn_in_folder:
                    if folder not in partn_in_file:
                        partn_to_write.append(folder)
                for folder in folders_content:
                    for partn in partn_to_write:
                        if re.search(partn, folder):
                            refs_to_write.append(folder)

                #        0                 1               2
                # 'folder number', 'catalog number', 'description',
                for n, folder in enumerate(refs_to_write):
                    ldf.write(b'ENTRY,,')
                    ldf.write(make_catnum(folder.split(' ')[1]).encode('cp1251'))
                    ldf.write(b',,')
                    ldf.write(b'1')
                    ldf.write(b',,,')
                    ldf.write(str(n + 1).encode('cp1251'))
                    ldf.write(b',,,,')
                    ldf.write(folder.split(' ')[0].encode('cp1251'))  # в ldf прописывается ссылка
                    ldf.write(b',,\n')
                    ldf.write(b'FIELD,Translate,')
                    ldf.write(entry(get_name(os.path.join(path, folder))).encode('cp1251'))
                    ldf.write(b'\n')

        else:  # если ldf файл делается из папки

            #        0                 1               2
            # 'folder number', 'catalog number', 'description',
            for n, folder in enumerate(folders_content):
                ldf.write(b'ENTRY,,')
                ldf.write(make_catnum(folder.split(' ')[1]).encode('cp1251'))
                ldf.write(b',')
                ldf.write(entry(get_name(os.path.join(path, folder))).encode('cp1251'))
                ldf.write(b',')
                ldf.write(b'1')
                ldf.write(b',,,')
                ldf.write(str(n + 1).encode('cp1251'))
                ldf.write(b',,,,')
                ldf.write(folder.split(' ')[0].encode('cp1251'))  # в ldf прописывается ссылка
                ldf.write(b',,\n')

        # копирование и переименовывание всех изображений, относящихся к данному ldf
        for pic_num, pic in enumerate(pics):
            # ссылка на изображение (без расширения)
            new_picname = ldf_name + '-' + str(pic_num)

            # ссылка на изображение (с расширением)
            new_picpath = ''.join([new_picname, pic[1]])
            new_picpath = os.path.join(dest_path, new_picpath)

            shutil.copyfile(pic[0], new_picpath)

            ldf.write(b'PICTURE,')
            ldf.write(str(pic_num + 1).encode('cp1251'))
            ldf.write(b',')
            ldf.write(new_picname.encode('cp1251'))
            ldf.write(b',,')
            ldf.write(b'0,1,1,')
            ldf.write(''.join([new_picname, pic[1]]).encode('cp1251'))
            ldf.write(b',,,\n')

            ldf.write(b'PICTUREFRAME,')
            ldf.write(str(pic_num + 1).encode('cp1251'))
            ldf.write(b',0,0,')
            ldf.write(str(pic[2][0]).encode('cp1251'))
            ldf.write(b',')
            ldf.write(str(pic[2][1]).encode('cp1251'))
            ldf.write(b'\n')


def folder_surf(path, dest_path):
    content = os.listdir(path)
    ent_count = entity_define(path)
    write_ldf(path, dest_path, ent_count)
    for item in content:
        if os.path.isdir(os.path.join(path, item)):
            folder_surf(os.path.join(path, item), dest_path)


start_time = time.time()
curr_path = os.getcwd() # текущая директория
# проверка длины путей файлов
path_to_review = []
for root, dirs, files in os.walk(curr_path):
    for item in files:
        if (item.endswith('.xlsx') or item.endswith('.png')) and item.count('~') == 0:
            item_path = os.path.abspath(item)
            path_length = len(item_path)
            if path_length >= 260:
                path_to_review.append(item_path)
if path_to_review:
    print('\033[91m' + 'Пути файлов превышают длину 260 символов:' + '\033[0m')
    for item in path_to_review:
        print('\033[91m' + f'{item}' + '\033[0m')
    sys.exit('Слишком длинный путь')

res_folder_name = os.path.split(curr_path)[1].split()[-1]        # название папки со скриптом (0000 NA HD1500 -> HD1500)
parent_path = os.path.join(curr_path, os.pardir)                 # родительская директория
dest_path = os.path.join(parent_path, res_folder_name + '_ldf')  # путь к формируемой директории
os.mkdir(dest_path)                                              # создание формируемой директории
folder_surf(curr_path, dest_path)

print("--- %s seconds ---" % (time.time() - start_time))
