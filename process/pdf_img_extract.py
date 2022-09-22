import fitz
from PIL import Image
import io
import os
import re
import sys
import shutil
import pandas as pd


#todo проверка наличия дублей изображений. удаление дублей
"""
если не выделяет цельные изображения:
pdf => OCR => разметить области изображений и заголовков => сохранить как pdf с возможностью поиска

обрабатывает файл .pdf в папке, в которой находится скрипт
помещает все изображения из файла .pdf в новую папку в директории скрипта

- разметка картинок в ОКР по шаблону приводит к однообразию разрешений картинок (problem?)
- картинки извлекаются в ориентации пдф файла(книжная или альбомная) (problem?) (установка нужной ориентации в hotpoint)

если в директории скрипта лежит папка с каталогом (название папки = название pdf)
(в форме множества папок)-
помещает изображения по совпадению названия/кат. номера изображения
в одну из папок каталога
"""


def get_catnum_or_name(s):
    tokens = s.split('\n')
    catnum = ''
    for token in tokens:
        if re.search(r'[A-Z]?[\d-]{5,}[A-Z]?', token):
            m = re.search(r'[A-Z]?[\d-]{5,}[A-Z]?', token)
            catnum = m.group(0).replace(' ', '')
            break
    if catnum:
        return catnum
    else:
        return 'UNDEFINED'


def make_num(num):

    if len(str(num)) == 1:
        num = ''.join(['000', str(num)])
    elif len(str(num)) == 2:
        num = ''.join(['00', str(num)])
    elif len(str(num)) == 3:
        num = ''.join(['0', str(num)])
    elif len(str(num)) == 4:
        num = str(num)
    return num


def clean_str(el):

    if re.search(r'[\\/:*?«<>|+=;."\n]+', el):
        el = re.sub(r'[\\/:*?«<>|+=;."\n]+', '', el)

    return el


curr_path = os.getcwd()
filenames = os.listdir(curr_path)
pdf_list = []
for file in filenames:
    if file.endswith('.pdf') and file.count('~') == 0:
        pdf_list.append(file)
if len(pdf_list) != 1:
    print(pdf_list)
    sys.exit("Количество файлов pdf для обработки не = 1")

filename = pdf_list[0]
image_folder = f'{filename[:-4]}_изображения_не_перенесены'
image_path = os.path.join(curr_path, image_folder)
os.mkdir(image_path)

# open file
with fitz.open(filename) as my_pdf_file:
    # loop through every page
    pdf_len = len(my_pdf_file)
    for page_number in range(1, pdf_len + 1):

        # acess individual page
        page = my_pdf_file[page_number - 1]
        text = page.get_text('text', sort=True)
        header = get_catnum_or_name(text)

        # accesses all images of the page
        images = page.get_images()

        # check if images are there
        if images:
            print(f"{len(images)} изображений на стр {page_number}/{pdf_len}[+]", f'CAT_NUM = {header}')
        else:
            print(f"0 изображений на стр {page_number}[!]")

        # loop through all images present in the page
        for image_number, image in enumerate(images, start=1):
            # access image xref
            xref_value = image[0]

            # extract image information
            base_image = my_pdf_file.extract_image(xref_value)

            # access the image itself
            image_bytes = base_image["image"]

            # get image extension
            # ext = base_image["ext"]

            # load image
            image = Image.open(io.BytesIO(image_bytes))
            page_number_str = make_num(page_number)
            image_number_str = make_num(image_number)
            header_str = clean_str(header)
            # save image locally
            save_path = os.path.join(image_folder, f"pg{page_number_str}im{image_number_str}-{header_str}.png")
            image.save(open(save_path, "wb"), format='png')

# ================= удаление ненужных изображений============
# папка для перемещения ненужных изображений
bad_image_folder = os.path.join(curr_path, f'{filename[:-4]}_изображения_не_перенесены_2')
os.mkdir(bad_image_folder)

# определение часто встречающихся разрешений картинок
content = os.listdir(image_path)
dimensions = []
for file in content:
    if file.endswith('.png'):
        path = os.path.join(image_path, file)
        with Image.open(path) as img:
            dimensions.append(img.size)
commons = pd.DataFrame({'dims' : dimensions})
print('Разрешение | Количество изображений\n', commons['dims'].value_counts())

border = ''
while type(border) != int:
    try:
        border = int(input(f'Проверьте папку {image_folder} и '
                           f'введите количество изображений выше которого будут отсеяны'
                            ' все изображения с повторяющимся разрешением:'))
        break
    except ValueError:
        print('Необходимо ввести число')
print('Moving pictures...')
most_common_dims = commons['dims'].value_counts()[lambda x: x > border].index.tolist()  # разрешения картинок

# определение ненужных изображений
files_to_delete = []
for dim in most_common_dims:
    for file in content:
        if file.endswith('.png'):
            path = os.path.join(image_path, file)
            with Image.open(path) as img:
                if img.size == dim:
                    files_to_delete.append(file)

# перемещение ненужных изображений
for file in files_to_delete:
    path = os.path.join(image_path, file)
    shutil.move(path, bad_image_folder)

# перемещение нужных изображений
# поиск папки с каталогом в форме !плоской! директории
cat_dir = []
for file in filenames:
    if filename[:-4] == file and\
    os.path.isdir(os.path.join(curr_path, file)):  # папка с каталогом должна называться в точности как файл pdf
        cat_dir.append(file)
if len(cat_dir) != 1:
    print(f'Директории каталога: {cat_dir}')
    sys.exit(f"Не определена директория каталога, изображения находятся в папке {image_folder}")

move_dest = os.path.join(curr_path, cat_dir[0])  # директория с каталогом
image_path = os.path.join(curr_path, image_folder)
content = os.listdir(image_path)
all_pics_count = len(content)
moved_pics_count = 0

dirs = [i for i in os.listdir(move_dest) if os.path.isdir(os.path.join(move_dest, i))]  # только папки
for pic in content:
    pic_name = '-'.join(pic.split('-')[1:]).rpartition('.')[0]  # название у pic после первого "-" и до ".png"
    pic_name = pic_name.strip(' .,')
    for folder in dirs:
        if re.search(pic_name, folder, flags=re.IGNORECASE):
            shutil.move(os.path.join(image_path, pic), os.path.join(move_dest, folder))
            moved_pics_count += 1
            break
print(f'Перемещено {moved_pics_count} изображений из {all_pics_count}\n'
      f'Папка каталога {move_dest}\n'
      f'{all_pics_count - moved_pics_count} файлов в папке {image_path}')
