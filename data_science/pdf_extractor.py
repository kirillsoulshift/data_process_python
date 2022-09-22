import fitz
from PIL import Image
import io
import os
import re
import shutil
import pandas as pd
#todo проверка извлечения PermissionError: [WinError 32] Процесс не может получить доступ к файлу, так как этот
# файл занят другим процессом: 'C:\\Program Files\\Python\\PycharmProjects\\pythonProject\\app_test\\HD465-7Rtest
# изображения_не_перенесены\\pg0008im0001-A1310-002001.png' -> 'C:\\Program Files\\Python\\PycharmProjects\\
# pythonProject\\app_test\\HD465-7Rtest_изображения_не_востребованы\\pg0008im0001-A1310-002001.png'#
# During handling of the above exception, another exception occurred:
# Traceback (most recent call last):
# File "C:\Program Files\Python\Python38\lib\tkinter\__init__.py", line 1883, in __call__
#     return self.func(*args)
#   File "C:/Program Files/Python/PycharmProjects/pythonProject/oop/main.py", line 266, in set_border_value
#     message = self.extr.move_imgs(border)
#   File "C:\Program Files\Python\PycharmProjects\pythonProject\oop\pdf_extractor.py", line 160, in move_imgs
#     shutil.move(path, self.bad_image_folder)
#   File "C:\Program Files\Python\Python38\lib\shutil.py", line 803, in move
#     os.unlink(src)
# PermissionError: [WinError 32] Процесс не может получить доступ к файлу, так как этот файл занят другим процессом:
# 'C:\\Program Files\\Python\\PycharmProjects\\pythonProject\\app_test\\HD465-7Rtest_изображения_не_перенесены\\
# pg0008im0001-A1310-002001.png'

"""
Извлечение изображений pdf файла, отсеивание невостребованных изображений,
помещение востребованных изображений по папкам в директории каталога.

Если не выделяются цельные изображения:
pdf => Adobe Fine Reader => OCR-редактор => разметить области изображений и каталожных номеров => 
сохранить как pdf с возможностью поиска

Перемещение изображений в папку не востребованные происходит по соответствию группам разрещений.
Все извлеченные в ходе работы алгоритма изображения могут быть сгруппированы по разрешению.
Наиболее многочисленные группы формируют изображения не представляющие ценности. Это могут быть 
полностью черные изображения, неполные изображения, в зависимости от разновидности и способа формирования 
исходного pdf файла. Для работы алгоритма требуется указать число изображений в группе с самым многочисленным
содержанием изображений представляющими ценность. Те группы, число изображений в которых превышают указанное число,
будут отмечены как невостребованные.

помещает изображения по совпадению каталожного номера изображения (со  страницы на которой находилось изображение)
в одну из папок каталога с таким же каталожным номером в названии.
"""

def wrap_quot(s):
    return '"'+s+'"'

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


def clean_str(el):

    if re.search(r'[\\/:*?«<>|+=;."\n]+', el):
        el = re.sub(r'[\\/:*?«<>|+=;."\n]+', '', el)
    return el


class Extractor():
    def __init__(self, file_path, target_path):
        """
        :param file_path: Путь к PDF файлу
        :param target_path: Путь к целевой папке для выгрузки изображений
        """

        self.target_path = os.path.realpath(target_path)
        self.file_path = os.path.realpath(file_path)
        self.filename = os.path.split(file_path)[1][:-4]
        self.image_folder = f'{self.filename}_изображения_не_перенесены'
        self.image_path = os.path.join(self.target_path, self.image_folder)
        if os.path.exists(self.image_path):
            raise Exception(f'Папка {self.image_path} уже существует')
        os.mkdir(self.image_path)
        self.bad_image_folder = os.path.join(self.target_path, f'{self.filename}_изображения_не_востребованы')
        if os.path.exists(self.bad_image_folder):
            raise Exception(f'Папка {self.bad_image_folder} уже существует')
        os.mkdir(self.bad_image_folder)
        self.commons = None
        self.content = None

    def extract(self):

        # open file
        with fitz.open(self.file_path) as my_pdf_file:
            # loop through every page
            pdf_len = len(my_pdf_file)
            for page_number in range(1, pdf_len + 1):

                # access individual page
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
                    page_number_str = str(page_number).zfill(4)
                    image_number_str = str(image_number).zfill(4)
                    header_str = clean_str(header)
                    # save image locally
                    save_path = os.path.join(self.image_path,
                                             f"pg{page_number_str}im{image_number_str}-{header_str}.png")
                    image.save(open(save_path, "wb"), format='png')

        # определение часто встречающихся разрешений картинок
        self.content = os.listdir(self.image_path)
        dimensions = []
        for file in self.content:
            if file.endswith('.png'):
                path = os.path.join(self.image_path, file)
                with Image.open(path) as img:
                    dimensions.append(img.size)
        self.commons = pd.DataFrame({'dims': dimensions})
        print('Разрешение | Количество изображений\n', self.commons['dims'].value_counts())

        return f'Проверьте разрешения изображений в папке {self.image_path} и '\
               'в поле "Граница группировки" введите количество изображений, выше которого будут отсеяны '\
               'изображения в группах, количество изображений '\
               'в которых больше введенного числа.'

    @staticmethod
    def move_imgs(border: int, image_folder, bad_image_folder, directory_folder):
        """
        Перемещение изображений по индивидуальным папкам в директории каталога
        :param border: число выше которого отеиваются более многочисленные группы
        изображениий по разрешению
        :param image_folder: папка изображений
        :param bad_image_folder: папка невостребованных изображений
        :param directory_folder: директория каталога с эксель файлами
        :return: str
        """
        image_folder = os.path.realpath(image_folder)
        bad_image_folder = os.path.realpath(bad_image_folder)
        directory_folder = os.path.realpath(directory_folder)

        content = os.listdir(image_folder)
        dimensions = []
        for file in content:
            if file.endswith('.png'):
                path = os.path.join(image_folder, file)
                with Image.open(path) as img:
                    dimensions.append(img.size)
        commons = pd.DataFrame({'dims': dimensions})
        most_common_dims = commons['dims'].value_counts()[lambda x: x > border].index.tolist()  # разрешения

        # определение ненужных изображений
        files_to_delete = []
        for dim in most_common_dims:
            for file in content:
                if file.endswith('.png'):
                    path = os.path.join(image_folder, file)
                    with Image.open(path) as img:
                        if img.size == dim:
                            files_to_delete.append(file)

        # перемещение ненужных изображений
        for file in files_to_delete:
            path = os.path.join(image_folder, file)
            shutil.move(path, bad_image_folder) # вызывает os.unlink(src) PermissionError: [WinError 32] на
            # рандомные файлы
            # os.system(f'mv {wrap_quot(path)} {self.bad_image_folder}')

        # перемещение нужных изображений
        content = os.listdir(image_folder)
        all_pics_count = len(content)
        moved_pics_count = 0
        # только папки:
        dirs = [i for i in os.listdir(directory_folder) if os.path.isdir(os.path.join(directory_folder, i))]
        for pic in content:
            pic_name = '-'.join(pic.split('-')[1:]).rpartition('.')[0]  # название у pic после первого "-" и до ".png"
            pic_name = pic_name.strip(' .,')
            for folder in dirs:
                if re.search(pic_name, folder, flags=re.IGNORECASE):
                    shutil.move(os.path.join(image_folder, pic), os.path.join(directory_folder, folder))
                    # os.system(f'mv {wrap_quot(os.path.join(self.image_path, pic))} '
                    #           f'{wrap_quot(os.path.join(self.directory_path, folder))}')
                    moved_pics_count += 1
                    break
        return f'Перемещено {moved_pics_count} изображений из {all_pics_count}\n'\
               f'Папка каталога {directory_folder}\n'\
               f'{all_pics_count - moved_pics_count} файлов в папке {image_folder}.'
