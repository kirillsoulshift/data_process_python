import fitz
from PIL import Image
import io
import os
import sys
import shutil
import pandas as pd
from cells import *
from multiprocessing import Pool, cpu_count
import time
"""
Извлечение изображений pdf файла, отсеивание невостребованных изображений,
помещение востребованных изображений по папкам в директории каталога.

Если не выделяются цельные изображения:
Adobe Fine Reader / OCR-редактор:
1) разметить области каталожных номеров текстовым полем под нумерацией 1
2) разметить области заголовков текстовым полем под нумерацией 2
3) разметить области изображений полем для изображений под нумерацией 3 
4) распознать документ
5) сохранить как pdf с возможностью поиска

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


def get_catnum_or_name(s):
    """Извлекает партномер со страницы по первому попавшемуся
    совпадению глобального паттерна, в дальнейшем используется
    для помещения в название извлеченного со страницы изображения"""
    tokens = s.split('\n')
    catnum = ''
    for token in tokens:
        if re.search(PARTNUM_REGEX, token):
            m = re.search(PARTNUM_REGEX, token)
            catnum = m.group(0).replace(' ', '')
            if catnum.count('FIG.'):
                catnum = catnum.replace('FIG.', '')
            if re.search(r'—+|[—-]{2,}', catnum):
                catnum = re.sub(r'—+|[—-]{2,}', '-', catnum)
            if re.search(r'[^\w.,-]', catnum):
                catnum = re.sub(r'[^\w.,-]', '', catnum)
            if re.search(r'[a-zа-я]', catnum):
                catnum = catnum.upper()
            catnum = catnum.strip(' -()[]')
            break
    if catnum:
        return catnum
    else:
        name = ''
        # b=иначе по наименованию попадает не в ту папку
        # for token in tokens:
        #     if re.search(DESC_REGEX, token):
        #         m = re.search(DESC_REGEX, token)
        #         name = m.group(0)[:10].strip()
        #         break
        if name:
            return name
        else:
            return 'UNDEFINED'


def clean_folder_name(el):
    """Очищает строку от недопустимых для наименования
    папки символов"""
    if re.search(r'[\\/:*?«<>|+=;"\n]+', el):
        el = re.sub(r'[\\/:*?«<>|+=;"\n]+', '', el)
    return el


def open_pdf(vector : list):
    """Извлечение изображений из PDF, в название изображения помещаются
    номер страницы, номер изображения, партномер со страницы на которой
    было извлечено изображение
    :param vector: list, [segment number, number of CPU, filename, image_folder]
    :return: None
    """
    # this is the segment number we have to process
    idx = vector[0]
    # number of CPUs
    cpu = vector[1]
    # open file
    my_pdf_file = fitz.open(vector[2])
    # get number of pages
    num_pages = my_pdf_file.page_count

    # pages per segment: make sure that cpu * seg_size >= num_pages!
    seg_size = int(num_pages / cpu + 1)
    seg_from = idx * seg_size  # our first page number
    seg_to = min(seg_from + seg_size, num_pages)  # last page number

    for page_number in range(seg_from, seg_to):

        # acess individual page
        page = my_pdf_file[page_number]
        text = page.get_text('text', sort=True)
        header = get_catnum_or_name(text)

        # accesses all images of the page
        images = page.get_images()

        # check if images are there
        if images:
            sys.stdout.write(f'{len(images)} изображений на стр {page_number}/{num_pages}[+] CAT_NUM = {header}\n')
        else:
            sys.stdout.write(f'0 изображений на стр {page_number}[!]\n')

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
            header_str = clean_folder_name(header)
            # save image locally
            save_path = os.path.join(vector[3], f"pg{page_number_str}im{image_number_str}-{header_str}.png")
            image.save(open(save_path, "wb"), format='png')

class Extractor():
    def __init__(self, file_path, target_path):
        """
        Класс для работы с PDF, создает папки в указанной директории,
        извлекает изображения из PDF с использованием PyMuPDF,
        перемещает извлеченные изображжения в папки из директории
        каталога по совпадению партномера в названии изображения и
        названиии папки
        :param file_path: Путь к PDF файлу
        :param target_path: Путь к целевой папке для выгрузки изображений
        """

        self.target_path = os.path.realpath(target_path)
        self.file_path = os.path.realpath(file_path)
        self.filename = os.path.split(file_path)[1][:-4]
        self.image_folder = f'{self.filename}_изображения_не_перенесены'
        self.image_path = os.path.join(self.target_path, self.image_folder)
        if os.path.exists(self.image_path):
            raise Exception(f'Папка {self.image_path} уже существует') from None
        os.mkdir(self.image_path)
        self.bad_image_folder = os.path.join(self.target_path, f'{self.filename}_изображения_не_востребованы')
        if os.path.exists(self.bad_image_folder):
            raise Exception(f'Папка {self.bad_image_folder} уже существует') from None
        os.mkdir(self.bad_image_folder)
        self.commons = None
        self.content = None

    def extract(self):
        """
        Метод стартует извлечение изображений из self.file_path
        через multiprocessing
        :return: None
        """

        cpu = cpu_count()
        # make vectors of arguments for the processes
        vectors = [(i, cpu, self.file_path, self.image_path) for i in range(cpu)]
        start_time = time.time()
        # pool = Pool()  # make pool of 'cpu_count()' processes
        # pool.map(open_pdf, vectors, 1)  # start processes passing each a vector
        with Pool() as pool:
            for _ in pool.imap_unordered(open_pdf, vectors):
                pass

        sys.stdout.write("---Извлечение завершено за %s seconds ---\n" % (time.time() - start_time))
        # # open file
        # with fitz.open(self.file_path) as my_pdf_file:
        #     # loop through every page
        #     pdf_len = len(my_pdf_file)
        #     for page_number in range(1, pdf_len + 1):
        #
        #         # access individual page
        #         page = my_pdf_file[page_number - 1]
        #         text = page.get_text('text', sort=True)
        #         header = get_catnum_or_name(text)
        #
        #         # accesses all images of the page
        #         images = page.get_images()
        #
        #         # check if images are there
        #         if images:
        #             print(f"{len(images)} изображений на стр {page_number}/{pdf_len}[+]", f'CAT_NUM = {header}')
        #         else:
        #             print(f"0 изображений на стр {page_number}[!]")
        #
        #         # loop through all images present in the page
        #         for image_number, image in enumerate(images, start=1):
        #             # access image xref
        #             xref_value = image[0]
        #
        #             # extract image information
        #             base_image = my_pdf_file.extract_image(xref_value)
        #
        #             # access the image itself
        #             image_bytes = base_image["image"]
        #
        #             # get image extension
        #             # ext = base_image["ext"]
        #
        #             # load image
        #             image = Image.open(io.BytesIO(image_bytes))
        #             page_number_str = str(page_number).zfill(4)
        #             image_number_str = str(image_number).zfill(4)
        #             header_str = clean_folder_name(header)
        #             # save image locally
        #             save_path = os.path.join(self.image_path,
        #                                      f"pg{page_number_str}im{image_number_str}-{header_str}.png")
        #             image.save(open(save_path, "wb"), format='png')

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

        print(f'1) Проверьте разрешения изображений в папке {os.path.split(self.image_path)[1]}\n'
               '2) В поле "Граница группировки" на вкладке "Перемещение" введите количество изображений'
              ', выше которого будут отсеяны '
               'изображения в группах, количество изображений '
               'в которых больше введенного числа.')

    @staticmethod
    def move_imgs(border: int, image_folder, bad_image_folder, directory_folder):
        """
        Перемещение изображений по индивидуальным папкам в директории каталога
        по совпадению партномера в названии изображения и названии папки
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
            shutil.move(path, bad_image_folder)

        # перемещение нужных изображений
        content = os.listdir(image_folder)
        all_pics_count = len(content)
        moved_pics_count = 0
        # только папки:
        dirs = [i for i in os.listdir(directory_folder) if os.path.isdir(os.path.join(directory_folder, i))]
        for pic in content:
            pic_name = '-'.join(pic.split('-')[1:]).rpartition('.')[0]  # название у pic после первого "-" и до ".png"
            pic_name = pic_name.strip(' .,()[]')
            for folder in dirs:
                if re.search(pic_name, folder, flags=re.IGNORECASE):
                    shutil.move(os.path.join(image_folder, pic), os.path.join(directory_folder, folder))
                    moved_pics_count += 1
                    break
        return f'Перемещено {moved_pics_count} изображений из {all_pics_count}\n'\
               f'Папка каталога {directory_folder}\n'\
               f'{all_pics_count - moved_pics_count} файлов в папке {image_folder}.'
