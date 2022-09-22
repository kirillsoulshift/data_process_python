import re
from metadata import make_meta
from table import *
from pdf_extractor import *
from ldf import *
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog
from tkinter.messagebox import showinfo
"""
+ очистить и выгрузить эксель 
+ выгрузить слова для перевода в эксель
+ перевести и выгрузить очищенный эксель
+ создать директорию каталога
+ выгрузка и помещение изображений из PDF
+ создать директорию ldf
"""


class XLFrame(ttk.Frame):
    def __init__(self, container):
        super().__init__(container)
        self.table = None
        # field options
        options = {'padx': 5, 'pady': 5}
        # 0==============================================================================================
        # source path label
        self.source_path_label = ttk.Label(self, text='Исходный файл:', anchor=tk.W)
        self.source_path_label.grid(column=0, row=0, sticky=tk.W, **options)

        # source path entry
        self.source_path = tk.StringVar()
        self.source_path_entry = ttk.Entry(self, textvariable=self.source_path, width=50)
        self.source_path_entry.grid(column=1, row=0, **options)

        # source path browse button
        self.source_path_browse_button = ttk.Button(self, text='Browse',
                                                    command=self.browse_source_file)
        self.source_path_browse_button.grid(column=2, row=0, sticky=tk.W, **options)
        # 1==============================================================================================
        # target folder path label
        self.target_folder_path_label = ttk.Label(self, text='Целевая папка:', anchor=tk.W)
        self.target_folder_path_label.grid(column=0, row=1, sticky=tk.W, **options)

        # target folder path entry
        self.target_folder_path = tk.StringVar()
        self.target_folder_path_entry = ttk.Entry(self, textvariable=self.target_folder_path, width=50)
        self.target_folder_path_entry.grid(column=1, row=1, **options)

        # target folder path browse button
        self.target_folder_path_browse_button = ttk.Button(self, text='Browse',
                                                           command=self.browse_target_folder)
        self.target_folder_path_browse_button.grid(column=2, row=1, sticky=tk.W, **options)
        # 2==============================================================================================
        # Clean and upload button
        self.clean_button = ttk.Button(self, text='Clean',
                                       command=self.clean_and_upload)
        self.clean_button.grid(column=2, row=2, sticky=tk.W, **options)
        self.clean_button.state(['disabled'])
        # 3==============================================================================================
        # Translation lists upload button
        self.transl_lists_upload_button = ttk.Button(self, text='Trados',
                                                     command=self.translation_upload)
        self.transl_lists_upload_button.grid(column=2, row=3, sticky=tk.W, **options)
        self.transl_lists_upload_button.state(['disabled'])
        self.grid(padx=10, pady=10, sticky=tk.NSEW)

    def update(self):
        if self.source_path.get() and self.target_folder_path.get():
            self.table = Table(self.source_path.get(), self.target_folder_path.get())

    def browse_source_file(self):
        self.source_path.set(filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')]))
        self.update()
        if self.source_path.get() and self.target_folder_path.get():
            self.clean_button.state(['!disabled'])
            self.transl_lists_upload_button.state(['!disabled'])

    def browse_target_folder(self):
        self.target_folder_path.set(filedialog.askdirectory())
        self.update()
        if self.source_path.get() and self.target_folder_path.get():
            self.clean_button.state(['!disabled'])
            self.transl_lists_upload_button.state(['!disabled'])

    def clean_and_upload(self):
        if self.table is None:
            self.table = Table(self.source_path.get(), self.target_folder_path.get())
        self.table.process()
        self.table.make_xl(is_translated=False)

    def translation_upload(self):
        if self.table is None:
            self.table = Table(self.source_path.get(), self.target_folder_path.get())
            self.table.make_translation_lists()
        else:
            self.table.make_translation_lists()
        print(f'Файлы для перевода загружены в {os.path.split(self.target_folder_path.get())[1]}')


class TranslateFrame(ttk.Frame):
    def __init__(self, container):
        super().__init__(container)
        self.table = None
        # field options
        options = {'padx': 5, 'pady': 5}
        # 0==============================================================================================
        # source path label
        self.source_path_label = ttk.Label(self, text='Исходный файл:', anchor=tk.W)
        self.source_path_label.grid(column=0, row=0, sticky=tk.W, **options)

        # source path entry
        self.source_path = tk.StringVar()
        self.source_path_entry = ttk.Entry(self, textvariable=self.source_path, width=50)
        self.source_path_entry.grid(column=1, row=0, **options)

        # source path browse button
        self.source_path_browse_button = ttk.Button(self, text='Browse',
                                                    command=self.browse_source_file)
        self.source_path_browse_button.grid(column=2, row=0, sticky=tk.W, **options)
        # 1==============================================================================================
        # target folder path label
        self.target_folder_path_label = ttk.Label(self, text='Целевая папка:', anchor=tk.W)
        self.target_folder_path_label.grid(column=0, row=1, sticky=tk.W, **options)

        # target folder path entry
        self.target_folder_path = tk.StringVar()
        self.target_folder_path_entry = ttk.Entry(self, textvariable=self.target_folder_path, width=50)
        self.target_folder_path_entry.grid(column=1, row=1, **options)

        # target folder path browse button
        self.target_folder_path_browse_button = ttk.Button(self, text='Browse',
                                                           command=self.browse_target_folder)
        self.target_folder_path_browse_button.grid(column=2, row=1, sticky=tk.W, **options)
        # 2==============================================================================================
        # Dictionary path label
        self.dict_path_label = ttk.Label(self, text='Файл словаря:', anchor=tk.W)
        self.dict_path_label.grid(column=0, row=2, sticky=tk.W, **options)

        # Dictionary path entry
        self.dict_path = tk.StringVar()
        self.dict_path_entry = ttk.Entry(self, textvariable=self.dict_path, width=50)
        self.dict_path_entry.grid(column=1, row=2, **options)

        # Dictionary path browse button
        self.transl_lists_browse_button = ttk.Button(self, text='Browse',
                                                     command=self.browse_dict_file)
        self.transl_lists_browse_button.grid(column=2, row=2, sticky=tk.W, **options)
        # 3==============================================================================================
        # Affixless path label
        self.afx_path_label = ttk.Label(self, text='Файл\nсопоставления:', anchor=tk.W)
        self.afx_path_label.grid(column=0, row=3, sticky=tk.W, **options)

        # Affixless path entry
        self.afx_path = tk.StringVar()
        self.afx_path_entry = ttk.Entry(self, textvariable=self.afx_path, width=50)
        self.afx_path_entry.grid(column=1, row=3, **options)

        # Affixless path browse button
        self.afx_browse_button = ttk.Button(self, text='Browse',
                                            command=self.browse_afx_file)
        self.afx_browse_button.grid(column=2, row=3, sticky=tk.W, **options)
        # 4==============================================================================================
        # Translate and upload button
        self.translate_button = ttk.Button(self, text='Translate',
                                           command=self.translate)
        self.translate_button.grid(column=2, row=4, sticky=tk.W, **options)
        self.translate_button.state(['disabled'])
        self.grid(padx=10, pady=10, sticky=tk.NSEW)

    def update(self):
        if self.source_path.get() and self.target_folder_path.get():
            self.table = Table(self.source_path.get(), self.target_folder_path.get())

    def browse_source_file(self):
        self.source_path.set(filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')]))
        self.update()
        if self.source_path.get() and self.target_folder_path.get():
            self.translate_button.state(['!disabled'])

    def browse_target_folder(self):
        self.target_folder_path.set(filedialog.askdirectory())
        self.update()
        if self.source_path.get() and self.target_folder_path.get():
            self.translate_button.state(['!disabled'])

    def browse_dict_file(self):
        self.dict_path.set(filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')]))
        if self.source_path.get() and self.target_folder_path.get():
            self.translate_button.state(['!disabled'])

    def browse_afx_file(self):
        self.afx_path.set(filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')]))

    def translate(self):
        """два сценария с аффиксами и без"""
        if self.afx_path.get():
            if self.table is None:
                self.table = Table(self.source_path.get(), self.target_folder_path.get())
                self.table.translate(self.dict_path.get(), self.afx_path.get())
                self.table.make_xl(is_translated=True)
            else:
                self.table.translate(self.dict_path.get(), self.afx_path.get())
                self.table.make_xl(is_translated=True)
        else:
            if self.table is None:
                self.table = Table(self.source_path.get(), self.target_folder_path.get())
                self.table.translate(self.dict_path.get())
                self.table.make_xl(is_translated=True)
            else:
                self.table.translate(self.dict_path.get())
                self.table.make_xl(is_translated=True)
        print(f'Переведеный файл загружен в {os.path.split(self.target_folder_path.get())[1]}')


class DirectoryFrame(ttk.Frame):
    def __init__(self, container):
        super().__init__(container)
        self.table = None
        options = {'padx': 5, 'pady': 5}
        # 0==============================================================================================
        # source path label
        self.source_path_label = ttk.Label(self, text='Исходный файл:', anchor=tk.W)
        self.source_path_label.grid(column=0, row=0, sticky=tk.W, **options)

        # source path entry
        self.source_path = tk.StringVar()
        self.source_path_entry = ttk.Entry(self, textvariable=self.source_path, width=50)
        self.source_path_entry.grid(column=1, row=0, **options)

        # source path browse button
        self.source_path_browse_button = ttk.Button(self, text='Browse',
                                                    command=self.browse_source_file)
        self.source_path_browse_button.grid(column=2, row=0, sticky=tk.W, **options)
        # 1==============================================================================================
        # target folder path label
        self.target_folder_path_label = ttk.Label(self, text='Целевая папка:', anchor=tk.W)
        self.target_folder_path_label.grid(column=0, row=1, sticky=tk.W, **options)

        # target folder path entry
        self.target_folder_path = tk.StringVar()
        self.target_folder_path_entry = ttk.Entry(self, textvariable=self.target_folder_path, width=50)
        self.target_folder_path_entry.grid(column=1, row=1, **options)

        # target folder path browse button
        self.target_folder_path_browse_button = ttk.Button(self, text='Browse',
                                                           command=self.browse_target_folder)
        self.target_folder_path_browse_button.grid(column=2, row=1, sticky=tk.W, **options)

        # 2==============================================================================================
        # Make catalog directory with excels button
        self.make_directory_button = ttk.Button(self, text='Make Dir',
                                                command=self.make_directory)
        self.make_directory_button.grid(column=2, row=2, sticky=tk.W, **options)
        self.make_directory_button.state(['disabled'])
        self.grid(padx=10, pady=10, sticky=tk.NSEW)

    def update(self):
        if self.source_path.get() and self.target_folder_path.get():
            self.table = Table(self.source_path.get(), self.target_folder_path.get())

    def browse_source_file(self):
        self.source_path.set(filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')]))
        self.update()
        if self.source_path.get() and self.target_folder_path.get():
            self.make_directory_button.state(['!disabled'])

    def browse_target_folder(self):
        self.target_folder_path.set(filedialog.askdirectory())
        self.update()
        if self.source_path.get() and self.target_folder_path.get():
            self.make_directory_button.state(['!disabled'])

    def make_directory(self):
        if self.table is None:
            self.table = Table(self.source_path.get(), self.target_folder_path.get())
            self.table.make_dir()
        else:
            self.table.make_dir()
        print(f'Директория каталога сформирована в {os.path.split(self.target_folder_path.get())[1]}')


class PDFFrame(ttk.Frame):
    def __init__(self, container):
        super().__init__(container)
        self.extr = None
        options = {'padx': 5, 'pady': 5}

        # 0==============================================================================================
        # PDF path label
        self.pdf_path_label = ttk.Label(self, text='Исходный\nPDF:', anchor=tk.E)
        self.pdf_path_label.grid(column=0, row=0, sticky=tk.W, **options)

        # target folder path entry
        self.pdf_path = tk.StringVar()
        self.pdf_path_entry = ttk.Entry(self, textvariable=self.pdf_path, width=50)
        self.pdf_path_entry.grid(column=1, row=0, **options)

        # target folder path browse button
        self.pdf_path_browse_button = ttk.Button(self, text='Browse',
                                                 command=self.browse_pdf)
        self.pdf_path_browse_button.grid(column=2, row=0, sticky=tk.W, **options)
        # 1==============================================================================================
        # PDF target path label
        self.pdf_target_path_label = ttk.Label(self, text='Целевая\nпапка:', anchor=tk.W)
        self.pdf_target_path_label.grid(column=0, row=1, sticky=tk.W, **options)

        # PDF target path entry
        self.pdf_target_path = tk.StringVar()
        self.pdf_target_path_entry = ttk.Entry(self, textvariable=self.pdf_target_path, width=50)
        self.pdf_target_path_entry.grid(column=1, row=1, **options)

        # PDF target path browse button
        self.pdf_target_path_browse_button = ttk.Button(self, text='Browse',
                                                 command=self.browse_pdf_target_path)
        self.pdf_target_path_browse_button.grid(column=2, row=1, sticky=tk.W, **options)
        # 2==============================================================================================
        # PDF Extract button
        self.extract_button = ttk.Button(self, text='Extract',
                                           command=self.pdf_extract)
        self.extract_button.grid(column=2, row=2, sticky=tk.W, **options)
        self.extract_button.state(['disabled'])
        self.grid(padx=10, pady=10, sticky=tk.NSEW)

    def browse_pdf(self):
        self.pdf_path.set(filedialog.askopenfilename(filetypes=[('PDF Files', '*.pdf')]))
        if self.pdf_path.get() and self.pdf_target_path.get():
            self.extract_button.state(['!disabled'])

    def browse_pdf_target_path(self):
        self.pdf_target_path.set(filedialog.askdirectory())
        if self.pdf_path.get() and self.pdf_target_path.get():
            self.extract_button.state(['!disabled'])

    def pdf_extract(self):
        self.extr = Extractor(file_path=self.pdf_path.get(),
                              target_path=self.pdf_target_path.get())
        message = self.extr.extract()
        print(message)
        showinfo(title=f'Extraction complete',
                 message=f'Извлечение изображений из {os.path.split(self.pdf_path.get())[1]}'
                         f'завершено. ' + message)
        self.extr = None


class MoveImgsFrame(ttk.Frame):
    def __init__(self, container):
        super().__init__(container)
        options = {'padx': 5, 'pady': 5}
        # 0==============================================================================================
        # images folder
        self.images_folder_path_label = ttk.Label(self, text='Не\nперемещенные\nизображения:', anchor=tk.E)
        self.images_folder_path_label.grid(column=0, row=0, sticky=tk.W, **options)

        # images folder path entry
        self.images_folder_path = tk.StringVar()
        self.images_folder_path_entry = ttk.Entry(self, textvariable=self.images_folder_path, width=50)
        self.images_folder_path_entry.grid(column=1, row=0, **options)

        # images folder path browse button
        self.images_folder_browse_button = ttk.Button(self, text='Browse',
                                                 command=self.browse_images_folder)
        self.images_folder_browse_button.grid(column=2, row=0, sticky=tk.W, **options)
        # 1==============================================================================================
        # bad images folder
        self.bad_images_folder_path_label = ttk.Label(self, text='Не\nвостребованные\nизображения:', anchor=tk.E)
        self.bad_images_folder_path_label.grid(column=0, row=1, sticky=tk.W, **options)

        # images folder path entry
        self.bad_images_folder_path = tk.StringVar()
        self.bad_images_folder_path_entry = ttk.Entry(self, textvariable=self.bad_images_folder_path, width=50)
        self.bad_images_folder_path_entry.grid(column=1, row=1, **options)

        # images folder path browse button
        self.bad_images_folder_browse_button = ttk.Button(self, text='Browse',
                                                      command=self.browse_bad_images_folder)
        self.bad_images_folder_browse_button.grid(column=2, row=1, sticky=tk.W, **options)
        # 2==============================================================================================
        # catalog directory folder
        self.cat_dir_folder_path_label = ttk.Label(self, text='Директория\nкаталога:', anchor=tk.E)
        self.cat_dir_folder_path_label.grid(column=0, row=2, sticky=tk.W, **options)

        # images folder path entry
        self.cat_dir_folder_path = tk.StringVar()
        self.cat_dir_folder_path_entry = ttk.Entry(self, textvariable=self.cat_dir_folder_path, width=50)
        self.cat_dir_folder_path_entry.grid(column=1, row=2, **options)

        # images folder path browse button
        self.cat_dir_folder_browse_button = ttk.Button(self, text='Browse',
                                                          command=self.browse_cat_dir_folder)
        self.cat_dir_folder_browse_button.grid(column=2, row=2, sticky=tk.W, **options)
        # 3==============================================================================================
        # border label
        self.border_label = ttk.Label(self, text='Граница\nгруппировки:', anchor=tk.W)
        self.border_label.grid(column=0, row=3, sticky=tk.W, **options)

        # border entry
        self.border_value = tk.StringVar()
        self.border_value_entry = ttk.Entry(self, textvariable=self.border_value, width=10)
        self.border_value_entry.grid(column=1, row=3, sticky=tk.W, **options)

        # Move imgs button
        self.move_button = ttk.Button(self, text='Move', command=self.move)
        self.move_button.grid(column=2, row=3, sticky=tk.W, **options)
        self.move_button.state(['disabled'])
        self.grid(padx=10, pady=10, sticky=tk.NSEW)

    def browse_images_folder(self):
        self.images_folder_path.set(filedialog.askdirectory())
        if self.images_folder_path.get() and \
                self.bad_images_folder_path.get() and \
                self.cat_dir_folder_path.get():
            self.move_button.state(['!disabled'])

    def browse_bad_images_folder(self):
        self.bad_images_folder_path.set(filedialog.askdirectory())
        if self.images_folder_path.get() and \
                self.bad_images_folder_path.get() and \
                self.cat_dir_folder_path.get():
            self.move_button.state(['!disabled'])

    def browse_cat_dir_folder(self):
        self.cat_dir_folder_path.set(filedialog.askdirectory())
        if self.images_folder_path.get() and \
                self.bad_images_folder_path.get() and \
                self.cat_dir_folder_path.get():
            self.move_button.state(['!disabled'])

    def move(self):
        if not self.border_value.get():
            raise ValueError(f'Требуется ввести число, получено: {self.border_value.get()}')
        try:
            border = int(self.border_value.get())
        except ValueError:
            raise ValueError(f'Требуется ввести число, получено: {self.border_value.get()}')
        message = Extractor.move_imgs(border,
                                      self.images_folder_path.get(),
                                      self.bad_images_folder_path.get(),
                                      self.cat_dir_folder_path.get())
        print(message)
        showinfo(title=f'Image movement complete', message=message)


class LDFFrame(ttk.Frame):
    def __init__(self, container):
        super().__init__(container)
        options = {'padx': 5, 'pady': 5}
        # 0==============================================================================================
        # directory folder
        self.directory_folder_path_label = ttk.Label(self, text='Директория\nкаталога:', anchor=tk.E)
        self.directory_folder_path_label.grid(column=0, row=0, sticky=tk.W, **options)

        # directory folder path entry
        self.directory_folder_path = tk.StringVar()
        self.directory_folder_path_entry = ttk.Entry(self, textvariable=self.directory_folder_path, width=50)
        self.directory_folder_path_entry.grid(column=1, row=0, sticky=tk.W, **options)

        # directory folder path browse button
        self.directory_folder_browse_button = ttk.Button(self, text='Browse',
                                                 command=self.browse_directory_folder)
        self.directory_folder_browse_button.grid(column=2, row=0, sticky=tk.W, **options)
        # 1==============================================================================================
        # BOOKCODE
        self.bookcode_label = ttk.Label(self, text='BOOKCODE:', anchor=tk.E)
        self.bookcode_label.grid(column=0, row=1, sticky=tk.W, **options)

        # BOOKCODE entry
        self.bookcode = tk.StringVar()
        self.bookcode_entry = ttk.Entry(self, textvariable=self.bookcode, width=20)
        self.bookcode_entry.grid(column=1, row=1, sticky=tk.W, **options)

        # translate check box
        self.translate = tk.BooleanVar(self, False)
        self.translate_check = ttk.Checkbutton(self, text='Столбец\nперевода', variable=self.translate,
                                               onvalue=True, offvalue=False)
        self.translate_check.grid(column=2, row=1, sticky=tk.W, **options)
        # 2==============================================================================================
        # equipmet type
        self.equipmet_type_label = ttk.Label(self, text='Тип\nмеханизма:', anchor=tk.E)
        self.equipmet_type_label.grid(column=0, row=2, sticky=tk.W, **options)

        # equipmet type entry
        self.equipmet_type = tk.StringVar()
        self.equipmet_type_entry = ttk.Entry(self, textvariable=self.equipmet_type, width=20)
        self.equipmet_type_entry.grid(column=1, row=2, sticky=tk.W, **options)

        # comment check box
        self.comment = tk.BooleanVar(self, False)
        self.comment_check = ttk.Checkbutton(self, text='Столбец\nкомментария', variable=self.comment,
                                               onvalue=True, offvalue=False)
        self.comment_check.grid(column=2, row=2, sticky=tk.W, **options)
        # 3==============================================================================================
        # model
        self.model_label = ttk.Label(self, text='Модель:', anchor=tk.E)
        self.model_label.grid(column=0, row=3, sticky=tk.W, **options)

        # model
        self.model = tk.StringVar()
        self.model_entry = ttk.Entry(self, textvariable=self.model, width=20)
        self.model_entry.grid(column=1, row=3, sticky=tk.W, **options)
        # 4==============================================================================================
        # serial number
        self.serial_number_label = ttk.Label(self, text='Серийный\nномер:', anchor=tk.E)
        self.serial_number_label.grid(column=0, row=4, sticky=tk.W, **options)

        # serial number entry
        self.serial_number = tk.StringVar()
        self.serial_number_entry = ttk.Entry(self, textvariable=self.serial_number, width=20)
        self.serial_number_entry.grid(column=1, row=4, sticky=tk.W, **options)

        # Make ldf button
        self.make_ldf_button = ttk.Button(self, text='Make .ldf', command=self.make_ldf)
        self.make_ldf_button.grid(column=2, row=4, sticky=tk.W, **options)
        self.make_ldf_button.state(['disabled'])
        self.grid(padx=10, pady=10, sticky=tk.NSEW)

    def browse_directory_folder(self):
        self.directory_folder_path.set(filedialog.askdirectory())
        self.make_ldf_button.state(['!disabled'])

    def make_ldf(self):
        if not self.bookcode.get():
            raise ValueError(f'Требуется заполнить поле BOOKCODE, получено: "{self.bookcode.get()}"')
        if not self.equipmet_type.get():
            raise ValueError(f'Требуется заполнить поле Тип механизма, получено: "{self.equipmet_type.get()}"')
        if not self.model.get():
            raise ValueError(f'Требуется заполнить поле Модель, получено: "{self.model.get()}"')
        if not self.serial_number.get():
            raise ValueError(f'Требуется заполнить поле Серийный номер, получено: "{self.serial_number.get()}"')

        curr_path = os.path.realpath(self.directory_folder_path.get())  # текущая директория
        # проверка названия директории каталога
        dir_name = os.path.split(curr_path)[1]
        if not re.search(r'^\d{4}\s(\w+|NA)\s[\s\w-]+$', dir_name):
            raise NameError(f'Название директории каталога {dir_name} не соответствует '
                             'паттерну "0000 <Партномер> <Заголовок>"')

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
            print('Пути файлов превышают длину 260 символов:')
            for item in path_to_review:
                print(f'{item}')
            raise ValueError('Слишком длинный путь')

        # название директории каталога (0000 NA HD1500 -> HD1500)
        res_folder_name = os.path.split(curr_path)[1].split()[-1]
        parent_path = os.path.join(curr_path, os.pardir)                 # родительская директория
        dest_path = os.path.join(parent_path, res_folder_name + '_ldf')  # путь к формируемой директории
        os.mkdir(dest_path)                                              # создание формируемой директории
        folder_surf(curr_path, dest_path)
        make_meta(dest_path,
                  book_code=self.bookcode.get(),
                  eq_type=self.equipmet_type.get(),
                  model=self.model.get(),
                  serial=self.serial_number.get(),
                  translate=self.translate.get(),
                  comment=self.comment.get())
        print(f'Папка .ldf файлов создана: {dest_path}')

class SubFrame(ttk.Frame):
    def __init__(self, container):
        super().__init__(container)
        # field options
        options = {'padx': 5, 'pady': 5}
        # 0==============================================================================================
        # Text widget
        self.text_widget = tk.Text(self, height=10, width=60, wrap='word')
        self.text_widget.grid(column=0, row=0, sticky=tk.W, **options)
        self.scrollbar = ttk.Scrollbar(self, orient='vertical', command=self.text_widget.yview)
        self.scrollbar.grid(column=1, row=0, sticky=tk.NS, **options)
        self.text_widget['yscrollcommand'] = self.scrollbar.set
        sys.stdout = StdoutRedirector(self.text_widget)
        sys.stderr = StdoutRedirector(self.text_widget)

        self.grid(padx=10, pady=10, sticky=tk.NSEW)



class StdoutRedirector(object):
    def __init__(self, text_widget):
        self.text_space = text_widget

    def write(self, string):
        self.text_space.insert('end', string)
        self.text_space.see('end')

    def flush(self):
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title('LinkOne FileReview')
        window_width = 530
        window_height = 450

        # get the screen dimension
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # find the center point
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)

        # set the position of the window to the center of the screen
        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        self.resizable(False, False)


if __name__ == "__main__":
    app = App()
    notebook = ttk.Notebook(app)
    notebook.grid(column=0, row=0)

    # create frames
    frame1 = XLFrame(notebook)
    frame2 = TranslateFrame(notebook)
    frame3 = DirectoryFrame(notebook)
    frame4 = PDFFrame(notebook)
    frame5 = MoveImgsFrame(notebook)
    frame6 = LDFFrame(notebook)

    # add frames to notebook
    notebook.add(frame1, text='Очистка')
    notebook.add(frame2, text='Перевод')
    notebook.add(frame3, text='Директория')
    notebook.add(frame4, text='PDF')
    notebook.add(frame5, text='Перемещение')
    notebook.add(frame6, text='ldf')

    SubFrame(app).grid(column=0, row=1)
    app.mainloop()