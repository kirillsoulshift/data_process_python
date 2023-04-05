
from metadata import make_meta
from table import Table
from pdf_extractor import Extractor
from ldf import folder_surf
from reference import *
from cells import *
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, Menu
from tkinter.messagebox import showinfo, askyesno
import cells
from multiprocessing import freeze_support

# +?вопрос: ===============предложитьь вести словарь англ терминов из базы переводов LO==-=============
# вопрос: можно дополнить справочник Бизнес-глоссарий.
# +?вопрос: удобно ли раскрывающееся меню нужно ли много вложенных страниц в каталоге LO
# +?вопрос: нужно ли указывать принадлежность позиции к узлу (дефисами) в спецификации LO (или делать
#  раскрывающийся список)
# +?вопрос: пользуется ли юзер поиском? если да, то каким: по англ наименованию, по русском, по партномеру

# 1?пересмотреть правила перевода каталогов, нужен регламент, длинные позиции не влазят в заголовки страниц LO
# 2?вопрос: где регламент по переводу, по содержанию перевода, длина перевода, верхний регистр
# 3?вопрос: как переводить каталоги на русском, если в базе переводов термин исходного файла не используется
# менять ли наименование узла на общепринятое внутри полюса или оставлять так как в оригинале

# ? ?вопрос: можно автоматизировать извлечение таблиц из PDF,
# ? нужен +camelot, +opencv, -Ghostscript

# TODO? ?вкладка? ?ldf: переделывать расстановку ссылок
# +вкладка XL: удалять лидирующие нули в референс и квантити
# + вкладка pdf: wrap extractor with thread (decision: multiprocessing)
# + вкладка паттерны: аплай не аплаит? дебаг класса Patterns
# TODO поиск подлежащего, добавить в приложение

# переименовать приложение
# поменять иконку
# выровнять основное окно
# выровнять окна справки
# TODO сделать общее описание

"""
+ очистить и выгрузить эксель 
+ выгрузить слова для перевода в эксель
+ перевести и выгрузить очищенный эксель
+ создать директорию каталога
+ выгрузка и помещение изображений из PDF
+ создать директорию ldf
"""


class XLFrame(ttk.Frame):
    """Вкладка основного окна, обработки эксель файлов и
    выгрузки слов на перевод по столбцам"""
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

        # XL LabelFrame==========================================================================
        self.xlf = ttk.LabelFrame(self, text='Обработка .xlsx')
        self.xlf.grid(column=0, row=2, columnspan=3, rowspan=3, sticky=tk.NSEW, **options)
        # 0==============================================================================================
        # Clean and upload button
        self.clean_button = ttk.Button(self.xlf, text='Clean',
                                       command=self.clean_and_upload)
        self.clean_button.grid(column=2, row=0, sticky=tk.W, **options)
        self.clean_button.state(['disabled'])

        # russian characters check check box
        self.is_russian = tk.BooleanVar(self.xlf, False)
        self.is_russian_check = ttk.Checkbutton(self.xlf, text='Проверять наличие\nсимволов русской\nраскладки',
                                           variable=self.is_russian,
                                           onvalue=True, offvalue=False)
        self.is_russian_check.grid(column=0, row=0, sticky=tk.W, **options)

        # slash in header check box
        self.is_slash = tk.BooleanVar(self.xlf, False)
        self.slash_check = ttk.Checkbutton(self.xlf, text='Двуязычное написание\nзаголовков каталога',
                                           variable=self.is_slash,
                                           onvalue=True, offvalue=False)
        self.slash_check.grid(column=1, row=0, sticky=tk.W, **options)
        # 1==============================================================================================
        # identical headers without part_n delete check box
        self.del_headers = tk.BooleanVar(self.xlf, False)
        self.del_headers_check = ttk.Checkbutton(self.xlf, text='Удалять идентичные\nзаголовки следующие\n'
                                                            'подряд без партномеров',
                                           variable=self.del_headers,
                                           onvalue=True, offvalue=False)
        self.del_headers_check.grid(column=0, row=1, sticky=tk.W, **options)

        # headers merge check box
        self.merge_headers = tk.BooleanVar(self.xlf, False)
        self.merge_headers_check = ttk.Checkbutton(self.xlf, text='Выполнять слияние\nзаголовков на\n'
                                                              'соседних ячейках',
                                                 variable=self.merge_headers,
                                                 onvalue=True, offvalue=False)
        self.merge_headers_check.grid(column=1, row=1, sticky=tk.W, **options)

        # 2==============================================================================================
        # identical headers in a row check box
        self.headers_in_a_row = tk.BooleanVar(self.xlf, False)
        self.headers_in_a_row_check = ttk.Checkbutton(self.xlf, text='Проверять наличие\nидентичных заголовков\n'
                                                                 'следующих подряд',
                                                   variable=self.headers_in_a_row,
                                                   onvalue=True, offvalue=False)
        self.headers_in_a_row_check.grid(column=0, row=2, sticky=tk.W, **options)

        # ammount of characters in partnumbers check box
        self.partnumber_characters = tk.BooleanVar(self.xlf, False)
        self.partnumber_characters_check = ttk.Checkbutton(self.xlf, text='Проверять количество\n'
                                                                      'символов в партномерах',
                                                      variable=self.partnumber_characters,
                                                      onvalue=True, offvalue=False)
        self.partnumber_characters_check.grid(column=1, row=2, sticky=tk.W, **options)

        # Upload LabelFrame==========================================================================
        self.ulf = ttk.LabelFrame(self, text='Выгрузка слов на перевод из .xlsx')
        self.ulf.grid(column=0, row=5, columnspan=3, sticky=tk.NSEW, **options)
        # 0==============================================================================================
        # need to download comment col for translation check box
        self.need_comment_col = tk.BooleanVar(self.ulf, False)
        self.need_comment_col_check = ttk.Checkbutton(self.ulf, text='Добавить содержимое столбца комментариев\n'
                                                                 'к списку слов на перевод',
                                                      variable=self.need_comment_col,
                                                      onvalue=True, offvalue=False)
        self.need_comment_col_check.grid(column=0, row=0, sticky=tk.W, padx=23, pady=5)

        # Translation lists upload button
        self.transl_lists_upload_button = ttk.Button(self.ulf, text='Upload',
                                                     command=self.translation_upload)
        self.transl_lists_upload_button.grid(column=1, row=0, sticky=tk.W, padx=23, pady=5)
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
        try:
            if self.table is None:
                self.table = Table(self.source_path.get(), self.target_folder_path.get())

            self.table.process(need_rus_check=self.is_russian.get(),
                               need_slash=self.is_slash.get(),
                               del_headers=self.del_headers.get(),
                               merge_headers=self.merge_headers.get(),
                               headers_in_a_row=self.headers_in_a_row.get(),
                               partnum_char_amm=self.partnumber_characters.get())
            self.table.make_xl(is_translated=False)
            print(f'Очищенный файл загружен в {self.target_folder_path.get()}')
        except TypeError:
            print(f'В поле "Исходный файл" требуется заново указать путь к файлу\n'
                  f'{self.source_path.get()}')

    def translation_upload(self):
        if self.table is None:
            self.table = Table(self.source_path.get(), self.target_folder_path.get())
            self.table.make_translation_lists(download_comment_col=self.need_comment_col.get())
        else:
            self.table.make_translation_lists(download_comment_col=self.need_comment_col.get())
        print(f'Файлы для перевода загружены в {os.path.split(self.target_folder_path.get())[1]}')


class TranslateFrame(ttk.Frame):
    """Вкладка основного окна, перевод исходного эксель файла
    с помощью эксель словаря, содержит опцию создания списка
    соответствия терминов из спецификации с лидирующими
    символами принадлежности к подразделу и без него, список
    соответствия генерируется на вкладке XLFrame"""
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

        # translate headers check box
        self.translate_headers = tk.BooleanVar(self, False)
        self.translate_headers_check = ttk.Checkbutton(self, text='Переводить заголовки',
                                                   variable=self.translate_headers,
                                                   onvalue=True, offvalue=False)
        self.translate_headers_check.grid(column=1, row=4, sticky=tk.W, **options)
        # 5==============================================================================================
        # translate comment check box
        self.translate_comment = tk.BooleanVar(self, False)
        self.translate_comment_check = ttk.Checkbutton(self, text='Переводить комментарии',
                                                    variable=self.translate_comment,
                                                    onvalue=True, offvalue=False)
        self.translate_comment_check.grid(column=1, row=5, sticky=tk.W, **options)

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
        """два сценария с аффиксами и без;
        в случае пропущенных/дублирующихся значений в значении оригинала и перевода
        выводится диалоговое окно оповещения с возможностью удаления таких строк"""
        if self.table is None:
            self.table = Table(self.source_path.get(), self.target_folder_path.get())
        try:
            if self.afx_path.get():
                self.table.translate(dict_path=self.dict_path.get(), afx_less_path=self.afx_path.get(),
                                     translate_headers=self.translate_headers.get(),
                                     translate_comments=self.translate_comment.get())
                self.table.make_xl(is_translated=True)
            else:
                self.table.translate(dict_path=self.dict_path.get(),
                                     translate_headers=self.translate_headers.get(),
                                     translate_comments=self.translate_comment.get())
                self.table.make_xl(is_translated=True)
        except DictError as e:
            answer = askyesno(title='Оповещение',
                              message=str(e) + ' ДА: Удалить пропущенные/дублирующиеся значения и продолжить перевод'
                                               'НЕТ: Не применять никаких действий и продолжить перевод')
            # ответ 'да': перевод продолжается с удалением пропущенных/дублирующихся значений в оригинале и переводе
            if answer:
                if self.afx_path.get():
                    self.table.translate(dict_path=self.dict_path.get(), afx_less_path=self.afx_path.get(),
                                         null_ident=True,
                                         translate_headers=self.translate_headers.get(),
                                         translate_comments=self.translate_comment.get())
                    self.table.make_xl(is_translated=True)
                else:
                    self.table.translate(dict_path=self.dict_path.get(),
                                         null_ident=True,
                                         translate_headers=self.translate_headers.get(),
                                         translate_comments=self.translate_comment.get())
                    self.table.make_xl(is_translated=True)
            # ответ 'нет': перевод продолжается не применяя никаких действий к словарю
            else:
                if self.afx_path.get():
                    self.table.translate(dict_path=self.dict_path.get(), afx_less_path=self.afx_path.get(),
                                         null_ident=False,
                                         translate_headers=self.translate_headers.get(),
                                         translate_comments=self.translate_comment.get())
                    self.table.make_xl(is_translated=True)
                else:
                    self.table.translate(dict_path=self.dict_path.get(),
                                         null_ident=False,
                                         translate_headers=self.translate_headers.get(),
                                         translate_comments=self.translate_comment.get())
                    self.table.make_xl(is_translated=True)
        print(f'Переведеный файл загружен в {os.path.split(self.target_folder_path.get())[1]}')


class DirectoryFrame(ttk.Frame):
    """Вкладка основного окна, генерация директории каталога
    из исходного эксель файла"""
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
        try:
            if self.table is None:
                self.table = Table(self.source_path.get(), self.target_folder_path.get())
                self.table.make_dir()
            else:
                self.table.make_dir()
            print(f'Директория каталога сформирована в {os.path.split(self.target_folder_path.get())[1]}')
        except NoTOCError as e:
            answer = askyesno(title='Оповещение',
                              message=str(e) + ' ДА: Продолжить формирование директории каталога.'
                                               'НЕТ: Не применять никаких действий, завершить выполнение задачи.')
            if answer:
                self.table.make_dir(check_TOC=False)
                print(f'Директория каталога сформирована в {os.path.split(self.target_folder_path.get())[1]}')


class PDFFrame(ttk.Frame):
    """Вкладка основного окна, извлечение изображений из
    исходного пдф файла, алгоритм извлечения обычно использовался
    со сгенерированным файлом ABBYY FineReader <PDF с возможность поиска>
    после OCR распознавания. Генерирует две папки: с изображениями,
    и пустая папка для следующего этапа MoveImgsFrame"""
    def __init__(self, container):
        super().__init__(container)

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
        self.extract_button.state(['disabled'])
        print(f'Извлечение изображений из {os.path.split(self.pdf_path.get())[1]} ...')
        Extractor(self.pdf_path.get(), self.pdf_target_path.get()).extract()
        self.extract_button.state(['!disabled'])
        showinfo(title=f'Extraction complete',
                 message=f'Извлечение изображений из {os.path.split(self.pdf_path.get())[1]}'
                         f'завершено.')


class MoveImgsFrame(ttk.Frame):
    """Вкладка основного окна, распределеение извлеченных на
    этапе PDFFrame изображений в директории каталога, сгенерированной
    на этапе DirectoryFrame. Помещает изображения по индивидуальным
    папкам по совпадению партномера в названии изображения и
    названии папки в директории каталога"""
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
            raise ValueError(f'Требуется ввести число, получено: {self.border_value.get()}') from None
        try:
            border = int(self.border_value.get())
        except ValueError:
            raise ValueError(f'Требуется ввести число, получено: {self.border_value.get()}') from None
        message = Extractor.move_imgs(border,
                                      self.images_folder_path.get(),
                                      self.bad_images_folder_path.get(),
                                      self.cat_dir_folder_path.get())
        print(message)
        showinfo(title=f'Image movement complete', message=message)


class LDFFrame(ttk.Frame):
    """Вкладка основного окна, генерация папки с .ldf .png .inf .sdf .fmt
    на основе директории каталога, содержащей разветвленную директорию с
    файлами .xlsx .png .txt, размещает ссылки в файлах .ldf на номера вложенных
    папок и файлы вложенных в исходную папку изображений"""
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
            raise ValueError(f'Требуется заполнить поле BOOKCODE, получено: "{self.bookcode.get()}"') from None
        if not self.equipmet_type.get():
            self.equipmet_type.set('')
            # raise ValueError(f'Требуется заполнить поле Тип механизма, получено: "{self.equipmet_type.get()}"') from None
        if not self.model.get():
            self.model.set('')
            # raise ValueError(f'Требуется заполнить поле Модель, получено: "{self.model.get()}"') from None
        if not self.serial_number.get():
            self.serial_number.set('')
            # raise ValueError(f'Требуется заполнить поле Серийный номер, получено: "{self.serial_number.get()}"') from None

        curr_path = os.path.realpath(self.directory_folder_path.get())  # текущая директория
        # проверка названия директории каталога
        dir_name = os.path.split(curr_path)[1]
        if not re.search(r'^\d{4}\s\w+\s.+$', dir_name):
            raise NameError(f'Название директории каталога {dir_name} не соответствует '
                             'паттерну "0000 <Партномер> <Заголовок>"') from None

        # проверка длины путей файлов
        print(curr_path)
        path_to_review = []
        for root, dirs, files in os.walk(curr_path):
            for item in files:
                item_path = os.path.join(root, item).encode()
                path_len = len(item_path)
                if path_len >= 260:
                    path_to_review.append(item)
        if path_to_review:
            print('Пути файлов превышают длину 260 символов:')
            for item in path_to_review:
                print(f'{item}')
            raise ValueError('Слишком длинный путь') from None

        # название директории каталога (0000 NA HD1500 -> HD1500)
        res_folder_name = os.path.split(curr_path)[1].split()[-1]
        parent_path = os.path.join(curr_path, os.pardir)                 # родительская директория
        dest_path = os.path.join(parent_path, res_folder_name + '_ldf')  # путь к формируемой директории
        os.mkdir(dest_path)                                              # создание формируемой директории


        folder_surf(curr_path, dest_path, self.translate.get(), self.comment.get())
        make_meta(dest_path,
                  book_code=self.bookcode.get(),
                  eq_type=self.equipmet_type.get(),
                  model=self.model.get(),
                  serial=self.serial_number.get(),
                  translate=self.translate.get(),
                  comment=self.comment.get())
        print(f'Папка .ldf файлов создана: {dest_path}')

class SubFrame(ttk.Frame):
    """Рамка основного окна, находится внизу окна
    содержит текстовый виджет со скролбаром, виджет
    выводит сообщения и ошибки программы"""
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
    """Перенаправляет ошибки и сообщения
    программы из sys.stdout и sys.stderr"""
    def __init__(self, text_widget):
        self.text_space = text_widget

    def write(self, string):
        self.text_space.insert('end', string)
        self.text_space.see('end')

    def flush(self):
        pass


class App(tk.Tk):
    """Основное окно, содержит вкладки,
    нижнюю рамку, меню справкаб настройкиб итд"""
    def __init__(self):
        super().__init__()
        self.iconbitmap('Icon.ico')
        self.title('LinkOneCreator')
        window_width = 535
        window_height = 590

        # get the screen dimension
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # find the center point
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)

        # set the position of the window to the center of the screen
        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        self.resizable(False, False)

        self.menubar = Menu(self)
        self.config(menu=self.menubar)

        # create buttons in main menubar
        self.settings_menu = Menu(self.menubar, tearoff=0)
        self.help_menu = Menu(self.menubar, tearoff=0)

        # create the sub_menu
        self.sub_menu = Menu(self.help_menu, tearoff=0)
        self.sub_menu.add_command(label='Обработка .xlsx', command=self.clean_help)
        self.sub_menu.add_command(label='Выгрузка слов на перевод .xlsx', command=self.upload_help)

        # add menu items to the help menu
        self.help_menu.add_cascade(label='Очистка', menu=self.sub_menu)
        self.help_menu.add_command(label='Перевод', command=self.translate_help)
        self.help_menu.add_command(label='Директория', command=self.dir_help)
        self.help_menu.add_command(label='Изображение', command=self.pdf_help)
        self.help_menu.add_command(label='Перемещение', command=self.move_help)
        self.help_menu.add_command(label='Файлы LDF', command=self.ldf_help)

        self.settings_menu.add_command(label='Паттерны', command=self.pattern_settings)

        # add the File menu to the menubar
        self.menubar.add_cascade(label='Настройки', menu=self.settings_menu)
        self.menubar.add_cascade(label='Справка', menu=self.help_menu)

    def pattern_settings(self):
        window = Patterns(self)
        window.grab_set()

    def clean_help(self):
        window = AboutWindow(self, CLEAN_HELP)
        window.grab_set()

    def upload_help(self):
        window = AboutWindow(self, UPLOAD_HELP)
        window.grab_set()

    def translate_help(self):
        window = AboutWindow(self, TRANSLATE_HELP)
        window.grab_set()

    def dir_help(self):
        window = AboutWindow(self, DIRECTORY_HELP)
        window.grab_set()

    def pdf_help(self):
        window = AboutWindow(self, IMAGES_HELP)
        window.grab_set()

    def move_help(self):
        window = AboutWindow(self, MOVE_HELP)
        window.grab_set()

    def ldf_help(self):
        window = AboutWindow(self, LDF_HELP)
        window.grab_set()


class Patterns(tk.Toplevel):
    """Всплывающее окно раздела Паттерны в
    Настройках, настраивает регулярные выражения"""
    def __init__(self, parent):
        super().__init__(parent)

        self.geometry('440x700')
        self.resizable(False, False)
        self.title('Настройка')
        options = {'padx': 5, 'pady': 5}
        # Main Frame=====================================================================
        self.mainframe = ttk.Frame(self)
        self.mainframe.grid(sticky=tk.NSEW, **options)
            # Upper Frame=====================================================================
        self.frame = ttk.Frame(self.mainframe)
        self.frame.grid(column=0, row=0, sticky=tk.NSEW, **options)
                # Legend==========================================================================
        legend = 'Паттерны регулярных выражений будут использованы для определения\n' \
                 'пренадлежности ячеек таблицы к определенному типу данных. Каждый\n' \
                 'из типов данных имеет собственный алгоритм очистки'
        self.label = ttk.Label(self.frame, text=legend, anchor=tk.W)
        self.label.grid(column=0, row=0, columnspan=2, sticky=tk.NSEW, **options)
                # Partnumber==========================================================================
        self.partn_label = ttk.Label(self.frame, text='Партномер', anchor=tk.W)
        self.partn_label.grid(column=0, row=1, sticky=tk.W, **options)

        self.partn = tk.StringVar()
        self.partn.set(PARTNUM_REGEX)
        self.partn_entry = ttk.Entry(self.frame, textvariable=self.partn, width=40)
        self.partn_entry.grid(column=1, row=1, sticky=tk.W, **options)

        self.partn_error_label = ttk.Label(self.frame, text='', foreground='red', anchor=tk.W)
        self.partn_error_label.grid(column=1, row=2, sticky=tk.W)
                # Description==========================================================================
        self.desc_label = ttk.Label(self.frame, text='Описание', anchor=tk.W)
        self.desc_label.grid(column=0, row=3, sticky=tk.W, **options)

        self.desc = tk.StringVar()
        self.desc.set(DESC_REGEX)
        self.desc_entry = ttk.Entry(self.frame, textvariable=self.desc, width=40)
        self.desc_entry.grid(column=1, row=3, sticky=tk.W, **options)

        self.desc_error_label = ttk.Label(self.frame, text='', foreground='red', anchor=tk.W)
        self.desc_error_label.grid(column=1, row=4, sticky=tk.W)
                # subassy==========================================================================
        self.subassy_label = ttk.Label(self.frame, text='Символ принадлежности\n'
                                                        'к сборочной единице', anchor=tk.W)
        self.subassy_label.grid(column=0, row=5, sticky=tk.W, **options)

        self.subassy = tk.StringVar()
        self.subassy.set(SUBASSEMBLY_REGEX)
        self.subassy_entry = ttk.Entry(self.frame, textvariable=self.subassy, width=40)
        self.subassy_entry.grid(column=1, row=5, sticky=tk.W, **options)

        self.subassy_error_label = ttk.Label(self.frame, text='', foreground='red', anchor=tk.W)
        self.subassy_error_label.grid(column=1, row=6, sticky=tk.W)
                # posnum==========================================================================
        self.posnum_label = ttk.Label(self.frame, text='Номер позиции', anchor=tk.W)
        self.posnum_label.grid(column=0, row=7, sticky=tk.W, **options)

        self.posnum = tk.StringVar()
        self.posnum.set(ITEM_NUMBER_REGEX)
        self.posnum_entry = ttk.Entry(self.frame, textvariable=self.posnum, width=40)
        self.posnum_entry.grid(column=1, row=7, sticky=tk.W, **options)

        self.posnum_error_label = ttk.Label(self.frame, text='', foreground='red', anchor=tk.W)
        self.posnum_error_label.grid(column=1, row=8, sticky=tk.W)
            # Lower LabelFrame==========================================================================
        self.labelframe = ttk.LabelFrame(self.mainframe, text='Регулярные выражения')
        self.labelframe.grid(column=0, row=1, sticky=tk.NSEW, **options)
        regex_legend = "abc соответствует 'abc' (не 'ABC')\n" \
                       "[abc] соответствует 'a' или 'b' или 'c' (не 'A' или 'B' или 'C')\n" \
                       "[a-zA-Z] соответствует любой английской букве в обоих регистрах\n" \
                       "[0-9] соответствует любой цифре\n" \
                       "[^abc] соответствует любому символу кроме 'a' или 'b' или 'c'\n" \
                       "a? соответствует повторению 'a' 0 или 1 раз\n" \
                       "a+ соответствует повторению 'a' 1 или более раз\n" \
                       "a* соответствует повторению 'a' 0 или более раз\n" \
                       "a{3} соответствует 'aaa'\n" \
                       "a{1,3} соответствуе повторению 'a' от 1 до 3 раз\n" \
                       "a{3,} соответствуе повторению 'a' 3 или более раз\n" \
                       ". соответствует любому символу\n" \
                       "\. соответствует .\n" \
                       "\? соответствует ?\n" \
                       "\[ соответствует [\n" \
                       "\( соответствует (\n" \
                       "\w соответствует любой букве или цифре или символу '_'\n" \
                       "\d соответствует любой цифре\n" \
                       "\\n соответствует символу переноса строки\n" \
                       "^ обозначение начала строки\n" \
                       "$ обозночение конца строки\n" \
                       "abc|cba соответствует 'abc' или 'cba'"

        self.regex_label = ttk.Label(self.labelframe, text=regex_legend, anchor=tk.W)
        self.regex_label.grid(column=0, row=0, sticky=tk.NSEW, **options)
            # Buttons==========================================================================
        self.apply_button = ttk.Button(self.mainframe, text='Применить', command=self.apply_changes)
        self.apply_button.grid(column=0, row=2, sticky=tk.W, **options)
        # self.cancel_button = ttk.Button(self.mainframe, text='Отмена', command=self.destroy)
        # self.cancel_button.grid(column=1, row=2, sticky=tk.W, **options)

    def apply_changes(self):
        # print('BEFORE:')                                                                      # debug
        # print('cells.PARTNUM_REGEX:', cells.PARTNUM_REGEX)                                    # debug
        # print('cells.DESC_REGEX:', cells.DESC_REGEX)                                          # debug
        # print('cells.SUBASSEMBLY_REGEX:', cells.SUBASSEMBLY_REGEX)                            # debug
        # print('cells.ITEM_NUMBER_REGEX:', cells.ITEM_NUMBER_REGEX)                            # debug
        partn = self.partn.get()
        desc = self.desc.get()
        subassy = self.subassy.get()
        posnum = self.posnum.get()
        got_vars = [partn, desc, subassy, posnum]
        err_count = []
        for item in got_vars:
            try:
                if len(item) == 0:
                    raise re.error('Пустая строка в regex')
                re.compile(item)
                err_count.append(0)
            except re.error:
                err_count.append(1)
        # в случае правильности всех регулярных выражений, переменные модуля cells заменяются
        # новыми значениями
        if err_count == [0, 0, 0, 0]:
            cells.PARTNUM_REGEX = self.partn.get()
            cells.DESC_REGEX = self.desc.get()
            cells.SUBASSEMBLY_REGEX = self.subassy.get()
            cells.ITEM_NUMBER_REGEX = self.posnum.get()
            # print('AFTER:')                                                                   # debug
            # print('cells.PARTNUM_REGEX:', cells.PARTNUM_REGEX)                                # debug
            # print('cells.DESC_REGEX:', cells.DESC_REGEX)                                      # debug
            # print('cells.SUBASSEMBLY_REGEX:', cells.SUBASSEMBLY_REGEX)                        # debug
            # print('cells.ITEM_NUMBER_REGEX:', cells.ITEM_NUMBER_REGEX)                        # debug
            self.destroy()
        # в одном или более выражениях ошибка, нужный error_label отображает сообщение об ошибке
        # красным шрифтом, выражение с ошибкой (entry) подсвечивается красным шрифтом
        else:
            entries = [self.partn_entry,
                       self.desc_entry,
                       self.subassy_entry,
                       self.posnum_entry]
            error_labels = [self.partn_error_label,
                            self.desc_error_label,
                            self.subassy_error_label,
                            self.posnum_error_label]
            for entry, err in zip(entries, err_count):
                if err == 1:
                    entry['foreground'] = 'red'
            for error_label, err in zip(error_labels, err_count):
                if err == 1:
                    error_label['text'] = 'Ошибка в регулярном выражении'


class AboutWindow(tk.Toplevel):
    """Всплывающее окно справки с рамкой и скроллбаром,
    рамка отображает текст справки
    """
    def __init__(self, parent, about):
        super().__init__(parent)
        #               w x h
        self.geometry('700x770')
        # self.resizable(False, False)
        self.title('Справка')

        self.frame = ttk.Frame(self)
        self.frame.grid(padx=10, pady=10, sticky=tk.NSEW)

        self.canvas = tk.Canvas(self.frame, width=640, height=700, bg='white')
        self.canvas.grid(column=0, row=0, padx=5, pady=5, sticky=tk.E)

        about = re.sub(r'\t', '    ', about)
        self.canvas.create_text((10, 10), text=about)
        # self.text = tk.Text(self.frame, wrap='word', width=100, height=80)
        # about = GENERAL_HELP
        # self.text.insert('1.0', about)
        # self.text.grid(column=0, row=0, padx=5, pady=5, sticky=tk.W)

        self.scrollbar = ttk.Scrollbar(self.frame, orient='vertical', command=self.canvas.yview)
        self.scrollbar.grid(column=1, row=0, sticky=tk.NS, padx=5, pady=5)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        # self.canvas['yscrollcommand'] = self.scrollbar.set
        self.canvas.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox('all')))

        self.close = ttk.Button(self.frame, text='OK', command=self.destroy)\
            .grid(column=0, row=1, padx=5, pady=5, ipadx=5, ipady=5, sticky=tk.W)


if __name__ == "__main__":
    freeze_support()
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
    notebook.add(frame4, text='Изображения')
    notebook.add(frame5, text='Перемещение')
    notebook.add(frame6, text='Файлы LDF')

    SubFrame(app).grid(column=0, row=1)
    app.mainloop()