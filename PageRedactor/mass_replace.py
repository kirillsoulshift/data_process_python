import os
import re
import sys
import traceback
# import data.corrector as corrector

# import data.ldfconvertor as ldfconvertor
# import data.pixgenerator as pixgenerator
import data.indexer as indexer
import pandas as pd
from PyQt5 import QtGui, QtCore
from PyQt5.QtCore import QCoreApplication, QThread, pyqtSignal, Qt
from PyQt5.QtWidgets import QMainWindow, QDesktopWidget, QWidget, QLabel, QLineEdit, \
    QTextEdit, QGridLayout, QApplication, QMessageBox, QListWidgetItem, \
    QTableWidgetItem, QInputDialog, QGraphicsPixmapItem, QPushButton, QFileDialog
# import data.tasks as tasks
from components.worker import Worker
from data.bbiredactor import Redactor as bbiRedactor
import data.bliredactor as bliredactor
from data.crushDecompressor import CrushDecompressor
import data.crushCompressor as Compressor

import data.ldfconvertor as ldfconvertor

# from components.gui.gui import Ui_MainWindow
# import components.redact_params as redact_params
# import components.img_viewer as img_viewer
# from components.page_redactor import PageRedactor as PageRedactor
# from components.replace_form import ReplaceForm
# from data.crushCompressor import compress as crushCompressor
# from data.translator import translator as translator
# from data.crushTranslate import Translator


"""self.cat_path  = bbi_path

+ указать сколько каталогов
- указать сколько элементов заменено в каталоге
+ сделать выбор папки откуда берутся каталоги
+ сделать выбор папки откуда берутся словари (словари склеиваются)
+ массовая замена
- индексация

"""


def merge_dicts(path):
    print(f'gathering dicts from {path} ...')
    filenames = os.listdir(path)
    files = []
    for file in filenames:
        if file.endswith('.xlsx') and file.count('~') == 0:
            files.append(file)
    if len(files) == 0:
        raise Exception(f'Папка {path} не содержит словарей')
    for n, file in enumerate(files):
        if n == 0:
            df = pd.read_excel(os.path.join(path, file), header=None, dtype=str)
        else:
            df = pd.concat([df, pd.read_excel(os.path.join(path, file), header=None, dtype=str)])

    df.columns = ['ItemName', 'Translation']
    df.drop_duplicates(subset='ItemName', inplace=True, ignore_index=True)
    check_df_1 = df[df['ItemName'] == df['Translation']]
    check_df_2 = df[df['Translation'].isna() == True]
    check_df_3 = df[df['ItemName'].isna() == True]
    check_df_4 = df[df['Translation'] == '']
    check_df_5 = df[df['ItemName'] == '']
    check_list = list(set(check_df_1.index.tolist() +
                          check_df_2.index.tolist() +
                          check_df_3.index.tolist() +
                          check_df_4.index.tolist() +
                          check_df_5.index.tolist()))
    df.drop(check_list, inplace=True)
    df.reset_index(inplace=True, drop=True)

    print(f'gathered {len(files)} dicts')
    return df

def get_bli_path(cat_path) -> list:
    files = []
    for file in os.listdir(cat_path):
        if file.endswith('.bli') and file.count('~') == 0:
            file = os.path.join(cat_path, file).replace('\\', '/')
            files.append(file)
    if len(files) == 0:
        raise Exception(f'Путь {cat_path} не содержит .bli')
    return files

def get_bbi_path(cat_path) -> object:
    print(f'Collecting bbi_path in {cat_path} ...')
    files = []
    for file in os.listdir(cat_path):
        if file.endswith('.bbi') and file.count('~') == 0:
            file = os.path.join(cat_path, file).replace('\\', '/')
            files.append(file)
    if len(files) != 1:
        raise Exception(f'Путь {cat_path} не содержит .bbi / .bbi > 1')
    return files[0]


def get_LinkOne_dirs(path: str) -> (list, str):
    """Рекурсивный поиск директорий с контентом книги LinkOne
    в пути path,
    директория книги = папка: содиржит ровно один .cat, ровно один .bbi, хотя бы один .bli, не находящиеся
    в папке содержащей not_process в названии
    директория производителя = все остальные папки в path кроме содержащей not_process в названии

    :param path: папка с которой начинается поиск
    :return: список абсолютных путей книг LinkOne
    """

    print(f'Collecting LinkOne directories in {path} ...')

    LO_dirs = []
    Manu_dirs = []

    def find_LOs(directory: str, collector: list, manu_collector: list):
        content = os.listdir(directory)
        further_prc_content = []

        dir_content = [i for i in content if not i.count('not_process')
                       and os.path.isdir(os.path.join(directory, i))]

        for folder in dir_content:
            folder_content = os.listdir(os.path.join(directory, folder))
            suffixes = [i.rpartition('.')[-1] for i in folder_content if i.count('.')]
            if ('bli' in suffixes or 'BLI' in suffixes) and \
                    (suffixes.count('bbi') == 1 or suffixes.count('BBI') == 1) and \
                    (suffixes.count('cat') == 1 or suffixes.count('CAT') == 1):
                if os.path.join(directory, folder) not in collector:
                    collector.append(os.path.join(directory, folder))
            else:
                further_prc_content.append(os.path.join(directory, folder))
                manu_collector.append(folder)

        for folder in further_prc_content:
            find_LOs(folder, collector, manu_collector)

    find_LOs(path, LO_dirs, Manu_dirs)
    mes = f'Manufacturer directories: {", ".join(Manu_dirs)}\n' \
          f'LinkOne directories total: {len(LO_dirs)}'
    print(mes)
    return LO_dirs, mes

class ThreadWorker(QtCore.QObject):
    finished = QtCore.pyqtSignal()
    progress = QtCore.pyqtSignal(int)
    decompressor = CrushDecompressor()
    compressor = Compressor.CrushCompressor()
    bbiredactor = bbiRedactor()
    convertor = ldfconvertor.convertor()
    current_dir = None
    directories_set = []
    dictionary = None
    report = []
    finished_catalogs_count = 0
    bbiparams = {}


    def index_cat(self):

        cat_params = {'cat_path': self.current_dir,
                      'add_columns': self.bbiparams['ADDITIONAL_COLUMNS'],
                      'codepage': self.bbiparams['BUILD_CODEPAGE']}
        indexator = indexer.Indexator(cat_params)
        indexator.run()

    def launch_per_dir(self):
        catalogs_count = len(self.directories_set)
        for dir_ in self.directories_set:
            self.current_dir = dir_

            cat_name = os.path.split(self.current_dir)[1]
            cat_manufacturer = os.path.split(os.path.split(self.current_dir)[0])[1]

            try:
                print(f'replacing in {cat_name} ...')

                bbi_path = get_bbi_path(self.current_dir)

                # получение параметров книги и контента bbi
                temp_bbi_path, temp_bbi_size = self.decompressor.decompress_bbi(bbi_path)
                self.bbiredactor.read_file(temp_bbi_path)
                bbistructure = self.bbiredactor.get_struct()
                self.bbiparams = self.bbiredactor.get_book_bbi(bbistructure)

                cat_params = {}
                cat_params['add_columns'] = self.bbiparams['ADDITIONAL_COLUMNS']
                cat_params['cat_path'] = self.current_dir
                cat_params['codepage'] = self.bbiparams['BUILD_CODEPAGE']
                cat_params['columns'] = ['PARTNUMBER']

                result = {'elems_cnt': 0,
                          'changed': 0}
                elems_cnt = 0
                changed = 0

                files = os.listdir(self.current_dir)
                files = [i for i in files if i.endswith('.bli') or i.endswith('.BLI')]
                for file in files:
                    print(f'entered {file}')

                    original_path = os.path.join(self.current_dir, file)
                    original_path = original_path.replace('\\', '/')

                    temp_data = self.decompressor.decompress_bli(original_path)
                    temp_path, temp_size = temp_data[0], temp_data[1]
                    if temp_path is None:
                        continue
                    temp_path = temp_path.replace('\\', '/')
                    df, params = self.bliredactor.get_data(temp_path, cat_params)
                    print(f'got table {file}')

                    original_df = df.copy()
                    for col in cat_params['columns']:
                        df = self.bliredactor.translate_column(  # переведенный датафрейм
                            df, col, dictionary, cat_params)

                    original = original_df['PARTNUMBER'].to_list()
                    augmented = df['PARTNUMBER'].to_list()
                    test_df = pd.DataFrame({'original': original, 'augmented': augmented})
                    changed += len(test_df[test_df['original'] != test_df['augmented']])
                    elems_cnt += len(original_df['PARTNUMBER'])
                    print(f'{file} translated')
                    ## Конвертируем
                    params['codepage'] = cat_params['codepage']
                    edited_ldf = self.convertor.excel_to_bli(
                        temp_path, df,
                        params, False)

                    ## Запаковываем
                    self.compressor.compress_bli(temp_path, original_path)
                    print(f'{file} compressed')

                result['elems_cnt'] = elems_cnt
                result['changed'] = changed
                print(result)
                print(f'Indexing {cat_name} ...')
                self.index_cat()
                print(f'Indexing complete')
                self.report.append([cat_manufacturer, cat_name, 'COMPLETED', result['elems_cnt'], result['changed']])
                self.finished_catalogs_count += 1
                with open('REFERENCE_PARSER_LOG.txt', 'r+') as f:
                    f.seek(0, 2)
                    f.write(cat_manufacturer + ' ' + cat_name + ' COMPLETED ' + 'elems_cnt ' + str(elems_cnt) + ' changed '
                            + str(changed) + '\n')
                self.progress.emit(self.finished_catalogs_count)
            except Exception:
                print(repr(logging.exception('error msg')))
                self.report.append([cat_manufacturer, cat_name, 'FAILED', logging.exception('error msg'), ''])
                with open('REFERENCE_PARSER_LOG.txt', 'r+') as f:
                    f.seek(0, 2)
                    f.write(cat_manufacturer + ' ' + cat_name + ' FAILED ' + repr(logging.exception('error msg')) + '\n')
                self.progress.emit(self.finished_catalogs_count)
                continue

        print(f'FINISHED: {catalogs_count} catalogs, {self.finished_catalogs_count} complete')
        self.report = pd.DataFrame(self.report, columns=['man', 'cat', 'stat', 'elements', 'changed_el'])
        self.report.to_excel('mass_replace_report.xlsx', index=False)
        self.finished.emit()



class MassReplacer(QWidget):
    """'LinkOne Mass Replace'"""

    def __init__(self):
        super().__init__()
        self.thread = QThread()
        self.params = {}
        self.global_dict = None
        self.finished_catalogs_count = 0
        self.linkone_dirs_path = ''
        self.dicts_path = ''
        self.report = []
        self.lo_cat_dirs = []
        self.directories_message = ''
        self.current_dir = ''

        self.initUI()

    def initUI(self):

        self.search_folder = QLabel('Папка каталогов\nLinkOne')
        self.search_folder_edit = QLineEdit()
        self.browse_search_folder = QPushButton('Browse', self)
        self.browse_search_folder.clicked.connect(self.linkone_directory_browser)

        self.dicts_folder = QLabel('Папка словарей\nзамены партномеров')
        self.dicts_folder_edit = QLineEdit()
        self.browse_dicts_folder = QPushButton('Browse', self)
        self.browse_dicts_folder.clicked.connect(self.linkone_dicts_browser)

        self.launch_btn = QPushButton('Launch', self)
        self.launch_btn.clicked.connect(self.launch_non_parallel)

        grid = QGridLayout()
        grid.setSpacing(5)

        grid.addWidget(self.search_folder, 1, 0)
        grid.addWidget(self.search_folder_edit, 1, 1)
        grid.addWidget(self.browse_search_folder, 1, 2)

        grid.addWidget(self.dicts_folder, 2, 0)
        grid.addWidget(self.dicts_folder_edit, 2, 1)
        grid.addWidget(self.browse_dicts_folder, 2, 2)

        grid.addWidget(self.launch_btn, 3, 2)

        self.setLayout(grid)

        self.setGeometry(300, 300, 850, 200)
        self.setWindowTitle('LinkOne Mass Replace')
        self.show()

    def linkone_directory_browser(self):
        self.linkone_dirs_path = QFileDialog.getExistingDirectory(self,
                                                                  'Выберите папку с каталогами LinkOne для замены'
                                                                  'партномеров:',
                                                                  'C:\\',
                                                                  QFileDialog.ShowDirsOnly)

        self.search_folder_edit.setText(self.linkone_dirs_path)

    def linkone_dicts_browser(self):
        self.dicts_path = QFileDialog.getExistingDirectory(self,
                                                           'Выберите папку со словарями для замены'
                                                           'партномеров:',
                                                           'C:\\',
                                                           QFileDialog.ShowDirsOnly)

        self.dicts_folder_edit.setText(self.dicts_path)

    def replace_data(self, path):

        cat_params = {}
        cat_params['add_columns'] = self.params['ADDITIONAL_COLUMNS']
        cat_params['cat_path'] = path
        cat_params['codepage'] = self.params['BUILD_CODEPAGE']
        cat_params['columns'] = []

        # cols = ['NN', 'ITEMID', 'PARTNUMBER', 'DESCRIPTION',
        #         'ALTERNATIVE', 'LINKPUB', 'LINKBOOK',
        #         'LINKREF', 'LINKITEM',
        #         'WEIGHT', 'UNITS', 'QUANTITY']
        cat_params['columns'].append('PARTNUMBER')

        ## Объявляем поток
        replace_data_w = Worker()
        replace_data_w.moveToThread(self.thread)

        ## Подключаем переменные
        replace_data_w.cat_params = cat_params
        replace_data_w.dic = self.global_dict

        ## Подключаем слоты
        # replace_data_w.newText.connect(self.addNewText)
        # replace_data_w.mess.connect(self.mess)
        replace_data_w.running.connect(self.finishThread)
        # replace_data_w.get_dict.connect(save_data)
        self.thread.started.connect(replace_data_w.replace_data)

        ## Запускаем
        self.thread.start()

    def index_cat(self, path):

        cat_params = {'cat_path': path,
                      'add_columns': self.params['ADDITIONAL_COLUMNS'],
                      'codepage': self.params['BUILD_CODEPAGE']}

        indexator = indexer.Indexator(cat_params)
        indexator.run()


    def parallel_launch(self):
        decompressor = CrushDecompressor()
        bbiredactor = bbiRedactor()

        self.global_dict = merge_dicts(self.dicts_path)

        report = []

        LO_cat_dirs, directories_message = get_LinkOne_dirs(self.linkone_dirs_path)
        catalogs_count = len(LO_cat_dirs)
        finished_catalogs_count = 0
        for dir in LO_cat_dirs:
            cat_manufacturer = os.path.split(os.path.split(dir)[0])[1]
            cat_name = os.path.split(dir)[1]
            # try:

            bbi_path = get_bbi_path(dir)
            print(f'Getting parameters of {os.path.split(dir)[1]} ...')
            # получение параметров книги и контента bbi
            temp_bbi_path, temp_bbi_size = decompressor.decompress_bbi(bbi_path)
            bbiredactor.read_file(temp_bbi_path)
            bbistructure = bbiredactor.get_struct()
            self.params = bbiredactor.get_book_bbi(bbistructure)
            print(f'Replacing {os.path.split(dir)[1]} ...')
            self.replace_data(dir)
            print(f'Indexing {os.path.split(dir)[1]} ...')
            self.index_cat(dir)
            print(f'Indexing complete')
            report.append([cat_manufacturer, cat_name, 'COMPLETED', ''])
            finished_catalogs_count += 1

            # except:
            #     report.append([cat_manufacturer, cat_name, 'FAILED', traceback.format_exc().splitlines()[-1]])

        print(directories_message)
        print(f'{catalogs_count} catalogs, {finished_catalogs_count} complete')
        if catalogs_count != finished_catalogs_count:
            report = pd.DataFrame(report, columns=['MANUFACTURER', 'CAT_NAME', 'STATUS', 'MESSAGE'])
            report.to_excel('mass_replace_report.xlsx', index=False)

    def launch_in_thread(self):
        print('launch in thread started...')
        try:
            f = open('REFERENCE_PARSER_LOG.txt', 'x')
            f.write('\n')
            f.close()
        except FileExistsError:
            print('log file found: REFERENCE_PARSER_LOG.txt')

        self.global_dict = merge_dicts(self.dicts_path)
        self.lo_cat_dirs, self.directories_message = get_LinkOne_dirs(self.linkone_dirs_path)
        print(self.directories_message)
        catalogs_count = len(self.lo_cat_dirs)
        print('catalogs_count:', catalogs_count)

        worker = ThreadWorker()
        worker.directories_set = self.lo_cat_dirs
        worker.dictionary = self.global_dict
        worker.moveToThread(self.thread)
        self.thread.started.connect(worker.launch_per_dir)
        worker.finished.connect(self.thread.quit)
        worker.finished.connect(worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        worker.progress.connect(self.reportProgress)
        self.thread.start()



    def launch_non_parallel(self):
        print('Non-parallel launch started...')
        try:
            f = open('REFERENCE_PARSER_LOG.txt', 'x')
            f.write('\n')
            f.close()
        except FileExistsError:
            print('log file found')

        decompressor = CrushDecompressor()
        compressor = Compressor.CrushCompressor()
        bbiredactor = bbiRedactor()

        convertor = ldfconvertor.convertor()

        self.global_dict = merge_dicts(self.dicts_path)
        report = []
        LO_cat_dirs, directories_message = get_LinkOne_dirs(self.linkone_dirs_path)
        print(directories_message)
        catalogs_count = len(LO_cat_dirs)
        print('catalogs_count:', catalogs_count)
        finished_catalogs_count = 0
        for dir in LO_cat_dirs:


            cat_name = os.path.split(dir)[1]
            cat_manufacturer = os.path.split(os.path.split(dir)[0])[1]
            try:
                print(f'replacing in {cat_name} ...')
                # bli_paths = get_bli_path(dir)
                bbi_path = get_bbi_path(dir)

                # получение параметров книги и контента bbi
                temp_bbi_path, temp_bbi_size = decompressor.decompress_bbi(bbi_path)
                bbiredactor.read_file(temp_bbi_path)
                bbistructure = bbiredactor.get_struct()
                self.params = bbiredactor.get_book_bbi(bbistructure)

                cat_params = {}
                cat_params['add_columns'] = self.params['ADDITIONAL_COLUMNS']
                cat_params['cat_path'] = dir
                cat_params['codepage'] = self.params['BUILD_CODEPAGE']
                cat_params['columns'] = ['PARTNUMBER']

                result = {'elems_cnt': 0,
                          'changed': 0}
                elems_cnt = 0
                changed = 0

                files = os.listdir(dir)
                files = [i for i in files if i.endswith('.bli') or i.endswith('.BLI')]
                for file in files:
                    print(f'entered {file}')

                    original_path = os.path.join(dir, file)
                    original_path = original_path.replace('\\', '/')

                    temp_data = decompressor.decompress_bli(original_path)
                    temp_path, temp_size = temp_data[0], temp_data[1]
                    if temp_path is None:
                        continue
                    temp_path = temp_path.replace('\\', '/')
                    df, params = bliredactor.get_data(temp_path, cat_params)
                    print(f'got table {file}')

                    original_df = df.copy()
                    for col in cat_params['columns']:
                        df = bliredactor.translate_column(  # переведенный датафрейм
                            df, col, self.global_dict, cat_params)
                    if 'PARTNUMBER' not in original_df.columns.to_list():
                        continue
                    original = original_df['PARTNUMBER'].to_list()
                    augmented = df['PARTNUMBER'].to_list()
                    test_df = pd.DataFrame({'original': original, 'augmented': augmented})
                    changed += len(test_df[test_df['original'] != test_df['augmented']])
                    elems_cnt += len(original_df['PARTNUMBER'])
                    print(f'{file} translated')
                    ## Конвертируем
                    params['codepage'] = cat_params['codepage']
                    edited_ldf = convertor.excel_to_bli(
                        temp_path, df,
                        params, False)

                    ## Запаковываем
                    compressor.compress_bli(temp_path, original_path)
                    print(f'{file} compressed')

                result['elems_cnt'] = elems_cnt
                result['changed'] = changed
                print(result)
                print(f'Indexing {os.path.split(dir)[1]} ...')
                self.index_cat(dir)
                print(f'Indexing complete')
                report.append([cat_manufacturer, cat_name, 'COMPLETED', result['elems_cnt'], result['changed']])
                finished_catalogs_count += 1
                with open('REFERENCE_PARSER_LOG.txt', 'r+') as f:
                    f.seek(0, 2)
                    f.write(cat_manufacturer+' '+cat_name+' COMPLETED '+'elems_cnt '+str(elems_cnt)+' changed '
                            +str(changed)+'\n')


            except Exception as e:
                print(traceback.format_exc())
                report.append([cat_manufacturer, cat_name, 'FAILED', traceback.format_exc(), ''])
                with open('REFERENCE_PARSER_LOG.txt', 'r+') as f:
                    f.seek(0, 2)
                    f.write('\n'+cat_manufacturer+' '+cat_name+' FAILED '+str(e)+'\n\n')
                continue


        print(directories_message)
        print(f'{catalogs_count} catalogs, {finished_catalogs_count} complete')
        if catalogs_count != finished_catalogs_count:
            report = pd.DataFrame(report, columns=['man', 'cat', 'stat', 'elements', 'changed_el'])
            report.to_excel('mass_replace_report.xlsx', index=False)




if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MassReplacer()
    app.exec_()
