# -*- coding: utf-8 -*-
import sys
import os

import pandas as pd

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QTableWidgetItem
from PyQt5.QtCore import Qt
import components.open_bli as open_bli
import components.save_bli as save_bli
from components.worker import Worker

import data.crushDecompressor as Decompressor
import data.crushCompressor as Compressor

from data.bbiredactor import Redactor as BBIRedactor
import data.catredactor as cat_redactor
import data.bliredactor as bli_redactor
import data.ldfconvertor as ldf_convertor

from components.gui.page_redactor_GUI import Ui_Page_redactor


class PageRedactor(QtWidgets.QWidget):
    def __init__(self):
        super(PageRedactor, self).__init__()
        self.ui = Ui_Page_redactor()
        self.ui.setupUi(self)

        self.read_history()
        self.binding()
        self.vars()

        self.compressor = Compressor.CrushCompressor()
        self.decompressor = Decompressor.CrushDecompressor()
        self.convertor = ldf_convertor.convertor()

    def binding(self):
        self.ui.open_action.triggered.connect(self.open_cat)
        self.ui.ldf_list.itemSelectionChanged.connect(self.read_bli)
        self.ui.save_button.clicked.connect(self.tbl2bli)
        self.ui.save_changes.triggered.connect(self.save_all_files)
        self.ui.to_excel_but.clicked.connect(self.tbl2excel)
        self.ui.to_ldf_but.clicked.connect(self.tbl2ldf)
        self.ui.from_excel_but.clicked.connect(self.excel2tbl)
        self.ui.from_ldf_but.clicked.connect(self.ldf2tbl)
        self.ui.show_pics_but.clicked.connect(self.load_pics)
        self.ui.all_to_excel_but.clicked.connect(self.alltbl2excel)

    def vars(self):
        self.original_df = []
        self.new_df = []
        self.changed = {}
        self.cat_params = {}

    def read_history(self):
        if os.path.exists('data/history'):
            self.ui.history.clear()

            with open('data/history', 'r', encoding='cp1251') as old_h:
                i = 0
                while i < 5:
                    link = old_h.readline()
                    if len(link.strip()) > 0:
                        self.ui.history.addAction(link, self.open_cat)

                    i += 1
        else:
            old_h = ''

    def load_hist(self):
        try:
            self.directory = self.sender().text().strip()
            self.get_dir(self.directory)
        except Exception as ex:
            print(ex)
            self.add_status('error: page_redactor.load_hist')

    def last_dir(self):
        if os.path.exists('data/history'):
            path = ''
            with open('data/history', 'r',
                      encoding='cp1251') as old_h:
                path = old_h.readline()
                path = path.strip()
            return path
        else:
            return ''

    def write_history(self, d):
        ## Читаем историю
        ## Если в истории каталог есть - удаляем из истории и пишем наверх
        new_h = ''
        if os.path.exists('data/history'):
            with open('data/history', 'r',
                      encoding='cp1251') as old_h:
                for line in old_h:
                    if d.strip() == line.strip():
                        new_h = ''.join([line.strip(), '\n', new_h])
                    else:
                        new_h = ''.join([new_h, line.strip(), '\n'])
            if not new_h.count(d.strip()):
                new_h = ''.join([d.strip(), '\n', new_h])
            ## Переписываем историю
            if len(new_h) > 0:
                with open('data/history', 'w',
                          encoding='cp1251') as old_h:
                    old_h.write(new_h)

                self.read_history()

    def open_cat(self):
        if self.sender() == self.ui.open_action:
            ## Выбираем каталог
            directory = QtWidgets.QFileDialog.getExistingDirectory(
                self,
                "Выберите папку",
                self.last_dir())
        elif self.sender().text().strip() == u'Сохранить все изменения':
            try:
                directory = self.cat_path
            except Exception as ex:
                print(ex)
        else:
            directory = self.sender().text().strip()

        if not os.path.exists(directory):
            return False

        self.cat_path = directory
        self.book_code = directory.replace("\\", "/").split('/')[-1]
        self.files = {}
        cat = os.path.join(directory, 'book.cat').replace('\\', '/')

        if not os.path.exists(cat):
            self.add_status('Отсутствует book.cat')
            return False

        ## Очищаем временную папку
        self.clear_tmp()
        ## Распаковываем cat
        try:
            cat, temp_size = self.decompressor.decompress_cat(cat)
        except Exception as ex:
            self.add_status('Ошибка распаковки, open_cat.decompress_cat')
            print(ex)
            return False

        ## Получаем список файлов
        self.vars()
        self.files = cat_redactor.run(cat)

        if not self.files.get('.bli'):
            self.add_status('Отсутствует book.bbi')
            return False

        ## Очищаем список файлов на GUI
        if self.files.get('.bli'):
            self.ui.ldf_list.clear()
            self.ui.ldf_list.setColumnCount(2)
            self.ui.ldf_list.setHorizontalHeaderLabels(['Страница', 'Заголовок'])
            self.ui.ldf_list.setRowCount(0)
        else:
            text = u'Статус: Нет страниц спецификаций'
            self.ui.status.setText(text)
            return False

        try:
            self.write_history(self.cat_path)
            ## Получаем структуру book.bbi
            bbi = os.path.join(directory, 'book.bbi').replace('\\', '/')
        except Exception as ex:
            text = ''.join([u'Статус: ', str(ex)])
            self.ui.status.setText(text)

        if not os.path.exists(bbi):
            text = u'Статус: Нет файла book.bbi'
            self.ui.status.setText(text)
            return False

        try:
            bbi, temp_size = self.decompressor.decompress_bbi(bbi)
        except Exception as ex:
            self.add_status('Ошибка распаковки, open_cat.decompress_bbi')
            input(ex)
            return False

        try:
            bbiredactor = BBIRedactor()
            bbiredactor.read_file(bbi)
            self.cat_struct = bbiredactor.get_struct()
            self.cat_params = bbiredactor.get_book_bbi(self.cat_struct)
        except Exception as ex:
            text = ''.join([u'Статус: ', str(ex)])
            self.ui.status.setText(text)

        ## Запускаем процесс обработки страниц
        try:
            self.addElemsThread = QtCore.QThread(parent=self)
            self.tc = open_bli.open_bli()
            self.tc.cat_path = self.cat_path
            self.tc.files = self.files['.bli']
            self.tc.bbi = self.cat_params
            self.tc.moveToThread(self.addElemsThread)
            self.tc.add_page.connect(self.add_page)
            self.tc.add_status.connect(self.add_status)
            self.addElemsThread.started.connect(self.tc.run)
            self.addElemsThread.finished.connect(
                self.addElemsThread.exit)
            self.addElemsThread.start()
        except Exception as ex:
            text = ''.join([u'Статус: ', str(ex)])
            self.ui.status.setText(text)

    def decompress_file(self, path):
        try:
            path = Decompressor.run(path)
            path = path.replace("\\", "/")

            return path
        except:
            return ''

    def compress_file(self, original_path, temp_path):
        original_path = original_path.replace("\\", "/")
        temp_path = temp_path.replace("\\", "/")

        if not os.path.exists(original_path) or not os.path.exists(temp_path):
            return False

        try:
            self.compressor.compress_bbi(temp_path, original_path)
            path = original_path.replace("\\", "/")

            return path
        except:
            return ''

    def clear_tmp(self):
        path = os.path.join('temp', self.book_code).replace('\\', '/')
        if os.path.exists(path):
            for root, dirs, files in os.walk(path):
                for file in files:
                    os.remove(os.path.join(root, file))
            os.removedirs(os.path.join('temp', self.book_code))

    @QtCore.pyqtSlot(str, str)
    def add_page(self, string, path):
        file = None
        try:
            if string.count(' | '):
                text, file = string.split(' | ')
                file = file.strip()

            if not file:
                return False

            #self.ui.ldf_list.addItem(string
            self.ui.ldf_list.horizontalHeaderItem(0).setTextAlignment(Qt.AlignLeft)
            self.ui.ldf_list.horizontalHeaderItem(1).setTextAlignment(Qt.AlignLeft)

            self.ui.ldf_list.setRowCount(self.ui.ldf_list.rowCount()+1)
            self.ui.ldf_list.setItem(self.ui.ldf_list.rowCount()-1, ## строка
                                     0, ## столбец
                                     QTableWidgetItem(str(file)))
            self.ui.ldf_list.setItem(self.ui.ldf_list.rowCount()-1, ## строка
                                     1, ## столбец
                                     QTableWidgetItem(str(text)))
            #self.ui.ldf_list.resizeColumnsToContents()

            #print(self.ui.ldf_list.width())
            self.files['.bli'][file]['path'] = str(path)
        except Exception as ex:
            print(ex, file)
            self.add_status('error: bli_redactor.add_page')

    @QtCore.pyqtSlot(str)
    def add_status(self, string):
        self.ui.status.setText(string)

    def read_bli(self):
        if self.sender() != self.ui.ldf_list:
            return False

        if self.ui.ldf_list.rowCount() == 0:
            return False

        if not self.ui.ldf_list.currentItem():
            return False

        row = self.ui.ldf_list.currentRow()
        file = self.ui.ldf_list.item(row, 0).text().strip()

        try:
            path = self.files['.bli'][file]['path']
        except Exception as ex:
            print(ex)
            return False
        try:
            cat_struct = {'add_columns': self.cat_params['ADDITIONAL_COLUMNS'],
                          'codepage': self.cat_params['BUILD_CODEPAGE']}
            ##cat_struct = self.cat_struct[]
            df, params = bli_redactor.get_data(
                path, cat_struct)

            self.files['.bli'][file]['params'] = params

            self.painting_table(df)
            self.original_df = df
        except Exception as ex:
            print(ex)
            self.add_status('error: page_redactor.read_bli, 0')
        ## Заполняем поля заголовков

        try:
            real_page = file.replace('.bli', '').upper()
            page_data = self.cat_params['PAGES'][real_page]
            ## print(page_data)
            self.ui.title_str.clear()
            self.ui.title_str.insert(str(page_data['page_name']))
            self.ui.group_str.clear()
            self.ui.group_str.insert(str(page_data['page_name']))
            self.ui.page_str.clear()
            self.ui.page_str.insert(str(page_data['page_num']))
            self.ui.ref_str.clear()
            self.ui.ref_str.insert(str(page_data['reference']))

            ## запрет на реактирование строчек
            self.ui.page_str.setReadOnly(True)
            self.ui.ref_str.setReadOnly(True)
        except Exception as ex:
            print(ex)
            self.add_status('error: page_redactor.read_bli, 1')

    def painting_table(self, df):
        ## Определяем блок таблицы
        tbl = self.ui.tableWidget
        ## Стлбцы
        tbl.setColumnCount(len(df.columns)-1)
        columns = df.columns.to_list()[1:]
        tbl.setHorizontalHeaderLabels(columns)
        ## Строки
        tbl.setRowCount(len(df))
        ## Настраиваем отображение по левому краю
        for n in [i for i in range(0, len(df.columns)-1)]:
            tbl.horizontalHeaderItem(n).setTextAlignment(Qt.AlignLeft)
        ## Заполняем ячейки
        for col_num, col in enumerate(columns):
            column = df[col].to_list()

            for str_num, elem in enumerate(column):
                if type(elem) == bytes:
                    try:
                        elem = elem.decode('utf-8')
                    except:
                        elem = elem.decode('cp1251')
                if str(elem)[-2:] == '.0':
                    elem = str(elem)[:-2]
                tbl.setItem(str_num, ## строка
                            col_num, ## столбец
                            QTableWidgetItem(str(elem)))
        ##
        tbl.resizeColumnsToContents()

    ##        self.ui.tableWidget.cellChanged.

    def tbl2df(self):
        tbl = self.ui.tableWidget
        columns = ['NN']
        for column in range(tbl.columnCount()):
            header = tbl.horizontalHeaderItem(column)
            columns.append(header.text())

        df = []
        for row in range(tbl.rowCount()):

            rowdata = [1]
            for column in range(tbl.columnCount()):
                item = tbl.item(row, column)
                if item is not None:
                    rowdata.append(str(item.text()))
                else:
                    rowdata.append('')

            df.append(rowdata)

        df = pd.DataFrame(df, columns=columns)

        return df

    def tbl2bli(self):
        try:
            new_df = self.tbl2df()
        except Exception as ex:
            print(ex)
            return False

        if (len(self.original_df) == 0 or len(new_df) == 0):
            return False

        try:
            row = self.ui.ldf_list.currentRow()
            file = self.ui.ldf_list.item(row, 0).text().strip()

            old_path = self.files['.bli'][file]['path']
            params = self.files['.bli'][file]['params']

            ## Перезаписываем параметры
            params['title'] = self.ui.title_str.text()
            params['group'] = self.ui.group_str.text()
            params['page'] = self.ui.page_str.text()
            params['ref'] = self.ui.ref_str.text()

            params['codepage'] = self.cat_params['BUILD_CODEPAGE']

            edited_ldf = self.convertor.excel_to_bli(
                old_path,
                new_df,
                params,
                False)

            ## Добавляем в список изменённых
            ##            self.changed.append(old_path)
            self.changed[params['page']] = {'lbl':params['title'],
                                            'ref':params['ref'],
                                            'path':old_path}
            ## Красим строку в списке
            item = self.ui.ldf_list.currentItem()
            brush = QtGui.QBrush
            item.setForeground(brush(QtGui.QColor("green")))

            row = self.ui.ldf_list.currentRow()
            self.ui.ldf_list.item(row, 0).setForeground(brush(QtGui.QColor("green")))
            self.ui.ldf_list.item(row, 1).setForeground(brush(QtGui.QColor("green")))

            self.add_status('Статус: изменения записаны, но не сохранены. '+
                            'Не забудьте нажать"Сохрнить все изменения" в '+
                            'разделе меню "Файл".')

        except Exception as ex:
            print(ex)
            self.add_status('error: page_redactor.tbl2bli')

    def save_all_files(self):
        if len(self.changed) == 0:
            return False

        ## Вносим изменения в book.bbi
        bbi = False
        try:
            bbi_path = os.path.join('temp', self.book_code,
                                    'book.bbi.txt')

            if os.path.exists(bbi_path):
                redactor = BBIRedactor()
                redactor.read_file(bbi_path)
                cat_struct = redactor.get_struct()
                structure = redactor.write_page_name(cat_struct, self.changed,
                                                     self.cat_params)
                temp_path = redactor.save_temp(structure, self.book_code)

                ## Запаковываем
                original_path = os.path.join(self.cat_path, 'book.bbi')
                original_path = original_path.replace('\\', '/')

                self.compressor.compress_bbi(temp_path, original_path)
                self.add_status(u'Статус: Все заголовков сохранены')
                bbi = True
            else:
                print('Отсутствует ', bbi_path)
                return False
        except Exception as ex:
            print(ex)
            self.add_status('error: page_redactor.save_all_files, 1')

        bli = False
        try:
            ## Инициируем класс для сохранения
            self.sc = save_bli.save_bli()
            self.sc.files = self.changed
            self.sc.cur_dir = self.cat_path
            self.sc.book_code = self.book_code

            self.savePageThread = QtCore.QThread()
            self.sc.moveToThread(self.savePageThread)

            self.sc.chColor.connect(self.chColor)
            self.savePageThread.finished.connect(self.savePageThread.exit)
            self.savePageThread.started.connect(self.sc.run)
            self.savePageThread.start()

            self.add_status(u'Статус: Изменения страниц сохранены')
            bli = True
        except Exception as ex:
            print(ex)
            self.add_status('error: page_redactor.save_all_files, 0')
            
        if bbi == True and bli == True:
            self.changed = {}
            self.add_status(u'Статус: Все изменения сохранены')

    def alltbl2excel(self):
        if len(self.cat_path) == 0:
            return False
        
        ## Инициируем класс для сохранения
        self.save_excel = Worker()
        self.saveExcelThread = QtCore.QThread()
        self.save_excel.moveToThread(self.saveExcelThread)
        ## Подключаем переменные
        cat_params = {'cat_path': self.cat_path,
                      'book_code': self.book_code,
                      'add_columns': self.cat_params['ADDITIONAL_COLUMNS'],
                      'codepage': self.cat_params['BUILD_CODEPAGE']}
        self.save_excel.cat_params = cat_params

        ## Подключаем слоты
        self.save_excel.running.connect(self.saveExcelThread.exit)
        self.saveExcelThread.started.connect(self.save_excel.save_excel)
        self.save_excel.newText.connect(self.add_status)
        self.saveExcelThread.start()

    def exitThread(self, thread):
        print(thread)

    def tbl2excel(self):
        
        df = ''
        try:
            df = self.tbl2df()
        except Exception as ex:
            print(ex)
            return False

        if len(df) == 0 or not isinstance(df, pd.DataFrame):
            return False

        try:
            row = self.ui.ldf_list.currentRow()
            item = self.ui.ldf_list.item(row, 0).text().strip()
            item_name = item[:item.rfind('.')]

            path = QtWidgets.QFileDialog.getSaveFileName(
                self, 'Сохранение файла xlsx',
                ''.join([str(item_name), '.xlsx']))

            if len(path[0]) > 0:
                path = path[0]
                if path.endswith('.xlsx'):
                    next
                else:
                    path = path + '.xlsx'

                try:
                    df.to_excel(path,
                                index=False)
                except Exception as ex:
                    print(ex)
                    self.add_status('error: page_redactor."df.to_excel", 1')
        except Exception as ex:
            print(ex)
            self.add_status('error: page_redactor.tbl2excel, 1')

    def tbl2ldf(self):
        try:
            df = self.tbl2df()
        except Exception as ex:
            print(ex)

        if len(df) == 0 or not isinstance(df, pd.DataFrame):
            return False

        try:
            row = self.ui.ldf_list.currentRow()
            item = self.ui.ldf_list.item(row, 0).text().strip()
            item_name = item[:item.rfind('.')]

            path = QtWidgets.QFileDialog.getSaveFileName(
                self, 'Сохранение файла xlsx',
                ''.join([str(item_name), '.ldf']))

            if len(path[0]) > 0:
                try:
                    params = {}
                    params['title'] = self.ui.title_str.text()
                    params['group'] = self.ui.group_str.text()
                    params['page'] = self.ui.page_str.text()
                    params['ref'] = self.ui.ref_str.text()
                    params['pics'] = self.files[
                        '.bli'][item[1].strip()]['params']['img']

                    ldf = self.convertor.create_ldf(df, params)
                    self.convertor.write_ldf(ldf, path[0])
                except Exception as ex:
                    print(ex)
                    self.add_status('error: page_redactor.tbl2ldf, 1.1')

        except Exception as ex:
            print(ex)
            self.add_status('error: page_redactor.tbl2ldf, 1')

    def excel2tbl(self):
        if self.ui.ldf_list.currentItem() is None:
            return False

        file = QtWidgets.QFileDialog.getOpenFileName(
            self,
            'Файл спецификации',
            '.',
            u'*.xlsx')

        if len(file[0]) > 0:
            try:
                self.new_df = self.convertor.read_file(file[0])
                self.painting_table(self.new_df)
            except Exception as ex:
                print(ex)
                self.add_status('error: page_redactor.excel2tbl, 1')

    def ldf2tbl(self):
        print('ldf2tbl1')

    def load_pics(self):
        print('load_pics')

    def scan_pos(self):
        print('scan_pos1')

    @QtCore.pyqtSlot(str)
    def chColor(self, string):
        brush = QtGui.QBrush
        
        row = self.ui.ldf_list.currentRow()
        item = self.ui.ldf_list.item(row, 0)
        item.setForeground(brush(QtGui.QColor(string)))

        item = self.ui.ldf_list.item(row, 1)
        item.setForeground(brush(QtGui.QColor(string)))

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = PageRedactor()
    window.show()  # Показываем окно
    app.exec_()
