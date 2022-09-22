from cells import *
import pandas as pd
import re
import os

"""

"""


class Table():
    def __init__(self, path, target_folder):
        """
        при инициации:
        - проверка наличия заголовков
        - определение столбца описания
        :param path: путь к исходному эксель файлу
        :param target_folder: путь к целевой папке
        """
        self.target = os.path.realpath(target_folder)
        path = os.path.realpath(path)
        print('Загрузка таблицы...')
        df = pd.read_excel(path, header=0, dtype=str, nrows=1)
        test_check = ''.join(df.columns.tolist())                                       # проверка наличия заголовков
        if not re.search(r'(?=.*part[\s_]?num)(?=.*qty)(?=.*desc)', test_check, flags=re.IGNORECASE):
            raise ColumnNameError(path)
        df = pd.read_excel(path, header=0, dtype=str)
        self.path = path
        self.destfilename = os.path.split(self.path)[1][:-5]
        self.df = df
        self.columns = df.columns.tolist()
        self.untranslated = pd.DataFrame({'empty': []})
        for col in self.columns:                                                        # определение столбца описания
            if re.search(r'desc', col, flags=re.IGNORECASE):
                self.desc_col = col
        self.find_bool_column()
        print('Таблица загружена.')

    def find_bool_column(self):
        for col in self.columns:
            if re.search(r'is_header', col, flags=re.IGNORECASE):
                self.df[col] = [find_bool(i) for i in self.df[col]]
                break

    def clean_df(self):
        """
        - Удалить пустые строки
        - Очищение ячеек таблицы
        - Создание колонки 'is_header'
        - Удалить строку в датафрейме совпадающую с наименованиями колонок
        """
        self.df.dropna(how='all', inplace=True)
        self.df.reset_index(drop=True, inplace=True)
        for col in self.df.columns:                                                     # очищение таблицы
            self.df[col] = [Cell(val, col).wright() for val in self.df[col]]
        self.df['is_header'] = [check_row(row) for row in self.df.values]               # принадлежность к заголовкам
        self.columns = self.df.columns.tolist()
        nums_to_del = []
        for num_del, value in enumerate(self.df[self.desc_col]):
            if value == value:
                if value.upper() == self.desc_col.upper():
                    nums_to_del.append(num_del)
        if len(nums_to_del) != 0:
            self.df.drop(nums_to_del, axis=0, inplace=True)
            self.df.reset_index(drop=True, inplace=True)

    def align_headers(self):
        """поиск заголовков ошибочно разделенных на несколько
        смежных ячеек по горизонтали;  соединение в одну ячейку"""
        fractured_headers = []
        for n in range(len(self.df)):
            if self.df['is_header'][n] and \
                    re.search(DESC_REGEX, self.df.iloc[n, 0]) and \
                    not qck_search(DESC_REGEX, self.df.iloc[n-1, 0]) and \
                    qck_search(DESC_REGEX, self.df.iloc[n+1, 0]):            # фрагмент заголовка является началом
                fractured_headers.append([self.df.iloc[n, 0], n, 'start'])

            elif self.df['is_header'][n] and \
                    re.search(DESC_REGEX, self.df.iloc[n, 0]) and \
                    qck_search(DESC_REGEX, self.df.iloc[n-1, 0]) and \
                    qck_search(DESC_REGEX, self.df.iloc[n+1, 0]):            # фрагмент заголовка является продолжением
                fractured_headers.append([self.df.iloc[n, 0], n, 'continuation'])

            elif self.df['is_header'][n] and \
                    re.search(DESC_REGEX, self.df.iloc[n, 0]) and \
                    qck_search(DESC_REGEX, self.df.iloc[n-1, 0]) and \
                    not qck_search(DESC_REGEX, self.df.iloc[n+1, 0]):        # фрагмент заголовка является окончанием
                fractured_headers.append([self.df.iloc[n, 0], n, 'continuation'])

        rows_for_del = [item[1] for item in fractured_headers if item[2] == 'continuation']
        for num, item in enumerate(fractured_headers):
            if item[2] == 'start':
                statement = item[0]
                count = num
                while len(fractured_headers) > count + 1 and fractured_headers[count + 1][2] == 'continuation':
                    statement += ' ' + fractured_headers[num + 1][0]
                    count += 1

                self.df.iloc[item[1], 0] = statement
        self.df.drop(index=rows_for_del, inplace=True)
        self.df.reset_index(drop=True, inplace=True)

    def del_mul_headers(self):
        """удалить заголовки, идентичные
        заголовку выше, проверка по
        партномерам и заголовкам производится отдельно"""
        headers = []
        for n, item in enumerate(self.df[self.columns[0]]):
            if self.df['is_header'][n]:
                headers.append(item)
        ser = pd.Series(headers, copy=False)
        mul_headers = ser.value_counts()[lambda x: x > 1].index.tolist()
        headers_to_del = []
        partnums_to_del = []
        for n, item in enumerate(self.df[self.columns[0]]):
            if item in mul_headers and \
                    re.search(DESC_REGEX, item) and \
                    item not in headers_to_del[-1:]:                             # заголовок отличается от предыдущего
                headers_to_del.append(item)                                      # добавление заголовка в чек-лист
            elif item in mul_headers and \
                    re.search(PARTNUM_REGEX, item) and \
                    item not in partnums_to_del[-1:]:                            # партномер отличается от предыдущего
                partnums_to_del.append(item)                                     # добавление партномера в чек-лист
            elif item in mul_headers and \
                    re.search(DESC_REGEX, item) and \
                    item in headers_to_del[-1:]:                                 # заголовок равен предыдущему
                self.df.drop(n, inplace=True)                                    # удаление текущей строки из датафрейма
            elif item in mul_headers and \
                    re.search(PARTNUM_REGEX, item) and \
                    item in partnums_to_del[-1:]:                                # партномер равен предыдущему
                self.df.drop(n, inplace=True)                                    # удаление текущей строки из датафрейма
        self.df.reset_index(drop=True, inplace=True)

    def fix_headers(self):
        """Проверить правильность заполнения
        партномеров и заголовков
        проверить порядок следования 'партномер -> заголовок'"""

        row_len = len(self.columns)
        empty_row = [float('nan') for i in range(row_len)]
        headers = []
        processed = []

        for n in range(len(self.df)):
            value = self.df.iloc[n, 0]
            new_list = empty_row[:]
            new_list.pop()
            new_list.insert(0, value)  # в первую ячейку пустой строки вставляется значение заголовка
            new_list.pop()
            new_list.append(True)  # в последнюю ячейку пустой строки вставляется значение True (это заголовок)
            if n != 0:
                if self.df['is_header'][n]:
                    if re.search(DESC_REGEX, value) \
                            and qck_search(PARTNUM_REGEX, self.df.iloc[n-1, 0]):              # заголовок с партномером
                        processed.append(new_list)
                        headers.append(new_list)

                    elif re.search(PARTNUM_REGEX, value) \
                            and qck_search(DESC_REGEX, self.df.iloc[n+1, 0]):                 # партномер с заголовком
                        processed.append(new_list)

                    elif re.search(DESC_REGEX, value) \
                            and not qck_search(PARTNUM_REGEX, self.df.iloc[n-1, 0]):          # заголовок без партномера
                        processed.append(empty_row)
                        processed.append(new_list)
                        headers.append(new_list)

                    elif re.search(PARTNUM_REGEX, value) \
                            and not qck_search(DESC_REGEX, self.df.iloc[n+1, 0]):             # партномер без заголовка
                        processed.append(new_list)
                        processed.append(headers[-1])

                else:
                    processed.append([self.df.iloc[n, i] for i in range(row_len)])
            else:
                if self.df['is_header'][n]:
                    if re.search(PARTNUM_REGEX, value):                 # партномер
                        processed.append(new_list)

                    elif re.search(DESC_REGEX, value):                  # заголовок
                        processed.append(empty_row)
                        processed.append(new_list)
                        headers.append(new_list)
                else:
                    if self.df.iloc[n, :].isna().all():                 # если все ячейки пустые
                        processed.append(empty_row)
                    else:
                        raise Exception('Первая строка таблицы не является заголовком')

        new_df = pd.DataFrame(processed)
        new_df.columns = self.columns
        self.df = new_df

    def dub_in_a_row(self):
        """Проверка на наличие нескольких одинаковых заголовков подряд"""

        headers = []
        headers_to_rev = []
        for n, value in enumerate(self.df[self.columns[0]]):
            if self.df['is_header'][n] and qck_search(DESC_REGEX, value):
                if value not in headers[-1:]:
                    headers.append(value)
                else:
                    headers_to_rev.append([n, value])
        if headers_to_rev:
            review = pd.DataFrame(headers_to_rev)
            review.columns = ['Num.', 'Header']
            print('Заголовки следующие подряд:')
            print(review)

    def is_russian_chars(self):
        """проверка наличия русских символов в столбцах"""
        for col in self.columns:
            if not self.df[col].apply(is_english).all():
                to_print = self.df[col].apply(is_english).loc[lambda x: x == False]
                print(f'=====Символы русской раскладки в столбце: <{col}> =====')
                print(self.df[col].iloc[to_print.index.tolist()].drop_duplicates())

    def partnum_len(self):
        """проверка длин всех партномеров в символах и вывод результата"""
        part_num_col = ''
        for col in self.columns:
            if re.search(r'part[\s_]?num', col, flags=re.IGNORECASE):
                part_num_col = col
        if part_num_col:
            res = self.df[part_num_col].dropna().apply(lambda x: len(x)).value_counts()
            df_values = pd.DataFrame({'Количество знаков партномера': res.index, 'Количество парномеров': res.values})
            print(df_values)
        else:
            print('Столбец Part Number не обнаружен')

    def process(self):
        self.is_russian_chars()                     # проверка на наличие символов русской раскладки
        self.clean_df()                             # удалить строки заголовков таблиц
        self.align_headers()                        # соединить заголовок разделенный на несколько соседних ячеек
        self.del_mul_headers()                      # удалить повторяющиеся заголовки и каталожные номера заголовков
        self.fix_headers()                          # проверка правильности заполнения заголовков и каталожных номеров
        self.dub_in_a_row()                         # проверка на наличие нескольких одинаковых заголовков подряд
        self.partnum_len()                          # проверка длин партномеров

    def make_xl(self, is_translated: bool):
        if is_translated:
            self.df.to_excel(os.path.join(self.target, self.destfilename + '_translated.xlsx'), index=False)
            if not self.untranslated.empty:
                self.untranslated.to_excel(os.path.join(self.target, self.destfilename + '_untranslated.xlsx'),
                                           header=False, index=False)
        else:
            self.df.to_excel(os.path.join(self.target, self.destfilename + '_cleaned.xlsx'), index=False)

    def make_translation_lists(self):
        headers = self.df[self.columns[0]]          # копирование столбца заголовков
        headers.dropna(inplace=True)
        headers.drop_duplicates(inplace=True)
        headers.reset_index(drop=True, inplace=True)
        idx_to_remove = []
        for n, item in enumerate(headers):
            if re.search(ITEM_NUMBER_REGEX, item):  # удаление партномеров и номеров позиций
                idx_to_remove.append(n)
        headers.drop(idx_to_remove, inplace=True)
        headers.to_excel(os.path.join(self.target, self.destfilename + '_headers.xlsx'), header=False, index=False)

        rows = self.df[self.desc_col]              # копирование столбца описания
        rows.dropna(inplace=True)
        rows.drop_duplicates(inplace=True)
        subassy = []
        for item in rows:                           # проверка наличия символов принадлежности к подузлу
            if re.search(SUBASSEMBLY_REGEX, item):
                subassy.append(item)
                break
        if subassy:
            rows_afx_less = pd.DataFrame(rows)
            rows_afx_less['2'] = rows.apply(lambda x: re.sub(SUBASSEMBLY_REGEX, '', x))
            rows_afx_less.to_excel(os.path.join(self.target, self.destfilename + '_rows_afx_less.xlsx'),
                                   header=False, index=False)
        rows.to_excel(os.path.join(self.target, self.destfilename + '_rows.xlsx'), header=False, index=False)

    def translate(self, dict_path, afx_less_path=None):

        # получение словаря
        dictionary = pd.read_excel(dict_path, header=None, dtype=str)
        dictionary.columns = ['source', 'target']

        # проверка наличия пустых ячеек
        if dictionary.isnull().values.any():
            raise DictError(dict_path)

        # проверка наличия идентичных значений в столбцах (source=target)
        check_df_1 = dictionary[dictionary['source'] == dictionary['target']]
        if len(check_df_1) != 0:
            raise DictError(dict_path)

        # проверка наличия дубликатов в словаре
        check_df_1 = dictionary[dictionary.duplicated(subset=['source'])]
        if len(check_df_1) != 0:
            dictionary.drop_duplicates(subset=['source'], inplace=True)
            print('Дубликаты словаря удалены')

        # получение списка соответствий
        if afx_less_path:
            afx_less = pd.read_excel(os.path.realpath(afx_less_path), header=None, dtype=str)
            afx_less.columns = ['source', 'afxless']
            afx_less['target'] = dictionary.iloc[:, 1]
            discrepancy = []
            for i in afx_less.index:
                if afx_less.iloc[i, 0] != afx_less.iloc[i, 1]:
                    discrepancy.append(i)
            for i in discrepancy:
                afx_less.iloc[i, 2] = what_prefix(afx_less, i, 0) + afx_less.iloc[i, 2]

            afx_less.drop(['afxless'], axis=1, inplace=True)
            dictionary = afx_less

        merged = self.df.merge(dictionary, how='left', left_on=self.df[self.desc_col], right_on='source')
        merged.drop(columns='source', inplace=True)

        # проверка наличия непереведенных позиций в description
        not_translated = []
        for item, translated in zip(merged[self.desc_col], merged['target']):
            if not isNaN(item) and \
                    isNaN(translated) and \
                    item not in not_translated and \
                    item.lower() != self.desc_col.lower():
                not_translated.append(item)
            elif item == translated:
                not_translated.append(item)

        # перевод столбца заголовков, первого столбца
        translated_headers = []
        untranslated_headers = []
        for item in merged[merged.columns[0]]:
            if dictionary['source'].isin([item]).any():
                translated_headers.append(dictionary[dictionary['source'] == item]['target'].values[0])
            else:
                if item == item and \
                        re.search(DESC_REGEX, item, flags=re.IGNORECASE) and \
                        item.lower() != self.desc_col.lower():
                    untranslated_headers.append(item)
                translated_headers.append(item)

        merged[merged.columns[0]] = translated_headers

        # перевод столбца 'Comment'
        comment_column = ''
        for col in merged.columns:  # определение столбца коментариев
            if re.search(r'comment', col, flags=re.IGNORECASE):
                comment_column = col
        if comment_column:
            translated_comments = []
            for item in merged[comment_column]:
                if dictionary['source'].isin([item]).any():
                    translated_comments.append(dictionary[dictionary['source'] == item]['target'].values[0])
                else:
                    translated_comments.append(item)

            merged[comment_column] = translated_comments

        # вывод непереведенных заголовков (если есть)
        if len(untranslated_headers) != 0 or len(not_translated) != 0:
            untranslated_headers.extend(not_translated)
            untranslated_headers = list(set(untranslated_headers))
            untranslated = pd.DataFrame({'untrnsl': untranslated_headers})
            untranslated.dropna(inplace=True)
            self.untranslated = untranslated
        self.df = merged
        self.columns = self.df.columns.tolist()

    def make_dir(self):
        """
        - проверка порядка столбцов
        - перестановка столбцов в соответствии с
            0            1            2          3          4              5            6
        'Item ID', 'Part Number', 'Quantity', 'Unit', 'Description', 'Translation', 'Comment'
        - проверка некодируемых символов
        - проверка начала таблицы
        - проверка наличия наименований разделов в таблице
        - создание директории каталога
        """

        # проверка наличия дополнительных столбцов
        need_rearrangement = False
        test_check = ''.join(self.columns)
        if not re.search(r'unit', test_check, flags=re.IGNORECASE):
            self.df['UNIT'] = [float('nan') for i in range(len(self.df))]
            self.columns = self.df.columns.tolist()
            need_rearrangement = True
        if not re.search(r'comment', test_check, flags=re.IGNORECASE):
            self.df['COMMENT'] = [float('nan') for i in range(len(self.df))]
            self.columns = self.df.columns.tolist()
            need_rearrangement = True
        if not re.search(r'transl|target', test_check, flags=re.IGNORECASE):
            raise TranslationColumnError(self.path)

        if len(self.columns) != 8:                                  # 7 колонок с данными + is_header = 8
            raise ColumnCountError(self.path)

        # проверка порядка столбцов
        if not need_rearrangement:
            for n, col in enumerate(self.columns):
                if n == 1 and not re.search(r'part\s?num', col, flags=re.IGNORECASE):
                    need_rearrangement = True
                    break
                if n == 2 and not re.search(r'qty', col, flags=re.IGNORECASE):
                    need_rearrangement = True
                    break
                if n == 3 and not re.search(r'unit', col, flags=re.IGNORECASE):
                    need_rearrangement = True
                    break
                if n == 4 and not re.search(r'desc', col, flags=re.IGNORECASE):
                    need_rearrangement = True
                    break
                if n == 5 and not re.search(r'transl|target', col, flags=re.IGNORECASE):
                    need_rearrangement = True
                    break
                if n == 6 and not re.search(r'comment', col, flags=re.IGNORECASE):
                    need_rearrangement = True
                    break

        # перестановка
        if need_rearrangement:
            for n, col in enumerate(self.columns):
                if re.search(r'part\s?num', col, flags=re.IGNORECASE):
                    part = col
                if re.search(r'qty', col, flags=re.IGNORECASE):
                    qty = col
                if re.search(r'unit', col, flags=re.IGNORECASE):
                    unit = col
                if re.search(r'transl|target', col, flags=re.IGNORECASE):
                    target = col
                if re.search(r'comment', col, flags=re.IGNORECASE):
                    comment = col
            self.df = self.df[[self.columns[0], part, qty, unit, self.desc_col, target, comment, 'is_header']]
            self.columns = self.df.columns.tolist()

        # проверка некодируемых символов
        ser = pd.concat([self.df[self.columns[0]], self.df[self.desc_col]])
        ser.dropna(inplace=True)
        ser = set(''.join(ser.tolist()))
        control = []
        for i in ser:
            try:
                i.encode('cp1251')
            except UnicodeEncodeError:
                control.append(i)
        if control:
            raise UnencodableError(self.path, control)

        # проверка начала датафрейма: пустая строка либо партномер
        check_cell = self.df.iloc[0, 0]
        if not isNaN(check_cell):
            if not re.search(PARTNUM_REGEX, check_cell):
                raise RowContentError(self.path)

        # проверка наличия наименований разделов в таблице
        for i in range(6):
            if i % 2 == 0:                                                      # четная строка == партномер
                if not isNaN(self.df.iloc[i, 0]):
                    if not re.search(PARTNUM_REGEX, self.df.iloc[i, 0]):
                        raise NoTOCError(self.path)
            else:                                                               # нечетная строка == заголовок раздела
                if not re.search(DESC_REGEX, self.df.iloc[i, 0]):
                    raise NoTOCError(self.path)

        # file_folder = os.path.split(self.path)[0]
        dest_folder = str(0).zfill(4) + ' ' + 'NA' + os.path.split(self.path)[1][:-5]
        dest_path = os.path.join(self.target, dest_folder)
        os.mkdir(dest_path)
        part_nums = []
        headers = []
        numbers = []
        folder_count = 0
        for row_num, header in enumerate(self.df[self.columns[0]]):

            if self.df['is_header'][row_num] and re.search(DESC_REGEX, header, flags=re.IGNORECASE):
                headers.append(header)
                numbers.append(row_num)
                part_nums.append(self.df.iloc[row_num - 1, 0])
                if len(numbers) > 1:
                    folder_count += 1
                    start = numbers[-2] + 1
                    end = numbers[-1] - 1
                    df_to_write = self.df.iloc[start:end].copy()
                    folder_name = str(folder_count).zfill(4) + \
                                  " " + get_partnum(part_nums[-2]) + \
                                  " " + clean_str(headers[-2], xlsx_name=False)
                    name = clean_str(headers[-2], xlsx_name=True)
                    os.mkdir(os.path.join(dest_path, folder_name))
                    df_to_write.to_excel(os.path.join(dest_path, folder_name, name + '.xlsx'),
                                         columns=self.columns[:-1],                        # все колонки кроме is_header
                                         index=False)
                    with open(os.path.join(dest_path, folder_name, 'name.txt'), 'w') as f:
                        f.write(headers[-2])        # полное наименование страницы записывается в текстовый файл в папке

        folder_count += 1
        start = numbers[-1] + 1
        end = len(self.df)
        df_to_write = self.df.iloc[start:end].copy()
        folder_name = str(folder_count).zfill(4) + \
                      " " + get_partnum(part_nums[-1]) + \
                      " " + clean_str(headers[-1], xlsx_name=False)
        name = clean_str(headers[-1], xlsx_name=True)
        os.mkdir(os.path.join(dest_path, folder_name))
        df_to_write.to_excel(os.path.join(dest_path, folder_name, name + '.xlsx'),
                             columns=self.columns[:-1],                                    # все колонки кроме is_header
                             index=False)
        with open(os.path.join(dest_path, folder_name, 'name.txt'), 'w') as f:
            f.write(headers[-1])                    # полное наименование страницы записывается в текстовый файл в папке

if __name__ == "__main__":
    curr_path = os.getcwd()

    filenames = os.listdir(curr_path)
    files = []
    for file in filenames:
        if file.endswith('.xlsx') and file.count('~') == 0:
            files.append(file)
    if len(files) != 1:
        print(f'Файлы: {files}')
        raise Exception

    file = files[0]
    Test = Table(file, curr_path)
    # Test.process()
    Test.clean_df()
    print(Test.df)
    Test.align_headers()
    print(Test.df)
    Test.del_mul_headers()
    print(Test.df)
    Test.fix_headers()
    print(Test.df)



