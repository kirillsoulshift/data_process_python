import re
import unicodedata

"""
Модуль содержит:
 - функции приложения
 - переменные глобальных для приложения патернов определения типа ячеек, 
 - паттерны определения наименований колонок 
 - Классы эксепшенов 
 - Класс ячейки, используется дя определения типа текстовых данных и очищения
"""


# class RegexHandler():
#     """
#     Используется для хранения настроек поисковых паттернов
#     по умолчанию, взаимодействует с окном настроек, в
#     котором паттерны могут быть изменены на текущую сессию
#     """
#
#     def __init__(self):
#
#         # изменяемые паттерны для определения ячеек партномера, описания,
#         # определение символа принадлежности к подразделу, определение
#         # ячейки ссылочного номера
#         self._part_regex = r'^\(?\w{0,2}[\d\s\.,-]{3,}\w?'
#         self._desc_regex = r'[A-ZА-Яa-zа-я\s()]{3,}'
#         self._subassy_regex = r'^[\s-]+'
#         self._itemnum_regex = r'^\w?[\d-]+\w?$'
#
#         self._desc_col = r'desc|описание'
#         self._ref_col = r'id|index|ref|поз|п/п|nc'
#         self._part_col = r'part[_\s]?n|деталь[_\s]?№'
#         self._qty_col = r'qty|quant|кол-во|колич'
#         self._comm_col = r'comment|remark|коммент|примеча'
#         self._unit_col = r'unit|ед\.\s?изм\.'
#         self._transl_col = r'transl|target|перевод'
#
#     @property
#     def part_regex(self):
#         return self._part_regex
#
#     @part_regex.setter
#     def part_regex(self, a):
#         try:
#             re.compile(a)
#             self._part_regex = a
#         except re.error:
#             raise ValueError(f'Ошибка в регулярном выражении {a}')
#
#     @property
#     def desc_regex(self):
#         return self._desc_regex
#
#     @desc_regex.setter
#     def desc_regex(self, a):
#         try:
#             re.compile(a)
#             self._desc_regex = a
#         except re.error:
#             raise ValueError(f'Ошибка в регулярном выражении {a}')
#
#     @property
#     def subassy_regex(self):
#         return self._subassy_regex
#
#     @subassy_regex.setter
#     def subassy_regex(self, a):
#         try:
#             re.compile(a)
#             self._subassy_regex = a
#         except re.error:
#             raise ValueError(f'Ошибка в регулярном выражении {a}')
#
#     @property
#     def itemnum_regex(self):
#         return self._itemnum_regex
#
#     @itemnum_regex.setter
#     def itemnum_regex(self, a):
#         try:
#             re.compile(a)
#             self._itemnum_regex = a
#         except re.error:
#             raise ValueError(f'Ошибка в регулярном выражении {a}')


PARTNUM_REGEX = r'^\(?\w{0,2}[\d\s.,-]{4,}\w?' # r'^[\d\s\.,-]{4,}$|^ZEGA'  # r'^\w?[\d\s-]{5,}\w?$|^FIG\.'
DESC_REGEX = r'[A-ZА-Яa-zа-я\s()]{3,}'
SUBASSEMBLY_REGEX = r'^[\s-]+'
ITEM_NUMBER_REGEX = r'^[A-ZА-Яa-zа-я]?[\d.-]{1,3}[A-ZА-Яa-zа-я]?$'

DESC_COL = r'desc|описание|name'
REF_COL = r'id|index|item|ref|поз|п/п|nc'
PARTNUM_COL = r'part[_\s]?n|деталь[_\s]?№'
QTY_COL = r'qty|quant|кол-во|колич'
COMMENT_COL = r'comment|remark|коммент|примеча'
UNIT_COL = r'unit|ед\.\s?изм\.'
TRANSL_COL = r'transl|target|перевод'


def find_bool(s):
    """Обрабатывает ячейки с данными
    для столбца is_header"""
    if s == s:
        if type(s) != bool:
            if re.search(r'истина|true', s, flags=re.IGNORECASE):
                return True
            elif re.search(r'ложь|false', s, flags=re.IGNORECASE):
                return False
            else:
                return float('nan')
        else:
            return s
    else:
        return s


def where_bytes_ser(ser):
    """Поиск байт-стринг в серии"""
    res = []
    for item in ser:
        if type(item) == bytes:
            res.append([f'{item=}', f'{type(item)=}'])
    return res


def where_bytes_df(df):
    """Поиск байт-стринг в датафрейме"""
    res = []
    for ser in df.columns.tolist():
        res.append(where_bytes_ser(df[ser]))
    return print(res)


def check_row(row):
    """
    проверка строки датафрейма на наличие
    значения в первом столбце, а во всех остальных NaN
    :param row: строка датафрейма
    :return: bool
    """
    row = row.tolist()
    col_count = len(row)
    check_list = [True]
    for i in range(col_count - 1):
        check_list.append(False)
    f = lambda x: True if x == x else False
    row = [f(i) for i in row]
    if row == check_list:
        return True
    else:
        return False


def qck_search(regex, statement):
    """
    :param regex: str to find
    :param statement: str where to find
    :return: bool
    """
    if statement == statement:
        if re.search(regex, statement):
            return True
        else:
            return False
    else:
        return False


def is_english(s):
    """Функция поиска символов русской раскладки
    в текстовой строке"""
    if s != s:
        return True
    elif isinstance(s, bool):
        return True
    elif re.search(r'[А-Яа-я]', s):
        return False
    else:
        return True


def isNaN(num):
    return num != num


def what_prefix(df, idx, col):
    """Извлекает символ принадлежности к подразделу
    из текстовой строки по совпадению глобального
    паттерна"""
    prefix = ''
    el = df.iloc[idx, col]
    if re.search(SUBASSEMBLY_REGEX, el):
        m = re.search(SUBASSEMBLY_REGEX, el)
        prefix = m.group(0)
    return prefix


def get_partnum(partnum):

    if partnum == '' or partnum != partnum:
        return 'NA'
    else:
        return partnum


def clean_str(el, xlsx_name: bool):
    """
    Используется для преобразования наименования папки или
    .xlsx файла
    :param el: str
    :param xlsx_name: является ли строка именем эксель файла
    :return: str
    """
    if xlsx_name:
        return re.sub(r'[\\/|:;*?"]', '', el[0:3].strip())  # далее название .xlsx файла не используется и может быть любым
    else:
        return re.sub(r'[\\/|:;*?"]', '', el[0:10].strip()) # наименование кратко записывается в название папки


class UnencodableError(Exception):

    def __init__(self, file, chars_list, *args):
        super().__init__(args)
        self.chars = chars_list
        self.file = file

    def __str__(self):
        return f'Некодируемые символы в файле {self.file}\n' \
               f'Некодируемые символы: {self.chars}'


class DictError(Exception):

    def __init__(self, dict_file, *args):
        super().__init__(args)
        self.file = dict_file

    def __str__(self):
        return f'Столбцы словаря {self.file} содержат пропущенные значения.\n' \
               f'Столбцы словаря {self.file} содержат идентичные значения (source=target).'


class ColumnNameError(Exception):

    def __init__(self, xl_file, *args):
        super().__init__(args)
        self.file = xl_file

    def __str__(self):
        return f'Столбцы исходного файла {self.file} должны иметь названия схожие с:\n' \
               f'Столбец партномеров = Part Number\n' \
               f'Столбец количество = Qty\n' \
               f'Столбец описания = Description'


class RowContentError(Exception):

    def __init__(self, xl_file, *args):
        super().__init__(args)
        self.file = xl_file

    def __str__(self):
        return f'Начало таблицы {self.file} не является партномером или пустой строкой'


class NoTOCError(Exception):

    def __init__(self, xl_file, *args):
        super().__init__(args)
        self.file = xl_file

    def __str__(self):
        return f'Начало таблицы {self.file} не содержит заголовков разделов каталога'


class TranslationColumnError(Exception):

    def __init__(self, xl_file, *args):
        super().__init__(args)
        self.file = xl_file

    def __str__(self):
        return f'Исходный файл {self.file} не содержит столбец перевода с названием' \
               f'Translation или Target'


class ColumnCountError(Exception):

    def __init__(self, xl_file, *args):
        super().__init__(args)
        self.file = xl_file

    def __str__(self):
        return f'Исходный файл {self.file} не содержит колонки c данными' \
               f'в требующемся количестве и в порядке:\n' \
               f"'Item ID','Part Number','Quantity','Unit','Description','Translation','Comment'"


class Cell():

    def __init__(self, value, col_name, slash_header = False):
        """
        Класс работает с каждой ячейкой индивидуально, производится очищщение и запись ячеки в датафрейм
        :param value: строка, содержимое ячейки эксель
        :param col_name: наименование столбца к которому принадлежит данная ячейка
        :param slash_header: требуется ли учитывать написание заголовков на двух языках через '/'
        """

        if value == value:
            self.col_name = col_name
            self.val = value
            self.slash_header = slash_header
            if re.search(PARTNUM_REGEX, self.val, flags=re.I) and \
                (re.search(REF_COL, self.col_name, flags=re.I) or
                 re.search(PARTNUM_COL, self.col_name, flags=re.I)):
                self.type = 'part_num'
            elif re.search(QTY_COL, self.col_name, flags=re.I):
                self.type = 'qty'
            else:
                self.type = 'desc'
        else:
            self.val = None

    def is_control(self):
        """
        удаление символов управления xml
        удаление символов управления unicode
        удаление ¤
        удаление китайских иероглифов
        """
        if re.search(r'&\w+;', self.val):
            self.val = re.sub(r'&amp;', ' AND ', self.val, flags=re.IGNORECASE)
            self.val = re.sub(r'&gt;', ')', self.val, flags=re.IGNORECASE)
            self.val = re.sub(r'&lt;', '(', self.val, flags=re.IGNORECASE)
            self.val = re.sub(r'&apos;', "'", self.val, flags=re.IGNORECASE)
            self.val = re.sub(r'&quot;', '"', self.val, flags=re.IGNORECASE)
        if re.search(r'[&<>]', self.val):
            self.val = self.val.replace(r'>', ')')
            self.val = self.val.replace(r'<', '(')
            self.val = self.val.replace('&', ' AND ')
        control_chars = [True for ch in self.val if unicodedata.category(ch)[0] == 'C']
        if control_chars:
            self.val = "".join([ch for ch in self.val if unicodedata.category(ch)[0] != 'C'])
        if re.search(r'[¤。º]', self.val):
            self.val = re.sub('[¤。º]', '', self.val)
        if re.search(u'[⺀-⺙⺛-⻳⼀-⿕々〇〡-〩〸-〺〻㐀-䶵一-鿃豈-鶴侮-頻並-龎]', self.val, re.U):
            self.val = re.sub(u'[⺀-⺙⺛-⻳⼀-⿕々〇〡-〩〸-〺〻㐀-䶵一-鿃豈-鶴侮-頻並-龎]', '', self.val)

    def clean(self):

        if self.type == 'desc':

            # удалять лидирующие нули в референс намбер
            if re.search(REF_COL, self.col_name, flags=re.I) and \
                    re.search(ITEM_NUMBER_REGEX, self.val, flags=re.I) and \
                    self.val.startswith('0'):
                self.val = self.val.lstrip(' 0')

            # в случае наименования странц через '/': <ENG>/<OTHER LANGUAGE>
            # остается первая половина строки в случаем множественных '/'
            if self.slash_header:
                if re.search(REF_COL, self.col_name, flags=re.IGNORECASE) and \
                        re.search(r'[a-zA-Z\s]/[a-zA-Z\s]', self.val):
                    c = '/'
                    pos_list = [pos for pos, char in enumerate(self.val) if char == c]
                    if len(pos_list) != 1:
                        pos_to_split = pos_list[int(len(pos_list) / 2)]
                    else:
                        pos_to_split = pos_list[0]
                    self.val = self.val[:pos_to_split].strip()

            #\) \ )
            if re.search(r'\\\s?\)|\\\s?\(', self.val):
                self.val = re.sub(r'\\\s?\)|\\\s?\(', '', self.val)

            # oring{oil
            if re.search(r'^.*\w[{}]\w.*$', self.val):
                self.val = re.sub(r'[{}]', ', ', self.val)

            # перевод позиций в верхний регистр
            if re.search(r'[a-zа-я]', self.val):
                self.val = self.val.upper()
            # очищение небуквенных символов в начале и конце строки
            if re.search(r'^[^\w.(-\[]|[^\w)\]]$', self.val):
                self.val = re.sub(r'^[^\w.(-\[]+|[^\w)\]]+$', '', self.val)
            # очищение многоточий
            if re.search(r'\.{2,}', self.val):
                self.val = re.sub(r'\.{2,}', '', self.val)
            # замена " ," на ","
            if self.val.count(' ,'):
                self.val = self.val.replace(' ,', ',')
            # замена "," с пропущенными пробелами
            if re.search(r'\d,[A-ZА-Я]|[A-ZА-Я],\d|[A-ZА-Я],[A-ZА-Я]|\),\(|\),\w|\w,\(', self.val):
                m = re.findall(r'\d,[A-ZА-Я]|[A-ZА-Я],\d|[A-ZА-Я],[A-ZА-Я]|\),\(|\),\w|\w,\(', self.val)
                m_rev = [i.replace(',', ', ') for i in m]
                for num, item in enumerate(m):
                    self.val = self.val.replace(item, m_rev[num])
            # замена '\w\(' и '\)\w'
            if re.search(r'[^\s]\(|\)[^\s]', self.val):
                self.val = re.sub(r'\)', ') ', self.val)
                self.val = re.sub(r'\(', ' (', self.val)
            # замена '\( \w' и '\w \)'
            if re.search(r'\(\s\w|\w\s\)', self.val):
                self.val = re.sub(r'\(\s', '(', self.val)
                self.val = re.sub(r'\s\)', ')', self.val)
            # замена ' . '
            if re.search(r'\s\.\s|[A-Z)]{2,}\.[A-Z(]{2,}|\s\.[A-Z]{2,}', self.val) and \
                    not re.search(r'WWW\.|HTTP', self.val):
                m = re.findall(r'\s\.\s|[A-Z)]{2,}\.[A-Z(]{2,}|\s\.[A-Z]{2,}', self.val)
                m_rev = [i.replace('.', ' ') for i in m]
                for num, item in enumerate(m):
                    self.val = self.val.replace(item, m_rev[num])
            # замена нескольких "-" и "—" на "-"
            if re.search(r'—+|[—-]{2,}', self.val):
                self.val = re.sub(r'—+|[—-]{2,}', '-', self.val)
            # замена "-" без пробелов на " - "
            if re.search(r'[^\s]-[^\s]|[^\s]-\s|\s-[^\s]', self.val):
                split_list = self.val.split('-')
                correct_list = []
                for i in split_list:
                    correct_list.append(i.strip())
                self.val = " - ".join(correct_list)
            if self.val.count('O - RING'):
                self.val = self.val.replace('O - RING', 'O-RING')
            # удаление ссылок на рисунки каталога
            if re.search(r'\(?SEE.*', self.val):
                self.val = re.sub(r'\(?SEE.*', '', self.val)
            # удаление ссылок на траницы каталога
            if re.search(r'\(?REFER.*', self.val):
                self.val = re.sub(r'\(?REFER.*', '', self.val)
            # замена нескольких " " одним " "
            if re.search(r'\s{2,}', self.val):
                self.val = re.sub(r'\s{2,}', ' ', self.val)
            self.val = self.val.strip(' ,.')
            return self.val

        elif self.type == 'part_num':
            if self.val.count('FIG.'):
                self.val = self.val.replace('FIG.', '')
            # замена нескольких "-" и "—" на "-"
            if re.search(r'—+|[—-]{2,}', self.val):
                self.val = re.sub(r'—+|[—-]{2,}', '-', self.val)
            # очищение небуквенных символов
            if re.search(r'[^\w.,-]', self.val):
                self.val = re.sub(r'[^\w.,-]', '', self.val)
            # перевод позиций в верхний регистр
            if re.search(r'[a-zа-я]', self.val):
                self.val = self.val.upper()
            # удаление цифр заключенных в скобки в конце партномера
            if re.search(r'\(\d{0,3}\)', self.val):
                self.val = re.sub(r'\(\d{0,3}\)', '', self.val)
            self.val = self.val.strip(' .,_-—()[]')
            return self.val

        elif self.type == 'qty':
            # очищение небуквенных символов
            if re.search(r'[^\w.,]', self.val):
                self.val = re.sub(r'[^\w.,]', '', self.val)
            # перевод позиций в верхний регистр
            if re.search(r'[a-zа-я]', self.val):
                self.val = self.val.upper()
            if self.val.startswith('0'):
                self.val = self.val.lstrip('0')
            self.val = self.val.strip(' .,_-—()[]')
            return self.val

    def wright(self):
        if self.val:
            self.is_control()
            return self.clean()
        else:
            return float('nan')
