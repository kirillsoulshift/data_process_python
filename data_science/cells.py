import re
import unicodedata

"""


"""
# TODO пересмотреть правила перевода каталогов, нужен регламент, переводить длинные позиции бессмыслено они не влазят в
# заголовок
PARTNUM_REGEX = r'^\w?[\d\s-]{5,}\w?$|^FIG\.'
DESC_REGEX = r'[A-ZА-Яa-zа-я\s()]{3,}'
SUBASSEMBLY_REGEX = r'^[\s-]+'
ITEM_NUMBER_REGEX = r'^\w?[\d-]+\w?$'





def find_bool(s):
    if s == s:
        if type(s) != bool:
            if re.search(r'истина|true', s, flags=re.IGNORECASE):
                return True
            elif re.search(r'ложь|false', s, flags=re.IGNORECASE):
                return False
    else:
        raise ValueError('В столбце is_header содержатся пропущенные значения')

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
    :param el: str
    :param xlsx_name: является ли строка именем эксель файла
    :return: str
    """
    if xlsx_name:
        return el[0:3]  # далее название .xlsx файла не используется и может быть любым
    else:
        return el[0:10].strip() # наименование кратко записывается в название папки


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
        return f'Столбцы словаря {self.file} содержат пропущенные значения\n' \
               f'Столбцы словаря {self.file} содержат идентичные значения (source=target)'


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

    def __init__(self, value, col_name):

        if value == value:
            self.val = value
            if re.search(PARTNUM_REGEX, self.val, flags=re.IGNORECASE):
                self.type = 'part_num'
            elif re.search(r'qty', col_name, flags=re.IGNORECASE):
                self.type = 'qty'
            else:
                self.type = 'desc'
        else:
            self.val = None

    def is_control(self):
        """
        удаление символов управления xml
        удаление символов управления unicode
        """
        if re.search(r'&\w+;', self.val):
            self.val = re.sub(r'&amp;', ' AND ', self.val, flags=re.IGNORECASE)
            self.val = re.sub(r'&gt;', ')', self.val, flags=re.IGNORECASE)
            self.val = re.sub(r'&lt;', '(', self.val, flags=re.IGNORECASE)
            self.val = re.sub(r'&apos;', '\'', self.val, flags=re.IGNORECASE)
            self.val = re.sub(r'&quot;', '"', self.val, flags=re.IGNORECASE)
        if re.search(r'[&<>]', self.val):
            self.val = self.val.replace(r'>', ')')
            self.val = self.val.replace(r'<', '(')
            self.val = self.val.replace('&', ' AND ')
        control_chars = [True for ch in self.val if unicodedata.category(ch)[0] == 'C']
        if control_chars:
            self.val = "".join([ch for ch in self.val if unicodedata.category(ch)[0] != 'C'])

    def clean(self):

        if self.type == 'desc':
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
            if re.search(r'\(SEE.*', self.val):
                self.val = re.sub(r'\(SEE.*', '', self.val)
            # удаление ссылок на траницы каталога
            if re.search(r'\(REFER.*', self.val):
                self.val = re.sub(r'\(REFER.*', '', self.val)
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
            if re.search(r'[^\w-]', self.val):
                self.val = re.sub(r'[^\w-]', '', self.val)
            # перевод позиций в верхний регистр
            if re.search(r'[a-zа-я]', self.val):
                self.val = self.val.upper()
            self.val = self.val.strip(' -')
            return self.val

        elif self.type == 'qty':
            # очищение небуквенных символов
            if re.search(r'[^\w.,]', self.val):
                self.val = re.sub(r'[^\w.,]', '', self.val)
            # перевод позиций в верхний регистр
            if re.search(r'[a-zа-я]', self.val):
                self.val = self.val.upper()
            self.val = self.val.strip(' .,-')
            return self.val

    def wright(self):
        if self.val:
            self.is_control()
            return self.clean()
        else:
            return float('nan')


if __name__ == '__main__':
    raise UnencodableError(file='tes_file', chars_list=['a', 'v', '<'])
