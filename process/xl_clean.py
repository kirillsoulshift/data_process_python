import pandas as pd
import re
import os
import sys
import time
import functools
import unicodedata


"""
ПРИМЕР ИСХОДНОГО .xlsx (выгрузка из OCR-проекта)
_________________________________________________________
 1 | REF.  | PART NO.  |  QTY.  |  UNIT  |  DESCRIPTION  |  # первая строка - наименовая столбцов
___|_______|___________|________|________|_______________|
 2 |       |           |        |        |               |
___|_______|___________|________|________|_______________|
 3 |577932 |           |        |        |               |  # каталожный номер страницы
___|_______|___________|________|________|_______________|
 4 |DECK   |           |        |        |               |  # заголовок страницы
___|_______|___________|________|________|_______________|  # может содержать пустые строки между
 5 |001----|-57775876..|1.......|EA-     |---DECKING/RAIL|  # названием страницы и спецификацией 
___|_______|___________|________|________|_______________|  # может содержать лишние ненужные
 6 |       |-57775884..|1.......|EA-     |------• _ DECK |  # наименовая столбцов
___|_______|___________|________|________|_______________|
 7 |002----|-57779302..|1.......|EA-     |---DECKING/DCS |
___|_______|___________|________|________|_______________|
   |       |           |        |        |               |
                      ...              
                      ...                
___|_______|___________|________|________|_______________|
38 |027----|-57774733..|4.......|EA-     |---WASHER,LOCK |  # конец первой страницы спецификации
___|_______|___________|________|________|_______________|            
   |       |           |        |        |
                      ...
           <любое кол-во пустых строк>              
                      ...
___|_______|___________|________|________|_______________|
45 |577305 |           |        |        |               |  # начало следующей страницы спецификации
___|_______|___________|________|________|_______________|
46 |DECKING|           |        |        |               |
___|_______|___________|________|________|_______________|
47 |001----|-57779386..|1.......|EA-     |---DECKING,ASSY|
___|_______|___________|________|________|_______________|                      
   |       |           |        |        |               |
                      ...              
                      ...                                     
                      ...               
порядок столбцов файлов .xlsx:

 ================================================================================================================
           >>>'Item ID', 'Part Number', 'Quantity', 'Unit', 'Description', 'Translation', 'Comment'<<<  
 ================================================================================================================   
                                  
Обрабатывает один .xlsx файл в папке в которой находится скрипт

ПЕРВАЯ СТРОКА ЭКСЕЛЯ = НАИМЕНОВАНИЯ КОЛОНОК

Удалить лишние и мусорные символы в начале и конце
строки, двойные пробелы, исправить
запятые и тире без пробелов итд. в столбцах .xlsx
удалить наименования столбцов таблиц внутри спецификаций
удалить повторяющиеся заголовки
проверить правильность заполнения 
заголовков и партномеров

скрипт формирует новый файл .xlsx в своей директории
"""


def is_called(func):
    """
    Проверка вызова функции
    """
    @functools.wraps(func)
    def wrapper(*args):
        wrapper.has_been_called = True
        return func(*args)
    wrapper.has_been_called = False
    return wrapper


def is_xml_control(el):
    """
    удаление символов управления xml
    """
    if re.search(r'&\w+;', el):
        el = el.replace(r'&AMP;', ' AND ')
        el = el.replace(r'&gt;', ')')
        el = el.replace(r'&lt;', '(')
        el = el.replace(r'&apos;', '\'')
        el = el.replace(r'&quot;', '"')

    if re.search(r'[&<>]', el):
        el = el.replace(r'>', ')')
        el = el.replace(r'<', '(')
        el = el.replace('&', ' AND ')
    return el


def is_unicode_control(el):
    """
    удаление символов управления unicode
    """
    control_chars = [True for ch in el if unicodedata.category(ch)[0] == 'C']
    if control_chars:
        el = "".join([ch for ch in el if unicodedata.category(ch)[0] != 'C'])
    return el


@is_called
def choose_replace(el, col):
    if el == el:
        if re.search(r'^[a-zа-я]?[\d\s-]{5,}|^[\d-]{3,}|^FIG\.', el, flags=re.IGNORECASE):
            return replace_partnumber(el, col)
        else:
            return replace_description(el, col)
    else:
        return el


@is_called
def replace_qty(el, col):
    if el == el:

        # очищение небуквенных символов
        if re.search(r'[^\w.,]', el):
            el = re.sub(r'[^\w.,]', '', el)

        # перевод позиций в верхний регистр
        if re.search(r'[a-zа-я]', el):
            el = el.upper()

        el = el.strip(' .,')

        return el
    else:
        return el


@is_called
def replace_partnumber(el, col):
    if el == el:

        if el.count('FIG.'):
            el = el.replace('FIG.', '')

        # замена нескольких "-" и "—" на "-"
        if re.search(r'—+|[—-]{2,}', el):
            el = re.sub(r'—+|[—-]{2,}', '-', el)

        # очищение небуквенных символов
        if re.search(r'[^\w-]', el):
            el = re.sub(r'[^\w-]', '', el)

        # перевод позиций в верхний регистр
        if re.search(r'[a-zа-я]', el):
            el = el.upper()

        el = el.strip(' -')

        return el
    else:
        return el


@is_called
def replace_description(el, col):

    if el == el:

        # замена нескольких " " одним " "
        if re.search(r'\s{2,}', el):
            el = re.sub(r'\s{2,}', ' ', el)

        # перевод позиций в верхний регистр
        if re.search(r'[a-zа-я]', el):
            el = el.upper()

        if re.search(r'Я', el):
            el = re.sub(r'Я', ' ', el)

        if re.search(r'[¤\\]', el):
            el = re.sub(r'[¤\\]', '', el)

        if re.search(r'\{', el):
            el = re.sub(r'\{', '(', el)

        if re.search(r'}', el):
            el = re.sub(r'}', ')', el)

        if re.search(r'\({2,}', el):
            el = re.sub(r'\({2,}', '(', el)

        if re.search(r'\){2,}', el):
            el = re.sub(r'\){2,}', ')', el)

        # очищение символов управления
        el = is_unicode_control(el)

        if re.search(r'[&<>]', el):
            el = is_xml_control(el)

        # очищение небуквенных символов в начале и конце строки
        if re.search(r'^[^\w.(-\[]|[^\w)\]]$', el):
            el = re.sub(r'^[^\w.(-\[]+|[^\w)\]]+$', '', el)

        if re.search(r'\.{2,}', el):
            el = re.sub(r'\.{2,}', '', el)

        # замена "," с пропущенными пробелами
        if re.search(r'\d,[A-ZА-Я]|[A-ZА-Я],\d|[A-ZА-Я],[A-ZА-Я]', el):
            m = re.findall(r'\d,[A-ZА-Я]|[A-ZА-Я],\d|[A-ZА-Я],[A-ZА-Я]', el)
            m_rev = [i.replace(',', ', ') for i in m]
            for num, item in enumerate(m):
                el = el.replace(item, m_rev[num])

        # замена '\w\(' и '\)\w'
        if re.search(r'[^\s]\(|\)[^\s]', el):
            el = re.sub(r'\)', ') ', el)
            el = re.sub(r'\(', ' (', el)

        # замена '\( \w' и '\w \)'
        if re.search(r'\(\s\w|\w\s\)', el):
            el = re.sub(r'\(\s', '(', el)
            el = re.sub(r'\s\)', ')', el)

        # замена ' . '
        if re.search(r'\s\.\s|[A-Z)]{2,}\.[A-Z(]{2,}|\s\.[A-Z]{2,}', el) and\
          not re.search(r'WWW\.|HTTP', el):
            m = re.findall(r'\s\.\s|[A-Z)]{2,}\.[A-Z(]{2,}|\s\.[A-Z]{2,}', el)
            m_rev = [i.replace('.', ' ') for i in m]
            for num, item in enumerate(m):
                el = el.replace(item, m_rev[num])

        # замена нескольких "-" и "—" на "-"
        if re.search(r'—+|[—-]{2,}', el):
            el = re.sub(r'—+|[—-]{2,}', '-', el)

        # замена "-" без пробелов на " - "
        if re.search(r'[^\s]-[^\s]|[^\s]-\s|\s-[^\s]', el):
            split_list = el.split('-')
            correct_list = []
            for i in split_list:
                correct_list.append(i.strip())
            el = " - ".join(correct_list)

        if el.count('O - RING'):
            el = el.replace('O - RING', 'O-RING')

        # удаление ссылок на рисунки каталога
        if re.search(r'\(SEE.*', el):
            el = re.sub(r'\(SEE.*', '', el)

        # удаление ссылок на траницы каталога
        if re.search(r'\(REFER.*', el):
            el = re.sub(r'\(REFER.*', '', el)

        # замена " ," на ","
        if el.count(' ,'):
            el = el.replace(' ,', ',')

        # замена нескольких " " одним " "
        if re.search(r'\s{2,}', el):
            el = re.sub(r'\s{2,}', ' ', el)

        el = el.strip(' ,.')

        return el
    else:
        return el


def is_english(s):
    if s != s:
        return True
    elif re.search(r'[А-Я]', s):
        return False
    else:
        return True


def check_row(row):
    """
    проверка строки датафрейма на наличие
    значения в первом столбце, а во всех остальных NaN
    :param row: строка датафрейма <list>
    :return: bool
    """
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


@is_called
def del_mul_headers(df):
    """
    удалить заголовки, идентичные
    заголовку выше
    проверка по партномерам и заголовкам производится
    отдельно
    :param df: DataFrame
    :return: DataFrame
    """

    data = df.values.tolist()
    headers = []
    for item in data:
        if check_row(item):
            headers.append(item[0])
    ser = pd.Series(headers, copy=False)
    mul_headers = ser.value_counts()[lambda x: x > 1].index.tolist()
    headers_to_del = []
    partnums_to_del = []
    for n, row in enumerate(data):
        if row[0] in mul_headers and \
        re.search(r'[A-ZА-Я\s]{4,}|^[A-ZА-Я]{3,}$', row[0]) and \
        row[0] not in headers_to_del[-1:]:                         # заголовок отличается от предыдущего
            headers_to_del.append(row[0])                          # добавление заголовка в чек-лист
        elif row[0] in mul_headers and \
        re.search(r'^\w?[\d-]+\w?$', row[0]) and \
        row[0] not in partnums_to_del[-1:]:                        # партномер отличается от предыдущего
            partnums_to_del.append(row[0])                         # добавление партномера в чек-лист
        elif row[0] in mul_headers and \
        re.search(r'[A-ZА-Я\s]{4,}|^[A-ZА-Я]{3,}$', row[0]) and \
        row[0] in headers_to_del[-1:]:                             # заголовок равен предыдущему
            df.drop(n, inplace=True)                               # удаление текущей строки из датафрейма
        elif row[0] in mul_headers and \
        re.search(r'^\w?[\d-]+\w?$', row[0]) and \
        row[0] in partnums_to_del[-1:]:                            # партномер равен предыдущему
            df.drop(n, inplace=True)                               # удаление текущей строки из датафрейма
    df.reset_index(drop=True, inplace=True)
    return df


@is_called
def clean_df(df):
    """
    Удалить строку в датафрейме
    совпадающую с наименованиями
    колонок

    :param df: dataframe
    :return: dataframe
    """

    df.reset_index(drop=True, inplace=True)

    for col in df.columns.to_list():
        if re.search(r'desc|назв', col, flags=re.IGNORECASE):
            result_col = col
    check_col = df[result_col].tolist()
    nums_to_del = []
    for num_del, cell in enumerate(check_col):
        if cell == cell:
            if cell.upper() == result_col.upper():
                nums_to_del.append(num_del)
    if len(nums_to_del) != 0:
        df.drop(nums_to_del, axis=0, inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


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


@is_called
def fix_headers(df):
    """
    Проверить правильность заполнения
    партномеров и заголовков
    проверить порядок следования 'партномер -> заголовок'
    :param df: DataFrame
    :return: DataFrame
    """

    data = df.values.tolist()
    row_len = len(data[0])
    empty_row = [float('nan') for i in range(row_len)]
    headers = []
    processed = []
    data.insert(0, empty_row)
    for n, row in enumerate(data):
        if n != 0:
            if check_row(row):
                # заголовок с партномером
                if re.search(r'[A-ZА-Я\s]{4,}|^[A-ZА-Я]{3,}$', row[0]) \
                        and qck_search(r'^\w?[\d-]+\w?$', data[n - 1][0]):
                    processed.append(row)
                    headers.append(row)
                # партномер с заголовком
                elif re.search(r'^\w?[\d-]+\w?$', row[0]) \
                        and qck_search(r'[A-ZА-Я\s]{4,}|^[A-ZА-Я]{3,}$', data[n + 1][0]):
                    processed.append(row)
                # заголовок без партномера
                elif re.search(r'[A-ZА-Я\s]{4,}|^[A-ZА-Я]{3,}$', row[0]) \
                        and not qck_search(r'^\w?[\d-]+\w?$', data[n - 1][0]):
                    processed.append(empty_row)
                    processed.append(row)
                    headers.append(row)
                # партномер без заголовка
                elif re.search(r'^\w?[\d-]+\w?$', row[0]) \
                        and not qck_search(r'[A-ZА-Я\s]{4,}|^[A-ZА-Я]{3,}$', data[n + 1][0]):
                    processed.append(row)
                    processed.append(headers[-1])

            else:
                processed.append(row)

    new_df = pd.DataFrame(processed)
    new_df.columns = df.columns.to_list()
    return new_df


@is_called
def dub_in_a_row(df):
    """
    Проверка на наличие нескольких
    одинаковых заголовков подряд
    """

    data = df.values.tolist()
    headers = []
    headers_to_rev = []
    for n, row in enumerate(data):
        if check_row(row) and re.search(r'[A-ZА-Я\s]{4,}|^[A-ZА-Я]{3,}$', row[0]):
            if row[0] not in headers[-1:]:
                headers.append(row[0])
            else:
                headers_to_rev.append([n, row[0]])
    if headers_to_rev:
        review = pd.DataFrame(headers_to_rev)
        review.columns = ['Num.', 'Header']
        review.drop_duplicates(inplace=True)
        print('\033[91m' + 'Заголовки следующие подряд:' + '\033[0m')
        print(review)


@is_called
def align_headers(df):
    """
    поиск заголовков ошибочно разделенных на несколько
    смежных ячеек по горизонтали;
    соединение в одну ячейку
    """

    data = df.values.tolist()
    fractured_headers = []
    for n, row in enumerate(data):
        if check_row(row) and \
        re.search(r'[A-ZА-Я\s()]{3,}', row[0]) and \
        not qck_search(r'[A-ZА-Я\s()]{3,}', data[n-1][0]) and \
        qck_search(r'[A-ZА-Я\s()]{3,}', data[n+1][0]):        # фрагмент заголовка является началом
            fractured_headers.append([row[0], n, 'start'])

        elif check_row(row) and \
        re.search(r'[A-ZА-Я\s()]{3,}', row[0]) and \
        qck_search(r'[A-ZА-Я\s()]{3,}', data[n - 1][0]) and \
        qck_search(r'[A-ZА-Я\s()]{3,}', data[n + 1][0]):      # фрагмент заголовка является продолжением
            fractured_headers.append([row[0], n, 'continuation'])

        elif check_row(row) and \
        re.search(r'[A-ZА-Я\s()]{3,}', row[0]) and \
        qck_search(r'[A-ZА-Я\s()]{3,}', data[n - 1][0]) and \
        not qck_search(r'[A-ZА-Я\s()]{3,}', data[n + 1][0]):      # фрагмент заголовка является окончанием
            fractured_headers.append([row[0], n, 'continuation'])

    rows_for_del = [item[1] for item in fractured_headers if item[2] == 'continuation']
    for num, item in enumerate(fractured_headers):
        if item[2] == 'start':
            statement = item[0]
            count = num
            while len(fractured_headers) > count + 1 and fractured_headers[count + 1][2] == 'continuation':
                statement += ' ' + fractured_headers[num + 1][0]
                count += 1

            df.iloc[item[1], 0] = statement
    df.drop(index=rows_for_del, inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


@is_called
def is_russian_chars(df):
    """
    проверка наличия русских символов в столбцах
    """
    for col in df.columns:
        test_series = df[col].copy()

        if not test_series.apply(is_english).all():
            to_print = test_series.apply(is_english).loc[lambda x: x == False].copy()

            print('\033[91m' + f'=====Символы русской раскладки в столбце: <{col}> =====' + '\033[0m')
            print(df[col].iloc[to_print.index.tolist()].drop_duplicates())
        else:
            print(f'=====Символы русской раскладки в столбце: <{col}> не обнаружены=====')


@is_called
def partnum_len(df):
    """
    проверка длин всех партномеров в символах и вывод результата
    """
    part_num_col = ''
    for col in df.columns:
        if re.search(r'part[\s_]?num', col, flags=re.IGNORECASE):
            part_num_col = col
    if part_num_col:
        res = df[part_num_col].dropna().apply(lambda x: len(x)).value_counts()
        df_values = pd.DataFrame({'Количество знаков партномера': res.index, 'Количество парномеров': res.values})
        print(df_values)
    else:
        print('\033[91m' + 'Столбец Part Number не обнаружен' + '\033[0m')





start_time = time.time()

curr_path = os.getcwd()
filenames = os.listdir(curr_path)
files = []
for file in filenames:
    if file.endswith('.xlsx') and file.count('~') == 0:
        files.append(file)
if len(files) != 1:
    print(f'Файлы: {files}')
    sys.exit(f'Файл для обработки в пути {curr_path} не определен')

file = files[0]
df = pd.read_excel(os.path.join(curr_path, file), header=0, dtype=str)
df.dropna(how='all', inplace=True)  # удалить пустые строки экселя
df.reset_index(drop=True, inplace=True)
columns_name = df.columns

test_check = ''.join(columns_name)
if not re.search(r'(?=.*part[\s_]?num)(?=.*(qty|quantity))(?=.*desc)', test_check, flags=re.IGNORECASE):
    print('\033[91m' + 'Столбцы исходного файла должны иметь названия схожие с:' + '\033[0m')
    print('\033[91m' + 'Столбец партномеров = Part Number' + '\033[0m')
    print('\033[91m' + 'Столбец количество = Qty' + '\033[0m')
    print('\033[91m' + 'Столбец описания = Description' + '\033[0m')
    print('\033[91m' + 'Первый столбец в таблице = Заголовки и номера позиций' + '\033[0m')
    sys.exit()

_vars = {}
for n, col in enumerate(columns_name):

    if n == 0:
        _vars[col] = [choose_replace(i, n) for i in df[col].tolist()]

    elif re.search(r'part[\s_]?num|unit', col, flags=re.IGNORECASE):
        _vars[col] = [replace_partnumber(i, n) for i in df[col].tolist()]

    elif re.search(r'qty|quantity', col, flags=re.IGNORECASE):
        _vars[col] = [replace_qty(i, n) for i in df[col].tolist()]

    else:
        _vars[col] = [replace_description(i, n) for i in df[col].tolist()]

for col in columns_name:
    df[col] = _vars[col]

# df = clean_df(df)                   # удалить строки заголовков таблиц
# df = align_headers(df)              # соединить заголовок разделенный на несколько соседних ячеек
# df = del_mul_headers(df)            # удалить повторяющиеся заголовки и каталожные номера заголовков
# df = fix_headers(df)                # проверка правильности заполнения заголовков и каталожных номеров
# dub_in_a_row(df)                    # проверка на наличие нескольких одинаковых заголовков подряд
is_russian_chars(df)                # проверка на наличие символов русской раскладки
# partnum_len(df)                     # проверка длин партномеров

# проверка вызова функций
if not (choose_replace.has_been_called and
        replace_qty.has_been_called and
        replace_partnumber.has_been_called and
        replace_description.has_been_called and
        del_mul_headers.has_been_called and
        clean_df.has_been_called and
        fix_headers.has_been_called and
        dub_in_a_row.has_been_called and
        is_russian_chars.has_been_called and
        partnum_len.has_been_called and
        align_headers.has_been_called):
    if not choose_replace.has_been_called:
        print('\033[91m' + 'Не использовалась функция выбора алгоритма очистки (для 1-ого столбца)' + '\033[0m')
    if not replace_qty.has_been_called:
        print('\033[91m' + 'Не использовалась функция очистки столбца "количество"' + '\033[0m')
    if not align_headers.has_been_called:
        print('\033[91m' + 'Не выполнялась проверка и соединение заголовков находящихся в нескольких ячейках' + '\033[0m')
    if not replace_partnumber.has_been_called:
        print('\033[91m' + 'Не использовалась функция очистки столбца "партномер"' + '\033[0m')
    if not replace_description.has_been_called:
        print('\033[91m' + 'Не использовалась функция очистки столбца "описание"' + '\033[0m')
    if not del_mul_headers.has_been_called:
        print('\033[91m' + 'Не использовалась функция удаления повторяющихся заголовков и каталожных номеров' + '\033[0m')
    if not clean_df.has_been_called:
        print('\033[91m' + 'Не использовалась функция удаления наименований столбцов внутри таблиц' + '\033[0m')
    if not fix_headers.has_been_called:
        print('\033[91m' + 'Не выполнялась проверка правильности заполнения исходного файла' + '\033[0m')
    if not dub_in_a_row.has_been_called:
        print('\033[91m' + 'Не выполнялась проверка на наличие одинаковых заголовков идущих подряд' + '\033[0m')
    if not is_russian_chars.has_been_called:
        print('\033[91m' + 'Не выполнялась проверка на наличие символов русской раскладки' + '\033[0m')
    if not partnum_len.has_been_called:
        print('\033[91m' + 'Не выполнялась проверка длин партномеров' + '\033[0m')
    resp = ''
    while not re.search(r'[yn]', resp, flags=re.IGNORECASE):
        resp = input('Формировать файл .xlsx (y/n)?')
    if re.search(r'n', resp, flags=re.IGNORECASE):
        sys.exit()

df.to_excel(file[:-5] + '_cleaned.xlsx', index=False)
print("--- %s seconds ---" % (time.time() - start_time))
