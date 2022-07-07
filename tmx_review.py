import pandas as pd
import os
import sys
import re
import unicodedata

"""
Создает новый файл .tmx на основании файлов .tmx и .xlsx находящихся
в директории скрипта

.xlsx файл должен содержать колонки source | target | date
.tmx файл должен содержать рабочую выгрузку из TM для копирования шапки файла
"""


def replace_description(el):
    # очищение символов управления
    setted = set(el)
    control_chars = [True for ch in setted if unicodedata.category(ch)[0] == 'C']
    if control_chars:
        el = "".join(ch for ch in el if unicodedata.category(ch)[0] != 'C')

    # очищение небуквенных символов в начале и конце строки
    if re.search(r'^[^\w.(-]|[^\w)]$', el):
        el = re.sub(r'^[^\w.(-]+|[^\w)]+$', '', el)

    # перевод позиций в верхний регистр
    if re.search(r'[a-zа-я]', el):
        el = el.upper()

    if el.count('&'):
        el = el.replace('&', ' AND ')

    if re.search(r'\.{2,}', el):
        el = re.sub(r'\.{2,}', '', el)

    # замена " ," на ","
    if el.count(' ,') != 0:
        el = el.replace(' ,', ',')

    # замена "," с пропущенными пробелами (замена повторов пробелов ниже)
    if re.search(r'\d,[A-ZА-Я]|[A-ZА-Я],\d|[A-ZА-Я],[A-ZА-Я]', el):
        el = el.replace(',', ', ')

    # замена '\w\(' и '\)\w'
    if re.search(r'[^\s]\(|\)[^\s]', el):
        el = re.sub(r'\)', ') ', el)
        el = re.sub(r'\(', ' (', el)

    # замена '\( \w' и '\w \)'
    if re.search(r'\(\s\w|\w\s\)', el):
        el = re.sub(r'\(\s', '(', el)
        el = re.sub(r'\s\)', ')', el)

    # замена ' . '
    if re.search(r'\s\.\s|[A-Z)]{2,}\.[A-Z(]{2,}|\s\.[A-Z]{2,}', el) and \
            not re.search(r'WWW\.|HTTP', el):
        m = re.findall(r'\s\.\s|[A-Z)]{2,}\.[A-Z(]{2,}|\s\.[A-Z]{2,}', el)
        m_rev = [i.replace('.', ' ') for i in m]
        for num, item in enumerate(m):
            el = el.replace(item, m_rev[num])

    # замена нескольких "-" и "—" на "-"
    if re.search(r'—+|[\-]{2,}', el):
        el = re.sub(r'—+|[\-]{2,}', '-', el)

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
    if re.search(r'\(SEE.*\)?', el):
        el = re.sub(r'\(SEE.*\)?', '', el)

    # замена нескольких " " одним " "
    if re.search(r'\s{2,}', el):
        el = re.sub(r'\s{2,}', ' ', el)

    el = el.strip(' ,.')

    return el


def is_rus(s):
    if s != s:
        return False
    elif re.search(r'[А-Я]', s):
        return True
    else:
        return False


def angle_exists(string):

    if '<' in string or '>' in string:
        return True
    else:
        return False


curr_path = os.getcwd()
filenames = os.listdir(curr_path)
files = []
for file in filenames:
    if file.endswith('.xlsx') and file.count('~') == 0 and file.count('tmx'):
        files.append(file)
if len(files) != 1:
    print(f'Файлы: {files}')
    sys.exit('Файл для обработки не определен')

file = files[0]
df = pd.read_excel(os.path.join(curr_path, file), header=None, dtype=str)  # 3 cols: source|target|date

df['datetime'] = pd.to_datetime(df.iloc[:, 2], yearfirst=True)

df.iloc[:, 0] = df.iloc[:, 0].apply(replace_description)
df.iloc[:, 1] = df.iloc[:, 1].apply(replace_description)
print('===Before processing===')
print(df.info())
print('Is any NA in df:', df.isna().any(axis=None))  # check entire dataframe


df.sort_values('datetime', ascending=True, inplace=True)
df.drop_duplicates(df.columns[0], keep='last', inplace=True, ignore_index=True)
df['is_target_in_rus'] = df.iloc[:, 1].apply(is_rus)
df = df[df['is_target_in_rus'] == True]
print('===After processing===')
print(df.info())
df['is_angle_in_source'] = df.iloc[:, 0].apply(angle_exists)
df['is_angle_in_target'] = df.iloc[:, 1].apply(angle_exists)
df_angled = df[(df['is_angle_in_source'] == True) or (df['is_angle_in_target'] == True)]
if not df_angled.empty:
    print(df_angled)
    resp = ''
    while not re.search(r'[yn]', resp, flags=re.IGNORECASE):
        resp = input('В датафрейме обнаружены символы "<, >", формировать tmx файл (y/n)?')
    if re.search(r'n', resp, flags=re.IGNORECASE):
        sys.exit()

df.to_excel('new_tmx_after_processing.xlsx', header=False, index=False)
tmxs = []
for file in filenames:
    if file.endswith('.tmx') and file.count('~') == 0:
        tmxs.append(file)
if len(tmxs) != 1:
    print(f'Файлы: {tmxs}')
    sys.exit('Файл tmx для изменения не определен')

header = ''
with open(tmxs[0], 'rb') as file:
    while line := file.readline():
        pr_string = line.decode()
        header += pr_string
        if pr_string.count('</header>'):
            break

with open('new.tmx', 'wb') as tmx:
    tmx.write(header.encode('utf-8'))
    tmx.write('  <body>\r\n'.encode('utf-8'))
    for sour, tar, dat in zip(df.iloc[:, 0], df.iloc[:, 1], df.iloc[:, 2]):
        tmx.write('    <tu '.encode('utf-8'))
        tmx.write(f'creationdate="{dat}T090240Z" '.encode('utf-8'))
        tmx.write('creationid="MOS\\nekrasovkn" '.encode('utf-8'))
        tmx.write(f'changedate="{dat}T090240Z" '.encode('utf-8'))
        tmx.write('changeid="MOS\\nekrasovkn" '.encode('utf-8'))
        tmx.write(f'lastusagedate="{dat}T090240Z" '.encode('utf-8'))
        tmx.write('usagecount="1">\r\n'.encode('utf-8'))
        tmx.write('      <prop type="x-LastUsedBy">MOS\\nekrasovkn</prop>\r\n'.encode('utf-8'))
        tmx.write('      <prop type="x-Origin">TM</prop>\r\n'.encode('utf-8'))
        tmx.write('      <prop type="x-ConfirmationLevel">Translated</prop>\r\n'.encode('utf-8'))
        tmx.write('      <tuv xml:lang="en-US">\r\n'.encode('utf-8'))
        tmx.write(f'        <seg>{sour}</seg>\r\n'.encode('utf-8'))
        tmx.write('      </tuv>\r\n'.encode('utf-8'))
        tmx.write('      <tuv xml:lang="ru-RU">\r\n'.encode('utf-8'))
        tmx.write(f'        <seg>{tar}</seg>\r\n'.encode('utf-8'))
        tmx.write('      </tuv>\r\n'.encode('utf-8'))
        tmx.write('    </tu>\r\n'.encode('utf-8'))
    tmx.write('  </body>\r\n'.encode('utf-8'))
    tmx.write('</tmx>'.encode('utf-8'))





