import pandas as pd
import re
import os
import sys
import time


"""
ПРИМЕР ИСХОДНОГО .xlsx (разделяемый файл)
________________________________________________________________________________
1  |REF.  | PART NO.  |  QTY.  |  UNIT  |  DESCRIPTION  |         TARGET        |  # первая строка наименования столбцов 
___|______|___________|________|________|_______________|_______________________|
2  |      |           |        |        |               |                       |
___|______|___________|________|________|_______________|_______________________|
3  |577932|           |        |        |               |                       |  # между каталожным номером страницы
___|______|___________|________|________|_______________|_______________________|  # и названием страницы нет ячеек
4  |КОРПУС|           |        |        |               |                       |   
___|______|___________|________|________|_______________|_______________________|
5  |      |           |        |        |               |                       |  # может содержать пустые строки между 
___|______|___________|________|________|_______________|_______________________|  # названием страницы и спецификацией 
6  |REF.  | PART NO.  |  QTY.  |  UNIT  |  DESCRIPTION  |                       |  # может содержать лишние ненужные
___|______|___________|________|________|_______________|_______________________|  # наименовая столбцов
7  |001   | 57793200  |  1     |  EA    |HOUSE,MACHINERY|МАШИННОЕ ОТДЕЛЕНИЕ, С Г|
___|______|___________|________|________|_______________|_______________________|
8  |      |           |        |        |               |                       |
___|______|___________|________|________|_______________|_______________________|
9  |577305|           |        |        |               |                       |
___|______|___________|________|________|_______________|_______________________|
10 |УСТАНО|           |        |        |               |                       |
___|______|___________|________|________|_______________|_______________________|
11 |001   | 57775876  |  1     |  EA    |DECKING/RAILING|ПЛАТФОРМА С ПЕРИЛАМИ, З|
___|______|___________|________|________|_______________|_______________________|
12 |      | 57775884  |  1     |  EA    |- DECKING,ASSEM|- ПЛАТФОРМА В СБОРЕ    |
___|______|___________|________|________|_______________|_______________________|
                                    ...
                           <любое кол-во строк>              
                                    ...

обрабатывает .xlsx файл в папке в которой находится скрипт

ПЕРВАЯ СТРОКА ЭКСЕЛЯ = НАИМЕНОВАНИЯ КОЛОНОК
порядок столбцов:

 ================================================================================================================
           >>>'Item ID', 'Part Number', 'Quantity', 'Unit', 'Description', 'Translation', 'Comment'<<<  
 ================================================================================================================
 
между ячейками с партномером и названием
страницы не должно быть ячеек

если партномера нет, над названием страницы пустая ячейка

Разделяет .xlsx выгруженный из OCR
на множество .xlsx по названиям заголовков
и помещает их каждый в свою папку

скрипт формирует папку в своей директории
"""


def make_folder_num(num):

    if len(str(num)) == 1:
        num = ''.join(['000', str(num)])
    elif len(str(num)) == 2:
        num = ''.join(['00', str(num)])
    elif len(str(num)) == 3:
        num = ''.join(['0', str(num)])
    elif len(str(num)) == 4:
        num = str(num)
    return num


def is_unencodable(df):
    """
    проверка некодируемых символов в
    колонках заголовков и описания на
    английском
    """

    df.reset_index(drop=True, inplace=True)
    headers = df.iloc[:, 0]  # series

    for col in df.columns.to_list():
        if re.search(r'desc|назв', col, flags=re.IGNORECASE):
            result_col = col

    description = df[result_col]  # series
    ser = pd.concat([headers, description])
    ser.dropna(inplace=True)
    ser = set(''.join(ser.tolist()))
    control = []
    for i in ser:
        try:
            i.encode('cp1251')
        except UnicodeEncodeError:
            control.append(i)
    if control:
        print('\033[91m' + 'Некодируемые символы в файле (cp1251):' + '\033[0m')
        print(control)
        sys.exit()


def clean_str(el, xlsx_name: bool):
    """
    заменить недопустимые для наименования папок символы
    :param el: str
    :param xlsx_name: является ли строка именем эксель файла
    :return: str
    """
    # if re.search(r'[:*?]', el):
    #     el = re.sub(r'[:*?]', '', el)
    # if re.search(r'[\\/|]', el):
    #     el = re.sub(r'[\\/|]', '%', el)
    # if el.count('"') != 0:
    #     el = el.replace('"', '\'\'')
    # if re.search(r'\s{2,}', el):
    #     el = re.sub(r'\s{2,}', ' ', el)
    if xlsx_name:
        return el[0:3]  # далее название .xlsx файла не используется и может быть любым
    else:
        return el[0:10].strip() # наименование кратко записывается в название папки


def get_partnum(partnum):

    if partnum == '' or partnum != partnum:
        return 'NA'
    else:
        return partnum


def check_row(row):
    """
    проверка строки датафрейма на наличие
    значения в первом столбце, а во всех остальных NaN
    :param row: номер строки в списке
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


start_time = time.time()
curr_path = os.getcwd()

filenames = os.listdir(curr_path)
files = []
for file in filenames:
    if file.endswith('.xlsx') and file.count('~') == 0:
        files.append(file)
if len(files) != 1:
    print(f'Файлы: {files}')
    sys.exit('Файл для обработки не определен')

file = files[0]
dest_folder = f'{file[:-5]}'
dest_path = os.path.join(curr_path, dest_folder)
os.mkdir(dest_path)

df = pd.read_excel(os.path.join(curr_path, file), header=0, dtype=str)
is_unencodable(df)
columns_name = df.columns.to_list()

headers_n_refs = df[columns_name[0]].tolist()
check = df.values.tolist()  # для проверки содержания nan во всех колонках кроме первой

part_nums = []
headers = []
numbers = []
folder_count = 0

for row_num, header in enumerate(headers_n_refs):

    if check_row(check[row_num]) and re.search(r'[А-Я\s]{3,}', header, flags=re.IGNORECASE):
        headers.append(header)
        numbers.append(row_num)
        part_nums.append(headers_n_refs[row_num - 1])
        if len(numbers) > 1:
            folder_count += 1
            start = numbers[-2] + 1
            end = numbers[-1] - 1
            df_to_write = df.iloc[start:end].copy()

            folder_name = make_folder_num(folder_count) + \
                          " " + get_partnum(part_nums[-2]) + \
                          " " + clean_str(headers[-2], xlsx_name=False)
            name = clean_str(headers[-2], xlsx_name=True)
            os.mkdir(os.path.join(dest_path, folder_name))
            df_to_write.to_excel(os.path.join(dest_path, folder_name, name+'.xlsx'), index=False)
            with open(os.path.join(dest_path, folder_name, 'name.txt'), 'w') as f:
                f.write(headers[-2])  # полное наименование страницы записывается в текстовый файл в папке

folder_count += 1
start = numbers[-1] + 1
end = len(df)
df_to_write = df.iloc[start:end].copy()

folder_name = make_folder_num(folder_count) + \
              " " + get_partnum(part_nums[-1]) + \
              " " + clean_str(headers[-1], xlsx_name=False)
name = clean_str(headers[-1], xlsx_name=True)
os.mkdir(os.path.join(dest_path, folder_name))
df_to_write.to_excel(os.path.join(dest_path, folder_name, name+'.xlsx'), index=False)
with open(os.path.join(dest_path, folder_name, 'name.txt'), 'w') as f:
    f.write(headers[-1])  # полное наименование страницы записывается в текстовый файл в папке
print("--- %s seconds ---" % (time.time() - start_time))
