import pandas as pd
import os
# import sys
import re

"""
обрабатывает все подходящие файлы в своей директории
из партномеров удаляет :
-символы не относящиеся к партномеру
-скобки
-пробелы
отовсюду удаляет :
-элементы замены backslashreplace(\0\,\7\, итд.)

"""
def clean_partnumber(el):
    el = clean_backslashreplace(el)
    if el == el:

        if el.count('FIG.'):
            el = el.replace('FIG.', '')
        # замена нескольких "-" и "—" на "-"
        if re.search(r'[—-]{2,}', el):
            el = re.sub(r'[—-]{2,}', '-', el)
        # очищение небуквенных символов
        if re.search(r'[^\w.,-]+', el):
            el = re.sub(r'[^\w.,-]+', '', el)
        # перевод позиций в верхний регистр
        if re.search(r'[a-zа-я]', el):
            el = el.upper()

        if len(el) == 0:
            el = float('nan')
    return el


def clean_backslashreplace(el):
    if el == el:
        if re.search(r'\\.\\', el):
            el = re.sub(r'\\.\\', '', el)
        if len(el) == 0:
            el = float('nan')
    return el


def clean_cols(df):
    for col in df.columns:
        if col == 'Каталожный номер':
            df[col] = df[col].apply(clean_partnumber)
        elif col == 'Наименование в каталоге производителя' or\
             col == 'Наименование на русском языке':
            df[col] = df[col].apply(clean_backslashreplace)


DESC_COL = r'desc|описание'
REF_COL = r'id|index|ref|поз|п/п|nc'
PARTNUM_COL = r'part[_\s]?n|деталь[_\s]?№'
QTY_COL = r'qty|quant|кол-во|колич'
COMMENT_COL = r'comment|remark|коммент|примеча'
UNIT_COL = r'unit|ед\.\s?изм\.'
TRANSL_COL = r'transl|target|перевод'

curr_path = os.getcwd()
filenames = os.listdir(curr_path)
files = []
for file in filenames:
    if file.endswith('.xlsx') and file.count('~') == 0:
        files.append(file)

print(f'Файлы для обработки: {files}')

for file in files:
    print(f'Обработка {file}...')
    # проверка наличия заголовков
    df = pd.read_excel(file, header=0, dtype=str, nrows=1)
    test_check = ''.join(df.columns.tolist())
    if not re.search(r'(?=.*('+PARTNUM_COL+r'))(?=.*('+DESC_COL+r'))', test_check, flags=re.I):
        raise Exception(f'Проверьте наименования колонок в файле {file}')

    df = pd.read_excel(file, header=0, dtype=str)

    # определение колонок
    desc = ''
    part = ''
    target = ''
    for col in df.columns:
        if re.search(DESC_COL, col, flags=re.IGNORECASE):
            desc = col
        if re.search(PARTNUM_COL, col, flags=re.IGNORECASE):
            part = col
        if re.search(TRANSL_COL, col, flags=re.IGNORECASE):
            target = col

    tech_columns = ['Изготовитель', 'Каталожный номер', 'Наименование в каталоге производителя',
                    'Наименование на русском языке', 'Вид техники', 'Модель']
    dummy_col = [float('nan') for _ in range(len(df))]
    tech_df = pd.DataFrame({key: dummy_col for key in tech_columns})
    tech_df['Каталожный номер'] = df[part]
    tech_df['Наименование в каталоге производителя'] = df[desc]
    if target:
        tech_df['Наименование на русском языке'] = df[target]
    else:
        tech_df['Наименование на русском языке'] = df[desc]

    clean_cols(tech_df)
    tech_df.dropna(subset='Каталожный номер', inplace=True)
    tech_df.drop_duplicates(subset='Каталожный номер', inplace=True)
    tech_df.to_excel(file[:-5] + '_tc.xlsx', index=False)
