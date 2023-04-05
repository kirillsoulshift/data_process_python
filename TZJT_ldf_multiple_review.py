import pandas as pd
import re
import os
import sys

"""
рекурсивная обработка файлов .ldf .xlsx в диреткории скрипта

- взять запретные символы из загрузки техкаталога
- проверять строку дескрипшн: в список 


"""


def append_comment(path):
    filenames = os.listdir(path)
    files = []
    for file in filenames:
        if file.endswith('.xlsx') and file.count('~') == 0:
            files.append(file)
    if len(files) > 1:
        print(f'Файлы: {files}')
        sys.exit('Файлов для обработки больше одного')

    elif len(files) == 1:
        file = files[0]
        file_path = os.path.join(path, file)
        df = pd.read_excel(file_path, header=0, dtype=str)
        df['Comment'] = [float('nan') for i in range(len(df))]
        rev_nums = []
        rem_qty = []
        rem_units = []
        for col in df.columns.tolist():

            if re.search(r'qty', col, flags=re.IGNORECASE):
                rev_list = df[col].tolist()

                for n, item in enumerate(rev_list):
                    if re.search(r'[.,]', item):
                        rev_nums.append(n)
                        rem_qty.append(item)
                        df.loc[n, col] = '1'

            if re.search(r'unit', col, flags=re.IGNORECASE) and rev_nums:
                for pos in rev_nums:
                    rem_units.append(df.loc[pos, col])
                    df.loc[pos, col] = 'EA'

            if re.search(r'comment', col, flags=re.IGNORECASE) and rev_nums:
                for n, pos in enumerate(rev_nums):
                    df.loc[pos, col] = rem_qty[n] + ' ' + rem_units[n]
        df.to_excel(file_path, index=False)


def reduce_afx(path):

    filenames = os.listdir(path)
    files = []
    for file in filenames:
        if file.endswith('.xlsx') and file.count('~') == 0:
            files.append(file)
    if len(files) > 1:
        print(f'Файлы: {files}')
        sys.exit('Файлов для обработки больше одного')

    elif len(files) == 1:
        file = files[0]
        file_path = os.path.join(path, file)
        df = pd.read_excel(file_path, header=0, dtype=str)
        processed_cols = 0
        for col in df.columns.tolist():

            if re.search(r'desc|targ', col, flags=re.IGNORECASE):
                processed_cols += 1
                process = df[col].tolist()
                processed = []
                for item in process:
                    if item == item:
                        if re.search(r'^[-_Dd\s]{3,}\w|^[-_Dd]{2,}\w', item):
                            processed.append(re.sub(r'^[-_Dd\s]+', '- ', item))
                        else:
                            processed.append(item)
                    else:
                        processed.append(item)
                df[col] = processed
        if processed_cols != 2:
            print(f'Файл: {file_path}')
            sys.exit('Столбцов для обработки больше одного')
        df.to_excel(file_path, index=False)


def folder_surfer(path):
    content = os.listdir(path)
    reduce_afx(path)
    append_comment(path)
    for item in content:
        if os.path.isdir(os.path.join(path, item)):
            folder_surfer(os.path.join(path, item))


# удаление ссылки на 2-е и последующие изображения в ldf по всей директории, (скрипт в род директории)
# curr_path = os.getcwd()
# filenames = os.listdir(curr_path)
# files = []
# for file in filenames:
#     if os.path.isdir(file):
#         files.append(file)
#
# if len(files) != 1:
#     print(f'Файлы: {files}')
#     raise Exception('не определена директория')
# print(f'Start review in {files[0]}')
# curr_folder = files[0]
# filenames = os.listdir(curr_folder)
# reviewed_count = 0
# for file in filenames:
#     if file.endswith('.ldf') and file.count('~') == 0:
#         with open(os.path.join(curr_folder, file), 'rb+') as f:
#             print(f'file {file} opened...')
#             for line in f.readlines():
#                 text = line.decode('cp1251')
#                 if text.count(',2,'):
#                     print(f'reviewing in {file}')
#                     text = text.replace(',2,', ',1,')
#                     f.seek(0)
#                     f.write(text.encode('cp1251'))
#                 break

# process ldf по всей директории, (скрипт в род директории)
# словарь triple_dict = датафрейм с тремя колонками Source | Target | Desc

triple_dict = pd.read_excel('_TZJT_WK35_tripledict.xlsx', header=None, dtype=str)
triple_dict.columns = ['Source', 'Target', 'Desc']
desc_list = triple_dict['Desc'].tolist()
desc_list_capital = [i.upper() for i in desc_list]
triple_dict['DESC_CAPITAL'] = desc_list_capital
# print(triple_dict)
# sys.exit()
curr_path = os.getcwd()
filenames = os.listdir(curr_path)
files = []
for file in filenames:
    if os.path.isdir(file) and not file.count('__'):
        files.append(file)

if len(files) != 1:
    print(f'Файлы: {files}')
    raise Exception('не определена директория')
print(f'Start review in {files[0]}')
curr_folder = files[0]
filenames = os.listdir(curr_folder)
decoderrors = []
for file in filenames:
    if file.endswith('.ldf') and file.count('~') == 0:
        with open(os.path.join(curr_folder, file), 'rb+') as f:
            print(f'file {file} opened...')
            new_content = []
            old_content = f.readlines()
            for line_num, line in enumerate(old_content):
                try:
                    text = line.decode('cp1251')
                except UnicodeDecodeError:
                    raise Exception(f'UnicodeDecodeError: file name of error: {file} in line number {line_num}')

                if text.count('ENTRY,'):
                    cells = text.split(',')
                    desc = cells[3].strip('"')
                    #   0   1     2                       3
                    # ENTRY,1,YJ56B1 - 1,Заземляющая щетка и её держатель,1,,,1,,,,,,
                    if desc in desc_list or desc in desc_list_capital:
                        was_replaced = False

                        old_partn = cells[2].strip('"').replace('.', r'\.')
                        print(f'match found {desc}')
                        mini_dict = triple_dict[triple_dict['DESC_CAPITAL'] == desc]
                        for idx in mini_dict.index:
                            if re.search(old_partn, mini_dict.loc[idx, 'Source'], flags=re.I):
                                newline = []
                                for n, i in enumerate(cells):
                                    if n != 2:
                                        newline.append(i)
                                    else:
                                        newline.append(mini_dict.loc[idx, 'Target'])
                                new_content.append(','.join(newline))
                                was_replaced = True
                                print(f'{old_partn} replaced by {mini_dict.loc[idx, "Target"]}')
                                break
                        if not was_replaced:
                            print(f'PARTNUMBER WASNT REPLACED AT {desc}')
                            new_content.append(text)

                    else:
                        new_content.append(text)
                else:
                    new_content.append(text)

            enc_new_content = [i.encode('cp1251') for i in new_content]

            if old_content != enc_new_content:
                if len(old_content) < len(enc_new_content):
                    new_content = new_content[:len(old_content)]
                    print(f'{file} got truncated by {len(enc_new_content)-len(old_content)} positions')
                elif len(old_content) > len(enc_new_content):
                    raise Exception(f'{len(old_content)-len(enc_new_content)} positions lost in {file}')
                f.truncate(0)
                f.seek(0)
                f.write(''.join(new_content).encode('cp1251'))
                print(f'{file} was replaced')
            else:
                print(f'{file} WAS NOT REPLACED')






