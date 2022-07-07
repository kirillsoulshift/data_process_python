import pandas as pd
import re
import os
import sys
import time

"""
рекурсивная обработка файлов .xlsx в диреткории скрипта
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


start_time = time.time()
curr_path = os.getcwd()
folder_surfer(curr_path)

print("--- %s seconds ---" % (time.time() - start_time))