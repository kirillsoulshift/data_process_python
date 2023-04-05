import pandas as pd
import os
import re

"""разделение строк одного столбца 
цифры | контент | 'S/N'

если колонка одна происходит разделение на три
если колонок 3 они объединяются"""

curr_path = os.getcwd()
filenames = os.listdir(curr_path)
files = []
for file in filenames:
    if file.endswith('.xlsx') and file.count('~') == 0:
        files.append(file)
if len(files) != 1:
    print(f'Файлы: {files}')
    raise Exception(f'Файл для обработки в пути {curr_path} не определен')

file = files[0]
df = pd.read_excel(os.path.join(curr_path, file), header=None, dtype=str)
source_row_count = len(df)
if len(df.columns) == 1:

    # если колонка одна происходит разделение на три
    df.columns = ['content']
    figures = []
    content = []
    sns = []
    for item in df['content']:
        # l = item.split()
        #
        # s = ' '.join(l[1:])
        # figure = l[0]
        # figures.append(figure)
        if re.search(r'^([a-zа-я]{0,2}[\d\s-]+)+', item, flags=re.I):
            figure = re.search(r'^([a-zа-я]{0,2}[\d\s-]+)+', item, flags=re.I).group(0)
            figures.append(figure)
            s = item[len(figure):]
        else:
            figure = ''
            figures.append(figure)
            s = item

        if re.search(r'S/N.+$', s):
            sn = re.search(r'S/N.+$', s).group(0)
            s = re.sub(r'S/N.+$', '', s).strip()
            content.append(s)
            sns.append(' '+sn)
        else:
            content.append(s)
            sns.append('')
    df['figures'] = figures
    df['content'] = content
    df['serial'] = sns
    df = df[['figures', 'content', 'serial']].copy()
    if len(df) != source_row_count:
        raise Exception('unequal rows count on explode')
    df.to_excel(file[:-5] + '_exploded.xlsx', index=False, header=False)

elif len(df.columns) == 3:

    # если колонок 3 они объединяются
    df.columns = ['figures', 'content', 'serial']
    df.fillna('', inplace=True)
    merged = []
    for n in range(len(df)):
        item = df.loc[n, 'figures'] + ' ' + df.loc[n, 'content'] + ' ' + df.loc[n, 'serial']
        merged.append(item.strip())

    df['merged'] = merged
    df = df[['merged']].copy()
    if len(df) != source_row_count:
        raise Exception('unequal rows count on merge')
    df.to_excel(file[:-5] + '_merged.xlsx', index=False, header=False)

else:
    print('количество колонок != 1|3')
