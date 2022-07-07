import pandas as pd
import os

"""
Обработка файла словаря .xlsx из двух колонок:
 - убрать из словаря непереведенные позиции и позиции таргет и сорс которых совпадает;
 - копировать их в отдельный файл, это будет файл непереведенных английских позиций;
 - удалить дубликаты по колонке source.
"""

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
df = pd.read_excel(os.path.join(curr_path, file), header=None, dtype=str)

nan_col = df[df[df.columns[1]] != df[df.columns[1]]]
equal = df[df[df.columns[0]] == df[df.columns[1]]]
all = equal.append(nan_col)

df = df[(df[df.columns[0]] != df[df.columns[1]])]
df.dropna(inplace=True)
df.drop_duplicates(df.columns[0], keep='last', inplace=True)
all.to_excel(file[:-5] + '_untranslated_eng_segments.xlsx', header=False, index=False)
df.to_excel(file[:-5] + '_valid_no_dub.xlsx', header=False, index=False)
