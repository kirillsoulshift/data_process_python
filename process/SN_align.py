import pandas as pd
import os
import re

"""проверка словарей на соответствие выражения после 'S/N'
проверка на отсутствие в переводе
проверка на идентичность в оригинале и переводе"""


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
df.columns = ['Source', 'Target']


missing_sn_idxs = []
inconsistent_sn_idxs = []
for n in range(len(df)):
    if df.loc[n, 'Source'].count('S/N') and not df.loc[n, 'Target'].count('S/N'):
        missing_sn_idxs.append(n)
        sn = ''.join(df.loc[n, 'Source'].rpartition('S/N')[1:])
        df.at[n, 'Target'] = df.loc[n, 'Target'] + ' ' + sn

    elif df.loc[n, 'Source'].count('S/N') and df.loc[n, 'Target'].count('S/N'):

        source_sn = ''.join(df.loc[n, 'Source'].rpartition('S/N')[1:])
        target_sn = ''.join(df.loc[n, 'Target'].rpartition('S/N')[1:])
        if source_sn != target_sn:
            if not re.search(r'[а-я]', target_sn, flags=re.I):

                inconsistent_sn_idxs.append(n)
                df.at[n, 'Target'] = ''.join(df.loc[n, 'Target'].rpartition('S/N')[0]) + source_sn

missing_sn_idxs = [i+1 for i in missing_sn_idxs] # for excel
inconsistent_sn_idxs = [i+1 for i in inconsistent_sn_idxs] # for excel
df.to_excel(file[:-5] + '_SN_aligned.xlsx', index=False, header=False)
print('S/N was missing at rows:', missing_sn_idxs)
print('S/N was inconsistent at rows:', inconsistent_sn_idxs)