import pandas as pd
import os
import sys
import re


"""
обрабатывает .xlsx файл в директории скрипта
Обрабатывает первую колонку:
Первая буква заглавная в строке кроме английских
обозначений и брендов
"""


def capital_all_column(s):
    if s == s:
        s = s.lower()
        sentence = s.split()
        new_sentence = []
        for n, word in enumerate(sentence):
            if n == 0:
                word = word.capitalize()
                new_sentence.append(word)
            elif n != 0 and re.search(r'[a-z]{5,}', word):
                word = word.capitalize()
                new_sentence.append(word)
            elif n != 0 and re.search(r'[a-z]+?', word):
                word = word.upper()
                new_sentence.append(word)
            else:
                new_sentence.append(word)
        s = ' '.join(new_sentence).strip()
        return s
    else:
        return s


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

df[df.columns[0]] = df[df.columns[0]].apply(capital_all_column)

df.to_excel(file[:-5] + '_cap.xlsx', index=False, header=False)
