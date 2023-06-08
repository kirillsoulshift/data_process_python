import pandas as pd
import re

'''с проверкой сетами
- проверка принадл символов
- пересмотреть проверку принадлежности символов
- М16Х40 БОЛТ М16Х40 если буква х русская, то в партномер не попадае + двойное попадание
числовых параметров в партномер
'''

PARTN_REGEX = r"[\s:!,№;%_\\|@#$^&<>()='\"?*`~]"

def is_rus_word(s):
    if re.search(r'^[а-яА-Я]+$', s):
        return True
    else:
        return False


triple_dict = pd.read_excel('toscript_TZJT_WK35_tripledict.xlsx', header=None, dtype=str)
triple_dict.columns = ['Source', 'Target', 'Desc']
triple_dict['list_of_sets'] = ['' for _ in range(len(triple_dict))]

for n, sn in enumerate(triple_dict['Source']):
    if not len(triple_dict[triple_dict['Source'] == sn]) == 1:
        str_set = triple_dict.loc[n, 'Desc'].split()
        uniq_desc = ''.join(list(set(str_set)))  # уникальный набор слов описания

        used_desc = ''
        if triple_dict['list_of_sets'].isin([uniq_desc]).any():                   # если набор слов уже встречался
            idx = triple_dict[triple_dict['list_of_sets'] == uniq_desc].index[0]  # поиск индекса с таким набором
            used_desc = triple_dict.loc[idx, 'Desc']                              # выбор уже использовавшегося описания
        else:
            triple_dict.at[n, 'list_of_sets'] = uniq_desc              # если не встречался, добавляется

        if used_desc:                                         #  выбор описания производился, строка описания выбирается
            str_set = triple_dict.loc[idx, 'Desc'].split()    #  из уже использованной
            str_features = [i for i in str_set if not is_rus_word(i)]
            str_features = ''.join(str_features)
            str_features = re.sub(r'[ЁёА-Яа-я]', '', str_features)
            target_sn = triple_dict.loc[n, 'Source'] + str_features
            triple_dict.at[n, 'Target'] = re.sub(PARTN_REGEX, '', target_sn)
        else:                                                     #  если не произведен выбор описания, создается новое
            str_features = [i for i in str_set if not is_rus_word(i)]   # из текущей строки
            str_features = ''.join(str_features)
            str_features = re.sub(r'[ЁёА-Яа-я]', '', str_features)

            target_sn = triple_dict.loc[n, 'Source'] + str_features
            triple_dict.at[n, 'Target'] = re.sub(PARTN_REGEX, '', target_sn)
    else:
        triple_dict.at[n, 'Target'] = re.sub(PARTN_REGEX, '', triple_dict.loc[n, 'Source'])

triple_dict.to_excel('partial_TZJT_WK35_tripledict.xlsx', header=False, index=False)





#
#
# desc_list = triple_dict['Desc'].tolist()
# desc_list_capital = [i.upper() for i in desc_list]
# triple_dict['DESC_CAPITAL'] = desc_list_capital
# WK35description = pd.read_excel('WK35description.xlsx', header=None, dtype=str)
# WK35description.columns = ['Desc']
# desc_list = WK35description['Desc'].tolist()
# set1 = set(desc_list_capital)
# set2 = set(desc_list)
# set3 = set1.intersection(set2)
# print(set3)

