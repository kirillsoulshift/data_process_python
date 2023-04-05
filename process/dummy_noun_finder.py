import pandas as pd
import os
import sys
import re


# vowels = ['а', 'я', 'у', 'ю', 'о', 'е', 'ё', 'э', 'и', 'ы']
# consonants = ['б', 'в', 'г', 'д', 'ж', 'з', 'й', 'к', 'л', 'м', 'н', 'п', 'р', 'с', 'т', 'ф', 'х', 'ц', 'ч', 'ш', 'щ', 'ь']
# cased_noun_endings = ['ов', 'ев', 'ей', 'ам', 'ям', 'ому', 'ему', 'ой', 'ей', 'ого', 'его', 'ой', 'ей', 'ам', 'ям', 'ому',
#                       'ему', 'ой', 'ей', 'ей',
# 'ой', 'ий', 'ый', 'ую', 'юю', 'ое', 'ее', 'ого', 'его', 'ой', 'ей', 'ами', 'ями', 'им', 'ым', 'ой', 'ей', 'ах', 'ях',
#                       'ой', 'ей', 'ом', 'ем', 'ом', 'ем', 'ой', 'ей',
# 'ых', 'их', 'ыми', 'ими', 'ые', 'ие', 'ых', 'их', 'ым', 'им', 'ым', 'им', 'ой', 'ий', 'ый', 'ая', 'яя', 'ое', 'ее', 'ые', 'ие']
# pronoun_endings = ['ий', 'юю', 'ему', 'его', 'ем', 'ому', 'ого', 'ими', 'ыми', 'ое', 'яя', 'ую', 'им', 'ая', 'ее', 'ым',
#                    'ых', 'ей', 'ой', 'ые', 'ие', 'их', 'ом']
cased_noun_pronoun = ['ев', 'ем', 'ах', 'ов', 'им', 'яя', 'ями', 'ым', 'ой', 'ое', 'ому', 'ого', 'ам', 'ых',
                      'ий', 'ом', 'ыми', 'юю', 'его', 'их', 'ые', 'ими', 'ами', 'ее', 'ям', 'ый', 'ей', 'ях',
                      'ему', 'ие', 'ая', 'ую', 'ия']
nouns = ['УПЛОТНЕНИЕ', 'ОГРАЖДЕНИЯ', 'РАЗЪЕМ', 'РАЗЪЁМ', 'ОСНОВАНИЕ', 'ЗАЖИМ']



def is_english_or_partn(words : list):
    """рекурсивная Функция поиска строк на первом месте из
    англ. символов или цифр в списке из строк"""


    if not re.search(r'[А-Яа-я]', ''.join(words)):
        return words
    if not re.search(r'[А-Яа-я]', words[0]):
        words.append(words.pop(0))
        return is_english_or_partn(words)
    else:
        return words


def ending_check(s : str, check_list : list):
    """
    True : строка не может являться подлежащим, False : строка может являться подлежащим
    :param s: строка позиции
    :param check_list: список окончаний по которым проводится проверка
    :return: bool
    """
    if not re.search(r'[А-Яа-я]', s):  # слово не содержит рус символов
        return True

    if re.search(r'[A-Za-z\d]', s):  # слово слитно с англ символами или цифрами
        return True

    if re.search(r'[А-Яа-я]+[^А-Яа-яёЁ\-]+[А-Яа-я]+', s):  # слово из рус симв но в середине не буквенные символы
        return True

    if len(s) < 3:  # слово является союзом или предлогом
        return True

    for ending in check_list:                    # слово имеет окончание из списка
        expression = ending + r'$'               # окончание из списка на конце строки
        ending_ready = re.sub(r'[^\w]+', '', s)  # на проверку окончания подается строка без знаков препинания
        if re.search(expression, ending_ready, flags=re.I):
            return True

    return False


def tail_dub(s):
    words = s.split()
    if len(words) < 2:
        return s
    sub1 = words[-1]
    sub2 = words[-2]
    if re.search(r'\d', sub1) and re.search(r'\d', sub2):
        if re.sub(r'[^\d]+', '', sub1) == re.sub(r'[^\d]+', '', sub2):
            return ' '.join(words[:-1])
        else:
            return s
    else:
        if sub1 == sub2:
            return ' '.join(words[:-1])
        else:
            return s


# +учесть знаки кроме букв и цифр в начале и конце предложения
# +учесть когда все слова существительные
# +учесть когда одно слово
# +учесть когда первое слово на англиийском
# +учесть конда в начале встречаются цифры
# +алгоритм перестановки не должен принимать во внимание слова из анг символов и цифр знаки препинания итд
# +учесть вариант когда в списке только одно слово из рус букв
# +учесть вариант когда слово на рус склеено с англ или цифрами
# +учесть когда слово из рус симв но в середине не буквенные символы
# +учесть удаление дублирование в результирующем списке
# +РЕЗИСТИВНО - ЁМКОСТНОЙ МОДУЛЬ 34535 34535 => РЕЗИСТИВНО - ЁМКОСТНОЙ МОДУЛЬ ?????????????
# +ЭЛЕКТРИЧЕСКАЯ ЩЁТКА КАБЕЛЬНОГО БАРАБАНА =>	БАРАБАНА ЭЛЕКТРИЧЕСКАЯ ЩЁТКА КАБЕЛЬНОГО
# +УСТАНОВКИ РЫЧАГА ОТКРЫВАНИЯ ДНИЩА КОВША => УСТАНОВКИ РЫЧАГА ОТКРЫВАНИЯ ДНИЩА КОВША
# +Φ2X500 ЖЕЛЕЗНАЯ ПРОВОЛОКА? 2Х500  => Φ2X500 ЖЕЛЕЗНАЯ ПРОВОЛОКА? 2Х500
# СФЕРО - КОНИЧЕСКИЙ ДВУХРЯДОВЫЙ РОЛИКОВЫЙ ПОДШИПНИК => СФЕРО - КОНИЧЕСКИЙ ДВУХРЯДОВЫЙ РОЛИКОВЫЙ ПОДШИПНИК

def noun_finder(s):
    # handle NaN
    if s != s:
        return s

    words = s.split()

    # исходная строка пустая или состоит из единственного слова
    if len(words) <= 1:
        return s

    # когда в списке только одно слово из рус букв или вообще его нет
    aux_word_count = 0
    for word in words:
        if not re.search(r'[А-Яа-я]', word):
            aux_word_count += 1
    if len(words) - aux_word_count < 2:
        words = is_english_or_partn(words)
        return tail_dub(' '.join(words))

    # если на первом месте в строке слово на анг языке или партномер
    words = is_english_or_partn(words)

    # далее нужно будет учитывать первые слова до точки или разделителя
    rest_words = []
    active_statement = []
    delimeter_number = 0                       # номер слова на котором остановилось отделение

    if re.search(r'[^\w]', s):

        for n, word in enumerate(words):
            if not re.search(r'[^\w]$', word) or len(word) == 1:
                active_statement.append(word)
            else:
                active_statement.append(word)
                delimeter_number = n
                break
        if delimeter_number > 0:
            rest_words = words[delimeter_number+1:]    # оставшаяся часть выражения

    if active_statement:
        words = active_statement
        if len(words) <= 1:
            words.extend(rest_words)
            return tail_dub(' '.join(words))

    # поиск подлежащего и помещение его на первое место
    noun_count = 0
    noun_number = 0
    for n, word in enumerate(words):
        if not ending_check(word, cased_noun_pronoun):
            noun_count += 1
            noun_number = n

    if noun_count == 1:
        result = []
        result.append(words.pop(noun_number))
        result.extend(words)
        if rest_words:
            result.extend(rest_words)
        return tail_dub(' '.join(result))
    elif noun_count == 0:
        return 'noun_count eq 0'
    else:
        return 'noun_count more than 1'



curr_path = os.getcwd()
filenames = os.listdir(curr_path)
files = []
for file in filenames:
    if file.endswith('.xlsx') and file.count('~') == 0:
        files.append(file)
if len(files) != 1:
    print(f'Файлы: {files}')
    sys.exit(f'Файл для обработки в пути {curr_path} не определен')

file = files[0]
df = pd.read_excel(os.path.join(curr_path, file), header=0, dtype=str)
df['PROCESSED'] = df['DESCRIPTION'].apply(noun_finder)
df.to_excel(file[:-5] + '_NOUN_FINDED.xlsx', index=False)

# s = 'СБОРКА 2 РАСПРЕДЕЛИТЕЛЯ'
#
# s = noun_finder(s)
# print(s)

