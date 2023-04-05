import pandas as pd
import re
# import os
import sys
import time
from tqdm import tqdm


def extract_parentheses(el):
    if el == el:
        el = re.sub(r'[^()]', '', el)
        return el
    else:
        return el


def extract_digits(el):
    if el == el:
        el = re.sub(r'[^\d]', '', el)
        return el
    else:
        return el


def extract_letters(el):
    if el == el:
        el = re.sub(r'[^a-zA-Z]', '', el)
        return el
    else:
        return el


def parentheses_check(df, col1 : str, col2 : str):
    """
    =====возвращает лист======
    :return: list
    """
    ser1 = df[col1].apply(extract_parentheses)
    ser2 = df[col2].apply(extract_parentheses)
    print('Parentheses check started...')
    res = []
    for i in tqdm(range(len(df))):
        if ser1[i] != ser2[i]:
            res.append(True)
        else:
            res.append(False)
    return res


def digits_check(df, col1 : str, col2 : str, check_col1 : str):
    """
    =====возвращает лист======
    :return: list
    """
    ser1 = df[col1].apply(extract_digits)
    ser2 = df[col2].apply(extract_digits)
    print('Digits check started...')
    res = []
    for i in tqdm(range(len(df))):
        if ser1[i] != ser2[i] and not df[check_col1][i]:
            res.append(True)
        else:
            res.append(False)
    return res


def letters_check(df, col1 : str, col2 : str, check_col1 : str, check_col2 : str):
    """
    =====возвращает лист======
    :return: list
    """
    ser1 = df[col1].apply(extract_letters)
    ser2 = df[col2].apply(extract_letters)
    print('Letters check started...')
    res = []
    for i in tqdm(range(len(df))):
        if ser1[i] != ser2[i] and not df[check_col1][i] and not df[check_col2][i]:
            res.append(True)
        else:
            res.append(False)
    return res


def status_insert(df, paran : str, digits : str, letters : str,):
    """
    =====возвращает лист======
    :return: list
    """
    print('Status insertion started...')
    res = []
    for i in tqdm(range(len(df))):
        if df[paran][i]:
            res.append('удалены параметры')
        elif df[digits][i]:
            res.append('параметры не соответствуют')
        elif df[letters][i]:
            res.append('название позиций не соответствует')
        else:
            res.append('соответствие')
    return res


start_time = time.time()
print('Started...')
dicted = pd.read_excel('LO_TK_dict.xlsx', header=0, dtype=str)
print("Got df in %s seconds" % (time.time() - start_time))


dicted['paranthesis'] = parentheses_check(dicted, 'DESCRIPTION', 'Description_TK')
dicted['digits'] = digits_check(dicted, 'DESCRIPTION', 'Description_TK', 'paranthesis')
dicted['letters'] = letters_check(dicted, 'DESCRIPTION', 'Description_TK', 'paranthesis', 'digits')
dicted['STATUS'] = status_insert(dicted, 'paranthesis', 'digits', 'letters')
dicted.drop(columns=['paranthesis', 'digits', 'letters'], inplace=True)

print("Excel file wrighting started...")
start_time = time.time()
dicted.to_excel('LO_TK_dict_status.xlsx', index=False)
print("Got excel file in %s seconds" % (time.time() - start_time))
