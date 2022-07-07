import pandas as pd
# import os
import re
# import numpy as np
# import sys
# import time
# import unicodedata

# import locale
# print(locale.getpreferredencoding())

df = pd.DataFrame({'0':[' - 123456','-2jk gd','- 3','8 - - ', '5-', '6'], '1':['566', '566', '566', '6666', '5', float('nan')]})

df['0'] = [i.upper() for i in df['0']]
print(df['0'])
# print("=============big.tmx==========")
# with open('big.tmx', 'rb') as file:
#
#     while line := file.readline():
#         pr_string = line.decode()
#
#         if pr_string.count('</tmx>'):
#             print(pr_string)
#             print(line)
#             break
# print("=============new.tmx==========")
# with open('new.tmx', 'rb') as file:
#
#     while line := file.readline():
#         pr_string = line.decode()
#         print(pr_string)
#         print(line)
#         if pr_string.count('</tu>'):
#             print(pr_string)
#             print(line)
#             break

# print(ord('é'))
# x = 'é'.encode('cp1251')
# print(x)
# print(0xD)
# print(0x1E)
# print(chr(0xD))
# print(ord('\r'))
# print(0xD)
# from unicodedata import category
# print(category('\n'))


