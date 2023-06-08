import fitz
import os
import sys
import time
import pandas as pd
import re
import numpy as np
import matplotlib.pyplot as plt
from sklearn.cluster import KMeans, FeatureAgglomeration, BisectingKMeans, MiniBatchKMeans, SpectralClustering, \
    SpectralBiclustering, SpectralCoclustering


CENTROID_WORDS = ['Index', 'Code', 'Part', 'Description', 'Qty', 'Serial']
# INDEX_REGEX = r'^\d{1,3}$'
# CODE_REGEX = r'^\w{1,2}$'
# PARTNUMBER_REGEX = r'^[-\w]{11,}$'
# QTY_REGEX = r'^\d{1,3}$'
# SERIALNUM_REGEX = r'^[-\d]{7,}$'


FIRST_HEADER_WORD = 'Index'  # присутствие слова в строке на первом месте удалит строку
CATALOG_COLUMNS_NUMBER = 6   # количество колонок спецификации каталога / количество кластеров
DROP_LAST_ROW = True         # на страницах содержится строка с номером страницы и сер. номером
DROP_FIRST_ROW = True        # на страницах содержится строка с партномером и заголовком


def get_table_header(df):

    for n, row in zip(df.index.tolist(), df[df.columns[0]]):
        if row[0].count(f'{FIRST_HEADER_WORD}'):
            headers = df.loc[n, df.columns[0]]
            header_x1 = df.loc[n, df.columns[1]]
            header_x2 = df.loc[n, df.columns[2]]
            df.drop([n], inplace=True)
            break
    return headers, header_x1, header_x2, df


def make_page_spec(df):
    """
    датафрейм таблицы
    """
    for row in df[df.columns[0]]:
        for word in row:
            pass


def merge_x1(args):
    arg_list = []
    for arg in args:
        arg_list.append(arg)
    return arg_list


curr_path = os.getcwd()
filenames = os.listdir(curr_path)
pdf_list = []
for file in filenames:
    if file.endswith('.pdf') and file.count('~') == 0:
        pdf_list.append(file)
if len(pdf_list) != 1:
    print(pdf_list)
    sys.exit("Количество файлов pdf для обработки не = 1")

filename = pdf_list[0]
with fitz.open(filename) as my_pdf_file:
    # loop through every page
    pdf_len = len(my_pdf_file)
    for page_number in range(1, pdf_len + 1):

        # access individual page
        page = my_pdf_file[page_number - 1]
        text = page.get_text('words', sort=False)

        # create a dataframe of words with bboxes
        df = pd.DataFrame(text, columns = ['y1', 'x1', 'y2', 'x2', 'desc', 'd1', 'd2', 'd3'])

        # dataframe with lists of words with equal y parameter
        res = df.sort_values(['y1', 'x1'], ascending = [True, False])[['y1', 'desc']]\
            .groupby(['y1'])\
            .agg(merge_x1)         # word
        res['x1_list'] = df.sort_values(['y1', 'x1'], ascending = [True, False])[['y1', 'x1']]\
            .groupby(['y1'])\
            .agg(merge_x1)['x1']   # first x-coordinate of bbox
        res['x2_list'] = df.sort_values(['y1', 'x1'], ascending=[True, False])[['y1', 'x2']] \
            .groupby(['y1']) \
            .agg(merge_x1)['x2']   # second x-coordinate of bbox

        table_headers, x1, x2, res = get_table_header(res)
        print(table_headers)
        print(type(table_headers))
        print(x1)
        print(type(x1))
        print(x2)
        print(type(x2))
        x1x2 = np.array([x1, x2])
        print(x1x2)
        print(type(x1x2))
        x1x2 = np.transpose(x1x2)
        print(x1x2)
        print(type(x1x2))
        means = np.mean(x1x2, axis=1)
        print(means)
        headers_n_means = {i : j for i, j in zip(table_headers, means)}
        print(headers_n_means)

        index = res.index.tolist()
        if DROP_LAST_ROW and DROP_FIRST_ROW:
            table_df = res.drop([index[0], index[len(res)-1]]) #  table without header and footer
        elif DROP_LAST_ROW:
            table_df = res.drop([index[len(res) - 1]])  # table without footer
        elif DROP_FIRST_ROW:
            table_df = res.drop([index[0]])  # table without header
        else:
            table_df = res.copy()

        # for row, cords in zip(table_df[table_df.columns[0]], table_df[table_df.columns[1]]):
        #     for word, cord in zip(row, cords):
        #         print(word, cord)


        index = table_df.index.tolist()
        # obtain array of coordinates of all words on a page
        arr = []
        for y_cord, listx1, listx2 in zip(index, table_df['x1_list'], table_df['x2_list']):
            for x1, x2 in zip(listx1, listx2):
                arr.append([x1, x2, y_cord])

        # get centroids array
        centroids = []
        for word in CENTROID_WORDS:
            centroids.append(headers_n_means[word])
        centroids_arr = []
        for centroid_x in centroids:
            centroids_arr.append([centroid_x, 1])

        X = np.array(arr)
        CENT = np.array(centroids_arr)


        # start_time = time.time()
        KM = KMeans(n_clusters=CATALOG_COLUMNS_NUMBER, init=CENT).fit(X)
        BKM = BisectingKMeans(n_clusters=CATALOG_COLUMNS_NUMBER).fit(X)
        MBKM = MiniBatchKMeans(n_clusters=CATALOG_COLUMNS_NUMBER, init=CENT, n_init=1).fit(X)

        # print("---learning time: %s seconds ---" % (time.time() - start_time))

        fig, ax = plt.subplots(2, 2, figsize=(10, 5))
        ax[0, 0].scatter(X[:, 0], X[:, 1], c=KM.labels_)
        ax[0, 0].set_title("KM")

        ax[0, 1].scatter(X[:, 0], X[:, 1], c=BKM.labels_)
        ax[0, 1].set_title("BKM")
        ax[1, 0].scatter(X[:, 0], X[:, 1], c=MBKM.labels_)
        ax[1, 0].set_title("MBKM")
        plt.show()
        break

