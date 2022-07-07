from PIL import Image
import os
import pandas as pd

"""
проверка разрешения картинок .png в папке
в которой находится скрипт

вывод таблицы разрешений и количества 
изображений с данным разрешением
"""

image_path = os.getcwd()
content = os.listdir(image_path)
dimensions = []
for file in content:
    if file.endswith('.png'):
        path = os.path.join(image_path, file)
        with Image.open(path) as img:
            dimensions.append(img.size)
commons = pd.DataFrame({'dims' : dimensions})

print('Разрешение | Число картинок\n', commons['dims'].value_counts())
