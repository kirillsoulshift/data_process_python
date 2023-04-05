import os
import sys
import struct



def get_size(file, pos):
    hex_size = file[pos[0]:pos[1]]
    try:
        dec_size = struct.unpack('<H', hex_size)[0]
    except:
        try:
            dec_size = struct.unpack('<I', hex_size)[0]
        except:
            dec_size = struct.unpack('<B', hex_size)[0]

    return dec_size


curr_path = os.getcwd()
filenames = os.listdir(curr_path)
files = []
for file in filenames:
    if (file.endswith('.pix') or file.endswith('.PIX')) and file.count('~') == 0:
        files.append(file)
if len(files) != 1:
    print(f'Файлы: {files}')
    sys.exit(f'Файл для обработки в пути {curr_path} не определен')

file = files[0]
with open(file, "rb") as f:

    file = f.read(1)
    pos =
    dec_size = get_size(file, pos)
