set PYTHONPATH=%USERPROFILE%\Desktop\Python38-32\Lib\site-packages;%PYTHONPATH%
cd %USERPROFILE%\Desktop\Python38-32\PageRedactor
%USERPROFILE%\Desktop\Python38-32\python.exe -i main.py

set PYTHONPATH=%USERPROFILE%\Desktop\Python38-32\Lib\site-packages;%PYTHONPATH%
cd %USERPROFILE%\Desktop\Python38-32\PageRedactor
%USERPROFILE%\Desktop\Python38-32\python.exe -i mass_replace.py

Модификация LinkOneCrush (WARNING: this repo do not contain folders data and components)

Обрабатывает книги LinkOne помещенные в папку to_process (находится в PageRedactor),
допускается добавление книг LinkOne внутри папок по производителю (так как они хранятся на Libraries\Mincom\Polyus)

Выгружает страницы в порядке следования по содержанию
со всем текстовым контентом по порядку и ссылками, содержащимися
на страницах. Результирующие файлы содержат выгрузку позиций из каталогов LinkOne с 
принадлежностью к системе. Каждому каталогу соответствует 
файл Excel (формируются в PageRedactor). Каждый файл содержит столбцы: Наименование, Номер по каталогу, 
последующие столбцы представляют собой уровни иерархии до корневой директории. 
(<Уровень 0> соответствует заголовку страницы, на которой встречается позиция в каталоге)

Для работы требуется файл словаря _global_dict.xlsx (находится в PageRedactor),

Если в обрабатываемом каталоге встречается позиция без перевода, выгрузка не происходит,
непереведенные позиции попадают в файл session_untranslated.xlsx (формируется в PageRedactor).
Для работы с непереведенными каталогами, требуется включить непереведенные позиции в _global_dict.xlsx
(с переводом во второй колонке), позиции из файла session_untranslated.xlsx (как есть)
в столбце перевода может быть любая строка, даже идентичная, но не пустая.

Пример:
SEALANT-GASKET			ПРОКЛАДКА ГЕРМЕТИЧНАЯ
  STUD (1/2-13X3.5-IN)		БОЛТ (1/2-13X3.5 ДЮЙМ.)
0000 834K S/N LWY00001-UP	0000 834K S/N LWY00001-UP
|0000000 D85EХ-15 S/N 20003-UP	|0000000 D85EХ-15 S/N 20003-UP
Housing, Blower			КОРПУС ВЕНТИЛЯТОРА ОБДУВА

В процессе работы скрипт выводит промежутчный отчет о состоянии в командную строку.
По окончании работы результат записывается в файл REFERENCE_PARSER_LOG.txt (находится в PageRedactor)
и в report.xlsx (формируется в PageRedactor, если количество обрабатываемых каталогов больше 4)

path = folder\pix_file_path

new_dir_path = folder\temp\folder

file_name = pix_file.txt

new_file_path = folder\temp\folder\pix_file.txt

file = self.read_file(path)

crushDecompressor
 crushDecompressor.decompress_pix()
  run(new_file_path, file, offset)

>>>bbi contains: 
		STRING :
[[b'CONTRACTOR_NAME', b'KOMATSU'], [b'SERIAL_NO', b'0036'], 
[b'VER', b'7'], [b'BUILD_MODE', b'PRODUCTION'], 
[b'EXPLICIT_PAGE_NUMBERS', b'TRUE'], [b'ALLOW_BOOK_PRINTING', b'FALSE'], 
[b'UNKNOWN_QUANTITY_STRING', b'AR'], [b'PRINT_TEMPLATE', b'SPLIT_FIXED'], 
[b'ALLOW_CONSOLIDATION', b'FALSE'], [b'SECURITY_LOCK', b'82263'], 
[b'SECURITY_LOCK', b'72259'], [b'SECURITY_LOCK', b'99527'], 
[b'SECURITY_LOCK', b'190605'], [b'BOOK_LOCK', b'190602'], 
[b'ABOUT_THIS_BOOK', b'Issued: Aug 28, 2019'], 
[b'ABOUT_THIS_BOOK', b'KOMATSU LTD. Tokyo'], 
[b'DEFAULT_SAVE_SEARCH_LEVEL', b'SAVE_ALL_IN_ALL_BOOKS'], 
[b'SAVE_ALL_IN_ALL_BOOKS', b'99527'], [b'ALLOW_SAVE_GRAPHIC', b'TRUE'], 
[b'SELECT_QTY_FITTED', b'TRUE'], [b'DEFAULT_SCREEN_LAYOUT', b'Standard'], 
[b'DEFAULT_PRINT_LAYOUT', b'Standard'], [b'MODEL', b'WA600-3'], 
[b'BUILD_CODEPAGE', b'utf-8']]
		BOOK :
[[b'wa600-6c', b'WA600-3 S/N 50001-UP', b'\xff\xff\xff\xff']]
		PUBLISHER :
[b'KOMATSU', b'Komatsu']
		LABELS :
['WA600-3 S/N 50001-UP', 'ENGINE RELATED PARTS', ...]
		ADDITIONAL_COLUMNS :
[b'SERIALNO']
		BUILD_CODEPAGE :
utf-8
		PAGES :
{'000000': {'page_num': '|000000', 'reference':
 '|000000', 'page_name': 'WA600-3 S/N 50001-UP'}, ... }

>>>Each pix contains: (на каждый столбец)
		DATA = {[POSITION, DATA], ...}

