

GENERAL_HELP = '''                                  
Обрабатывает .xlsx файл в папке в которой находится скрипт

+ очистить и выгрузить эксель 
+ выгрузить слова для перевода в эксель
+ перевести и выгрузить очищенный эксель
+ создать директорию каталога
+ выгрузка и помещение изображений из PDF
+ создать директорию ldf
'''

CLEAN_HELP = '''                                  
Для работы требуется файл с расширением .xlsx (выгрузка из OCR-проекта)

Обрабатывает .xlsx файл находящийся в пути "Исходный файл"
и формирует новый файл .xlsx в директории, указанной в пути "Целевая папка"

Для корректной работы алгоритма требуется:

    1) В первой строка в файле находятся наименования столбцов, для определения
    наименований столбцов используются паттерны регулярных выражений:
    Столбец описания: 'desc|описание'
    Столбец референса (первая): 'id|index|ref|поз|п/п|NC'
    Столбец партномеров: 'part[_\s]?n|деталь[_\s]?№'
    Столбец количества: 'qty|quant|кол-во|колич'
    Столбец комментариев: 'comment|remark|коммент|примеча'
    Столбец единиц измерения: 'unit|ед\.\s?изм\.'
    Столбец перевода: 'transl|target|перевод'
    
    2) Для начала работы алгоритма достаточными являются столбцы референса, описания
    и количества.

В результате работы происходит:
    - удаление избыточной пунктуации возникающей при
    OCR-распознавании текста в некоторых исходных файлах каталогов;
    - удаление лишних символов пунктуации в начале и конце строки;
    - удаление двойных пробелов, пробелов в конце и начале строки;
    - исправляется появление запятых, точек, тире итп. без пробелов;
    - в столбцах .xlsx удаляются наименования столбцов таблиц, 
    находящиеся внутри спецификаций;
    - удаляются пустые строки таблицы, (ни одна ячейка строки не заполнена); 
    - контент таблицы переводится в верхний регистр;
    - удаляются символы управления юникод
    - удаляются символы разметки html
    - удаляются китайские иероглифы

Описание функционала чекбоксов:

    Проверять наличие символов русской раскладки - по всем ячейкам таблицы
    производится поиск символов русской раскладки, функция используется для
    каталогов на английском языке с целью обнаружения ошибочно распознанных
    символов русской раскладки среди текстового контента на английском 
    (или другом) языке.
    Результат поиска (если есть) отображается в текстовом поле в нижней части
    окна программы.
    
    Двуязычное написание заголовков каталога - в некоторых каталогах встречаются
    наименования страниц и разделов на двух языках (MAINTENANCE / UNDERHALL),
    если подобные заголовки попали в обрабатываемый файл, включение данной функции
    оставляет вариант наименования встречающееся до символа "/" (т.е. на
    английском в примере выше).
    
    Удалять идентичные заголовки следующие подряд без партномеров - в исходных файлах 
    каталогов (.pdf) встречаются спецификации распространяющиеся на две или более страниц,
    при выгрузке из OCR-проекта в файл .xlsx, такая спецификация разделяется заголовком
    или заголовком с предшествующим партномером.
    Данная функция удаляет одинаковые заголовки и партномера встречающиеся подряд.
    Включите функцию, если обрабатываемый каталог имеет описанную выше структуру.
    Включите функцию, если используете шаблоны в OCR-редакторе.
    
    Выполнять слияние заголовков на соседних ячейках - в некоторых каталогах встречаются
    наименования страниц и разделов распространяющиеся на две и более строк, либо
    информация, требующая помещения в заголовок находится в разных местах на странице
    исходного каталога (.pdf). При выгрузке из OCR-проекта в файл .xlsx, такие заголовки
    помещаются на разные ячейки.
    Функция объединяет не идентичные заголовки стоящие на соседних ячейках.
    
    Проверять наличие идентичных заголовков следующих подряд - Функция позволяет отслеживать
    страницы каталога с одинаковыми заголовками встречающимися подряд.
    Результат поиска (если есть) отображается в текстовом поле в нижней части
    окна программы.
    
    Проверять количество символов в партномерах - отслеживает количество результирующих 
    (после очищения) символов в ячейках столбца партномеров.
    Результат отслеживания отображается в текстовом поле в нижней части окна программы.
'''

UPLOAD_HELP = '''
Для работы требуется файл с расширением .xlsx

Обрабатывает .xlsx файл находящийся в пути "Исходный файл"
и формирует новые файлы .xlsx в директории, указанной в пути "Целевая папка"

Функция производит выгрузку столбцов таблицы с удалением дубликатов и пустых
ячеек.
В первой строке в файле находятся наименования столбцов, для определения
наименований столбцов используются паттерны регулярных выражений:
    - Столбец описания: 'desc|описание' (обычно первый столбец, в нем же
    находятся заголовки)
    - Столбец референса (первая): 'id|index|ref|поз|п/п|NC'
    - Столбец комментариев: 'comment|remark|коммент|примеча' (если отмечен
    чекбокс "Добавить содержимое столбца комментариев к списку слов на перевод")

В результате в Целевой папке формируются отдельные файлы со списком слов на перевод из колонок
заголовков, описания и комментариев.

Если в ячейках в столбце описания присутствуют символы принадлежности к подразделам, файл
словаря описания будет представлен двумя столбцами, словами описания с предшествующими 
символами и словами описания без предшествующих символов (для загрузки в программу перевода).
Далее такой файл может быть использован в качестве файла сопоставления на стадии "Перевод".
'''

TRANSLATE_HELP = '''
Для работы требуется файлы с расширением .xlsx

Переводит .xlsx файл находящийся в пути "Исходный файл"
и формирует новый файл .xlsx в директории, указанной в пути "Целевая папка"
Исходный файл переводится с помощью словаря .xlsx находящимся в пути "Файл словаря"
Если в столбце описания исходного файла присутствуют символы принадлежности к подразделу, 
для перевода исходного файла требуется .xlsx файл сопоставления, который формируется на стадии
выгрузка слов на перевод. Файл указывается в пути "Файл сопоставления"

Функция переводит столбец описания путем создания столбца перевода в формируемом файле. Перевод
заголовков и комментариев производится на месте.
'''

DIRECTORY_HELP = '''
Для работы требуется файл с расширением .xlsx, выгруженный из OCR прошедший процедуры "Очистка" и
"Перевод"

Обрабатывает .xlsx файл находящийся в пути "Исходный файл"
и формирует директорию каталога в пути "Целевая папка"

Функция разделяет .xlsx выгруженный на множество .xlsx по наименованию страниц и помещает их каждый
в свою папку.

Для корректной работы функции требуется:

    1) партномер страницы занимает ячейку над наименованием страницы;

    2) между ячейками с партномером и наименованием страницы не должно быть ячеек;
    
    3) если у страницы отсутствует партномер, над названием страницы остается пустая ячейка;
    
    4) Опционально: поскольку для формирования папки раздела используется уникальный заголовок
    (с партномером или без), создание отдельных папок можно зарезервировать указанием
    названия страницы в столбце заголовков (первый столбец). Если у формируемого раздела отсутствует
    партномер см. п. 3. Для формирования папки табличная часть не требуется.
    
Результирующая папка будет содержать множество папок с названием типа "0236 3222332116 КОМПЛЕКТ Л",
где:
    1) 0236 - первые четыре цифры указывают номер будущей страницы .ldf;
    
    2) 3222332116 - следующая часть символов без пробелов - партномер раздела;
    
    3) КОМПЛЕКТ Л - остальная часть наименования папки представляет собой усеченное наименование
    страницы.

Внутри папки содержится файл name.txt с полным наименованием страницы внутри.
Если страница содержит спецификацию, внутри папки будет находится файл .xlsx с табличной частью
страницы из исходного файла.
'''

IMAGES_HELP = '''
Для работы требуется файл с расширением .pdf, исходный файл каталога

Обрабатывает .pdf файл находящийся в пути "Исходный файл"
и формирует две папки в пути "Целевая папка"

Функция извлекает изображения pdf файла, и помещает их в одну из формируемых папок.

Если не выделяются цельные изображения:
Adobe Fine Reader / OCR-редактор:
    1) разметить области каталожных номеров текстовым полем под нумерацией 1;
    2) разметить области заголовков текстовым полем под нумерацией 2;
    3) разметить области изображений полем для изображений под нумерацией 3;
    4) распознать документ;
    5) сохранить как pdf с возможностью поиска.
'''

MOVE_HELP = '''
Для работы требуются папки сформированные на стадиях "Директория" и "Изображения"

Функция помещает изображения по совпадению каталожного номера изображения (со страницы на которой 
находилось изображение) в одну из папок каталога с таким же каталожным номером в названии.

В пути "Не перемещенные изображения" указать одноименную папку сформированную на стадии
"Изображения" или иную. Данная папка должна содержать изображения с партномерами в названии.

В пути "Не востребованные изображения" указать одноименную папку сформированную на стадии
"Изображения" или иную. Данная папка не содержит изображений и служит контейнером изображений
не востребованных в книге LinkOne.

Граница группировки:
Перемещение изображений в папку не востребованные происходит по соответствию группам разрешений.
Все извлеченные в ходе работы алгоритма изображения могут быть сгруппированы по разрешению.
Наиболее многочисленные группы содержат изображения не представляющие ценности. Это могут быть 
полностью черные (белые) изображения, неполные изображения, изображение текстовой части 
в зависимости от разновидности и способа формирования исходного pdf файла. 
Для работы алгоритма требуется указать число изображений в группе с самым многочисленным 
содержанием изображений представляющими ценность. Те группы, число изображений в которых превышают 
указанное число, будут отмечены как невостребованные.

По завершении работы алгоритма папка "Не востребованные изображения" содержит все отсеянные
изображения (те изображения, что не попали в директорию каталога).

Папка "Не перемещенные изображения" может содержать ценные изображения, которые не переместились.
(у изображения отсутствует партномер, не было найдено совпадение по партномеру)

В папках директории каталога размещены изображения по совпадению партномера (на изображении и 
в названии папки)
'''

LDF_HELP = '''
Для работы требуется директория каталога сформированная на стадии "Директория" 
Данная стадия формирует папку с промежуточными файлами книги LinkOne (с расширением .ldf),
поэтому, для корректной работы, рекомендуется пройти стадии "Изображения" и "Перемещение".

Результирующая книга LinkOne обычно имеет вложенную структуру.
Формирование структуры книги LinkOne происходит на уровне создания промежуточных файлов .ldf. 
Для формирования структуры книги LinkOne, в промежуточных файлах .ldf должны быть размещены 
ссылки на соответствующие страницы каталога.

Например, для формирования многоступенчатой иерархии где заглавная страница каталога содержит
ссылки на главы, глава содержит ссылки на разделы, раздел содержит ссылки на подразделы, 
подраздел содержит ссылки на сборочную единицу итд.
Для автоматического размещения ссылок в файлах .ldf требуется создать аналогичную структуру 
папок и файлов в директории книги LinkOne файловой системы Windows.

Функция обрабатывает директорию каталога указанную в пути "Директория каталога"
В поле "BOOKCODE" указывается уникальный идентификатор книги LinkOne, назначающийся при
публикации каталога.

Поля "Тип механизма", "Модель", "Серийный номер" (опциональные) заполняются соответствующими
характеристиками обрабатываемой модели оборудования. При отсутствии данных в этих полях
сформированный файл BOOK.inf будет иметь пробелы (пустую строку) в соответствующих строках.

Чекбоксы "Столбец перевода" и "Столбец комментария" определяют будет ли создаваемая
книга иметь одноименные столбцы.

Алгоритм рекурсивно обрабатывает папки в Директории каталога, для каждой папки создается 
.ldf файл, в данный файл помещаются ссылки на изображения и папки, которые находились в этой
папке. 
Помещает созданные .ldf файлы с копиями изображений в новую папку '<BOOKCODE>_ldf',
которая создается в папке Директории каталога. Сформированная папка содержит файлы метаданных
для построения книги в Book Builder (это файлы .inf, .fmt, .sdf)
'''