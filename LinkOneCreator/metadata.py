import os


def make_meta(dest_path: str, book_code: str, eq_type: str, model: str, serial: str, translate: bool, comment: bool):
    """Создание файлов метаданных для книг LinkOne"""
    book_code = str(book_code).strip()
    eq_type = str(eq_type).strip()
    model = str(model).strip()
    serial = str(serial).strip()
    cat_name = model + ' S/N ' + serial
    # ====== BOOK.INF ======
    bookinf_path = os.path.join(dest_path, 'BOOK.INF')
    with open(bookinf_path, 'w') as inf:

        inf.write('PUBLISHER,Polyus,Polyus inc.\n')
        inf.write(f'BOOK,{book_code},{cat_name},{eq_type}\n')
        inf.write('STRING,EXPLICIT_PAGE_NUMBERS,TRUE\n')
        inf.write('STRING,CLOSEST_PRECEEDING_PARENT,FALSE\n')
        inf.write('STRING,SHOW_OPTION_NAMES,TRUE\n')
        inf.write('STRING,HIDE_REFERENCES,FALSE\n')
        inf.write('STRING,ALLOW_CONSOLIDATION,TRUE\n')
        inf.write('STRING,INVALIDATE_ORPHANS,FALSE\n')
        inf.write('STRING,SELECT_QTY_FITTED,FALSE\n')
        inf.write('STRING,ORDINAL_PAGE_NUMBERS,FALSE\n')
        inf.write('STRING,ALLOW_BOOK_PRINTING,TRUE\n')
        inf.write('STRING,ALLOW_SAVE_GRAPHIC,TRUE\n')
        inf.write(f'STRING,ABOUT_THIS_BOOK,Model {model}\n')
        inf.write(f'STRING,ABOUT_THIS_BOOK,S/N {serial}\n')
        inf.write('SEARCH_DEFINITION,Rus,Russian Translation,.\search.sdf\n')
        inf.write('STRING,DEFAULT_SEARCH_SET,Rus\n')
        if translate:
            inf.write('SEARCHABLE_IF,Part Number,Description,Translate\n')
        else:
            inf.write('SEARCHABLE_IF,Part Number,Description\n')
        inf.write('LAYOUT_DEFINITION,ruslayout,Layout с переводом,.\\ruslayout.fmt\n')
        inf.write('STRING,DEFAULT_SCREEN_LAYOUT,ruslayout\n')
        inf.write('STRING,DEFAULT_PRINT_LAYOUT,ruslayout\n')
        inf.write(f'BOOK_DESCRIPTION,RU,{model} S/N {serial}\n')
        inf.write(f'BOOK_DESCRIPTION,EN,{model} S/N {serial}\n')
        if translate:
            inf.write('FIELD_DEFINITION,Translate\n')
        if comment:
            inf.write('FIELD_DEFINITION,Comment\n')

    # ====== RUSLAYOUT.FMT ======
    ruslayout = os.path.join(dest_path, 'ruslayout.fmt')
    with open(ruslayout, 'w') as lay:

        lay.write('<!-- FAST KEYBOARD INPUT FOR HOTPOINT -->\n\
<LAYOUT DESCRIPTION="Example List Layout Definition">\n\
<FORMAT NOROWCOLLAPSE NOCOLUMNCOLLAPSE NOHIDEREPEATITEMIDS>\n\
  <HEADINGS>\n\
    <ROW>\n\
      <CELL COLSPAN=2>  <!-- over the two icon columns -->\n\
      <CELL> "Элемент"\n\
      <CELL> "Номер по каталогу"\n\
      <CELL> "Описание"\n')
        if translate:
            lay.write('      <CELL> "Наименование"\n')
        if comment:
            lay.write('      <CELL> "Комментарий"\n')
        lay.write('      <CELL> "Кол-во"\n\
      <CELL> "Ед. изм."\n\
    </ROW>\n\
  </HEADINGS>\n\
  <ROW>\n\
    <CELL COLGAP=0> <FIELD> Link Icon\n\
    <CELL COLGAP=0> <FIELD> Selection Icon\n\
    <CELL COLUMNCOLLAPSE> <FIELD> Item ID\n\
    <CELL COLUMNCOLLAPSE> <FIELD> Part Number\n\
    <CELL COLUMNCOLLAPSE> <FIELD> Description\n')
        if translate:
            lay.write('    <CELL COLUMNCOLLAPSE> <FIELD> Translate\n')
        if comment:
            lay.write('    <CELL COLUMNCOLLAPSE> <FIELD> Comment\n')
        lay.write('    <CELL COLUMNCOLLAPSE> <FIELD> Quantity\n\
    <CELL COLUMNCOLLAPSE> <FIELD> Units\n\
  </ROW>\n\
</FORMAT>\n\
</LAYOUT>')

    # ====== SEARCH.SDF ======
    search = os.path.join(dest_path, 'search.sdf')
    with open(search, 'w') as sdf:

        sdf.write(';This is an example of a Search Definition File.  It \n\
;defines a Search Set that matches the default Searches \n\
;provided by LinkOne.\n\
\n\
SET\n\
;This is an example of the Titles search. It\n\
;defines a Search that matches the default \n\
;Titles search provides by LinkOne.\n\
\n\
SEARCH,TITLES,Заголовки,,,Поиск по заголовкам...\n\
MENU,&Поиск по заголовкам...\n\
KEY,Title,LIST\n\
SEARCH_FIELD,Title\n\
\n\
SEARCH,REFERENCES,References,,,Matching References\n\
KEY,Reference,LIST\n\
SEARCH_FIELD,Page Reference\n\
\n\
;This is an example of the Page Numbers search. \n\
;It defines a Search that matchhes the default \n\
;Page Numbers search provides by LinkOne.\n\
;Note: This search is NOT on the menu.\n\
\n\
SEARCH,PAGE_NUMBERS,Номера страниц,,,Поиск по номерам страниц\n\
KEY,Page Number,LIST\n\
SEARCH_FIELD,Page Number\n\
\n\
SEARCH,DESCRIPTIONS,Английские описания,,,Поиск по английским описаниям...\n\
MENU,Поиск по английским описаниям...\n\
KEY,Description\n\
SEARCH_FIELD,Description\n\
\n\
;This is an example of the Part Numbers search. \n\
;It defines a Search that matchhes the default \n\
;Part Numbers search provides by LinkOne.\n\
\n\
SEARCH,PART_NUMBERS,Номера в каталоге,,,Поиск по каталожным номерам...\n\
DISPLAY_FIELD,Description\n\
MENU,&Поиск по номерам в каталоге...\n\
KEY,Part Number\n\
SEARCH_FIELD,Part Number\n\
\n')

        if translate:
            sdf.write('SEARCH,Translate,Поиск по русским описаниям,,,,Поиск по переводу\n\
MENU,&Поиск по русским описаниям\n\
KEY,Translate\n\
SEARCH_FIELD,Translate, Part Number, Description\n\
\n')

        sdf.write(';This is an example of the Part Numbers/Descriptions search. \n\
;It defines a Search that matchhes the default \n\
;Part Numbers/Descriptions search provides by LinkOne.\n\
;Note: This search is NOT on the menu.\n\
\n\
SEARCH,PNUM_DESCRIPTIONS,Поиск по каталожным номерам и описанию,,\
PAIRED,Поиск по каталожным номерам и описанию\n\
DISPLAY_FIELD,Quantity\n\
KEY,Part Number\n\
SEARCH_FIELD,Part Number\n\
KEY,Description\n\
SEARCH_FIELD,Description\n\
\n\
;This is an example of the Descriptions search. \n\
;It defines a Search that matchhes the default \n\
;Descriptions search provides by LinkOne.')
