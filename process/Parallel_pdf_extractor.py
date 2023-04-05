import fitz
from PIL import Image
import io
import os
import re
import sys
from multiprocessing import Pool, cpu_count
import time


mytime = time.clock if str is bytes else time.perf_counter


def get_catnum_or_name(s):
    tokens = s.split('\n')
    catnum = ''
    for token in tokens:
        if re.search(r'[A-Z]?[\d-]{5,}[A-Z]?', token):
            m = re.search(r'[A-Z]?[\d-]{5,}[A-Z]?', token)
            catnum = m.group(0).replace(' ', '')
            break
    if catnum:
        return catnum
    else:
        return 'UNDEFINED'


def clean_str(el):
    if re.search(r'[\\/:*?«<>|+=;."\n]+', el):
        el = re.sub(r'[\\/:*?«<>|+=;."\n]+', '', el)
    return el


def open_pdf(vector : list):
    """

    :param vector: list, [segment number, number of CPU, filename, image_folder]
    :return:
    """
    # this is the segment number we have to process
    idx = vector[0]
    # number of CPUs
    cpu = vector[1]
    # open file
    my_pdf_file = fitz.open(vector[2])
    # get number of pages
    num_pages = my_pdf_file.page_count

    # pages per segment: make sure that cpu * seg_size >= num_pages!
    seg_size = int(num_pages / cpu + 1)
    seg_from = idx * seg_size  # our first page number
    seg_to = min(seg_from + seg_size, num_pages)  # last page number

    for page_number in range(seg_from, seg_to):

        # acess individual page
        page = my_pdf_file[page_number]
        text = page.get_text('text', sort=True)
        header = get_catnum_or_name(text)

        # accesses all images of the page
        images = page.get_images()

        # check if images are there
        if images:
            print(f"{len(images)} изображений на стр {page_number}/{num_pages}[+]", f'CAT_NUM = {header}')
        else:
            print(f"0 изображений на стр {page_number}[!]")

        # loop through all images present in the page
        for image_number, image in enumerate(images, start=1):
            # access image xref
            xref_value = image[0]

            # extract image information
            base_image = my_pdf_file.extract_image(xref_value)

            # access the image itself
            image_bytes = base_image["image"]

            # get image extension
            # ext = base_image["ext"]

            # load image
            image = Image.open(io.BytesIO(image_bytes))
            page_number_str = str(page_number).zfill(4)
            image_number_str = str(image_number).zfill(4)
            header_str = clean_str(header)
            # save image locally
            save_path = os.path.join(vector[3], f"pg{page_number_str}im{image_number_str}-{header_str}.png")
            image.save(open(save_path, "wb"), format='png')


if __name__ == "__main__":
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
    image_folder = f'{filename[:-4]}_изображения'
    image_path = os.path.join(curr_path, image_folder)
    os.mkdir(image_path)

    cpu = cpu_count()

    # make vectors of arguments for the processes
    vectors = [(i, cpu, filename, image_folder) for i in range(cpu)]
    print("Starting %i processes for '%s'." % (cpu, filename))
    t0 = mytime()  # start a timer
    pool = Pool()  # make pool of 'cpu_count()' processes
    pool.map(open_pdf, vectors, 1)  # start processes passing each a vector
    t1 = mytime()  # stop the timer
    print("Total time %g seconds" % round(t1 - t0, 2))
