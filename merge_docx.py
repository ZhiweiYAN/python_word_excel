# encoding=utf-8
'''
DESC:
There are several MS word files(.docx). We combine them into
a single one keeping their original styles.

AUTHOR: Zhiwei YAN (jerodyan@163.com)
DATE:   2020-05-18
VERSION: 0.1
'''

import sys, os, glob, re
from time import sleep
import win32com.client as win32
from win32com.client import constants


from docxcompose.composer import Composer
from docx import Document as Document_compose


def print_usage():
    print('\"' * 80)
    print("The program deals with the MS WORD files in the dir (./input/) as inputs, ")
    print("writes an operation summary file 'output.docx' into the dir (./output/) as an output. ")
    print(' ')
    print('Copyright 2020, Zhiwei YAN. Contact: jerodyan@163.com.')
    print('\"' * 80)

def hr(msg):
    print(80 * '-')
    print(msg)
    sleep(0.5)

def check_folder(dir_name):
    if not os.path.exists(dir_name):
        os.makedirs(dir_name)
    return dir_name

def press_and_continue():
    a = input("Press ENTER key to continue.")

def press_and_exit(code):
    a = input("Press ENTER key to exit.")
    sys.exit(code)

def save_as_docx(path):
    # Opening MS Word application privided that there is an installed version of Word.
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(path)
    doc.Activate ()

    # Rename path with .docx
    new_file_abs = os.path.abspath(path)
    new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
    new_file_abs = re.sub(r'(.*)(\d{2})_(\d{2})_(\d{4})(.*)', r'\g<1>\g<4>\g<2>\g<3>\g<5>', new_file_abs )

    # Add a pagebreak on top of file
    pos_ins = word.ActiveDocument.Range(0,0)
    pos_ins.InsertBreak()

    # Add one blank page if the doc has odd pages.
    wdStatisticPages = 2
    doc_pages_orig = word.ActiveDocument.ComputeStatistics(wdStatisticPages)
    if (0==doc_pages_orig%2):
        word.ActiveDocument.Sections.Add()

    doc_pages = word.ActiveDocument.ComputeStatistics(wdStatisticPages)
    print('FILE: %s, PAGES: %d -> %d' %(new_file_abs, doc_pages_orig, doc_pages))

    # Save and Close
    word.ActiveDocument.SaveAs(
        new_file_abs, FileFormat=constants.wdFormatXMLDocument
    )
    doc.Close(False)
    # sys.exit(-1)

def combine_all_docx(filename_master,files_list, output):
    print('Merging: %s' %filename_master)
    number_of_sections=len(files_list)
    master = Document_compose(filename_master)
    composer = Composer(master)
    for i in range(0, number_of_sections):
        doc_temp = Document_compose(files_list[i])
        composer.append(doc_temp)
        print('Merging: %s' %(files_list[i]))
    composer.save(output)
    hr('OUTPUT:')
    print('Output file: %s' %output)

# Some Constants
input_dir = 'E:\\temp\\input\\'
input_files_extension_name = '*.doc'

output_dir = 'E:\\temp\\output\\'
output_file = 'output.docx'

if __name__ == "__main__":

    print_usage()

    # check the input dir and the output dir.
    check_folder(input_dir)
    check_folder(output_dir)

    # check and scan the input files
    hr('START: ')
    doc_files = sorted(glob.glob(input_dir + input_files_extension_name))
    if len(doc_files) == 0:
        print('ERROR: There are NOT %s files in the directory ./input/.' % input_files_extension_name)
        press_and_exit(1)
    assert len(doc_files) != 0, 'There are NOT files in the director ./input/.'
    print(doc_files)


    # convert file format from doc to docx.
    hr('SAVE AS DOCX:')
    for doc_file in doc_files:
        save_as_docx(doc_file)

    hr('MERGE DOCXs:')
    docx_files = sorted(glob.glob(input_dir + '*.docx'))
    combine_all_docx(docx_files[0], docx_files[1:], output_dir+output_file)

    hr('END.')

