# encoding=utf-8
'''
DESC:
The codes read the event files of operations from the DCS of powerplant, and
make the statistic investigation about the occurences of operations, then
packages the python source into a standalone Windows executables by means of
the command tool PyInstaller. eg. 'pyinstaller -F op-stat.py '

AUTHOR: Zhiwei YAN (jerodyan@163.com)
DATE:   2020-03-13
VERSION：0.1
'''

import sys, os, glob
from time import sleep
import pandas as pd

def hr(msg):
    print(80 * '-')
    print(msg)
    sleep(0.5)

def print_usage():
    print('\"'*80)
    print("The program deals with the MS EXCEL files in the dir (./input/) as inputs, ")
    print("writes an operation summary file 'output.csv' into the dir (./output/) as an output. ")
    print(' ')
    print('Copyright 2020, Rui-Dian-Sci-Tech. Contact: jerodyan@163.com.')
    print('\"'*80)

def check_folder(dir_name):
    if not os.path.exists(dir_name):
        os.makedirs(dir_name)
    return dir_name

def press_and_continue():
    a= input("Press ENTER key to continue.")

def press_and_exit(code):
    a = input("Press ENTER key to exit.")
    sys.exit(code)

# Some Constants
input_dir = './input/'
input_files_extension_name = '*.xls'

output_dir = './output/'
output_file = 'output.csv'

default_sheet_name = 'Sheet'
skip_excel_rows_num = 9
column_name = 'Message'
removed_chars = '1234567890.-'
event_column_name = 'Column-01-Operation-Event-Name'
counts_column_name = 'Column-02-Event-Counts'

if __name__ == "__main__":

    print_usage()

    # check the input dir and the output dir.
    check_folder(input_dir)
    check_folder(output_dir)
    
    # check and scan the input files
    hr('START: ')
    dcs_files = sorted(glob.glob(input_dir+input_files_extension_name))
    if  len(dcs_files) == 0:
        print('ERROR: There are NOT %s files in the directory ./input/.' %input_files_extension_name)
        press_and_continue()
    assert len(dcs_files)!=0, 'There are NOT files in the directory ./input/.'

    total_sheets = pd.Series([])
    for (idx, sample_file) in enumerate(dcs_files):
        print('processing file [%d] : %s' %(idx+1, sample_file), end=' ')
        try:
            xls = pd.read_excel(sample_file, skiprows=skip_excel_rows_num, sheet_name=default_sheet_name)
        except:
            print(', [Failed]')
            print("  - ERROR: Opening the file: %s, and Skip the file." %(sample_file))
            continue

        # remove white spaces or the numbers on the tail of event records.
        one_sheet = xls[column_name].str.rstrip(removed_chars).str.strip()
        one_sheet = one_sheet.str.replace('\s+', r' ', regex=True)

        # combine the records in all files.
        total_sheets = total_sheets.append(one_sheet)
        print(', [OK]')
        sleep(0.3)

    if total_sheets.empty:
        hr('WARNING: There are not any records to be analyzed.')
        print('Sorry, It failed. You SHOULD check your .xls files again. ')
        press_and_exit(1)

    res = total_sheets.value_counts(sort=True, ascending=False)
    # add the labels into the results.
    df = res.rename_axis(event_column_name).to_frame(counts_column_name)
    hr('Top 5 records are shown:')
    print(df.head())

    # save the result into the harddisk.
    df.to_csv(output_dir+output_file)
    hr('Save the result: ' + output_dir+output_file)

    hr("All Done, END.")
    press_and_exit(0)
