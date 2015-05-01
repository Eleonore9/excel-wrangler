import sys, os, re

from os import listdir

from os.path import isfile, join
from openpyxl import load_workbook

import pandas as pd


def pre_process_xlsx(input_xls, output_xls):
    """ Solve the issue with empty cells
    at the beginning of the file """
    print 'Pre processing ', input_xls
    wb = load_workbook(input_xls, use_iterators=False, data_only=True, guess_types=False, keep_vba=False)
    wb_data = wb.active
    # Adding str to the first empty cells: 
    if wb_data['A1'].value is None:
        wb_data['A1'] = 'Date time column'
    if wb_data['B1'].value is None:
        wb_data['B1'] = 'Date column'
    if wb_data['C1'].value is None:
        wb_data['C1'] = 'Time column'
    # Turning the first row into datetime:
    for row in wb_data.rows:
        import datetime
        if not isinstance(row[0].value, datetime.datetime):
            if isinstance(row[1].value, datetime.datetime) and isinstance(row[2].value, datetime.time):
                from datetime import datetime
                dt = datetime.combine(row[1].value, row[2].value)
                #print 'dt: ', dt, type(dt)
                row[0].value = dt

    wb.save(output_xls)

# This step messed up the first datetime column
# but it's hard to delete it with openpyxl so
# i'll use pandas:
## Note: not used anymore
def clean_xlsx(processed_xls, output_csv):
    df = pd.read_excel(processed_xls)
    df = df.drop('Date time column', 1)
    print 'File cleaned!'
    df.to_csv(path_or_buf=output_csv, encoding='utf-8', index=False)


def convert_one_xlsx_to_csv(input_file, output_file):
    """Converts one xls(x) file to csv.
    Use from the command line:
    python convert_xlsx_to_csv.py 'my_file.xlsx' 
    'my_file.csv'"""
    print "Reading file " + input_file + "\n..."
    df = pd.read_excel(input_file)
    #print "List of columns:\n", list(df.columns.values)
    df.to_csv(path_or_buf=output_file, encoding="utf-8", index=False)
    print "File " + input_file + " converted to " + output_file


def convert_all_xlsx_to_csv(pre_process, path=os.getcwd()):
    """Converts all the xls(x) files in the current
    directory or given path to csv files."""
    path_files = [f for f in listdir(path) if isfile(join(path,f))] or []
    files = [f for f in path_files if '#' not in f]
    xls = filter(lambda x: '.xls'in x, files)
    csv = filter(lambda x: '.csv' in x, files)
    if len(xls) == 0:
        print 'There are no excel files to convert!'
    elif len(xls) == len(csv):
        for i in range(len(xls)):
            if re.sub(r'(.xls).', '', xls[i]) == re.sub(r'.csv', '', csv[i]):
                print 'Seems like all excel files have been converted already!'
    elif len(xls) > 0 and len(csv) < len(xls):
        to_convert = [f for f in xls if re.sub(r'(.xls).?', '.csv', f) not in csv]
        print "%d files to be converted!" % len(to_convert)
        fail, success = [], 0
        
        for f in to_convert:
            if pre_process:
                new_f = re.sub(r'(.xls).?', '_modif.xlsx', f)
                csv_f = re.sub(r'(.xls).?', '_modif.csv', f)
                try:
                    pre_process_xlsx(f, new_f)
                    #clean_xlsx(new_f, csv_f)
                    convert_one_xlsx_to_csv(new_f, csv_f)
                    success += 1
                except:
                    fail.append(new_f)
            else:
                try:
                    convert_one_xlsx_to_csv(f, re.sub(r'(.xls).?', '.csv', f))
                    success += 1
                except:
                    fail.append(f)
        print "%d/%d files were converted." %(success, len(to_convert))
        if len(fail) > 0:
            print len(fail), ' files couldn\'t be converted:\n- ' + ('\n- ').join(fail)
    else:
         print 'No idea what\'s going on here!'


if __name__ == "__main__":
    ## Give the file path from the command line.

    ## Example 1: 
    # Convert one file: convert_one_xlsx_to_csv(sys.argv[1], sys.argv[2])
    # In the terminal: python convert_xlsx_to_csv file.xlsx file.csv

    ## Example 2: 
    # Convert all xls(x) files: convert_all_xlsx_to_csv(False, path=sys.argv[1])
    # In the terminal: python convert_xlsx_to_csv.py ./

    ## Example 3:
    # Try to pre-process: pre_process_xlsx(sys.argv[1], sys.argv[2])
    # In the terminal: python convert_xlsx_to_csv file_in.xlsx file_out.xlsx
