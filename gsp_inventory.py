#!/usr/bin/python3
# -*- coding: utf-8 -*-
'''
GSP Excel to CSV

@contributor: Hassan Ahmed
@contact: ahmed.hassan.112.ha@gmail.com
@owner: Patrick Mahoney
@version: 1.0.7

This module is created to Convert GSP Excel Inventory Feed file to TSV
    - To be run in the same directory the data file is located
    - Portable to run in linux or windows
TODO: All after this
'''

HELP_MESSAGE = '''Usage:
    The script is capable of running without any argument provided. All behavorial
    variables will be reset to default.

    $ python3 suredone_download.py [options]

Parameters/Options:
    -h  | --help            : View usage help and examples
    -d  | --delimter        : Delimiter to be used as the separator in the CSV file saved by the script
        |                       - Default is comma ','.
    -f  | --file            : Path to the configuration file containing API keys
        |                       - Default in %APPDATA%/local/suredone.yaml on Window
        |                       - Default in $HOME/suredone.yaml
    -o  | --output          : Path for the output file to be downloaded at
        |                       - Default in %USERPROFILE%/Downloads/SureDone_Downloads_yyyy_mm_dd-hh-mm-ss.csv
        |                       - Default in $HOME/downloads/SureDone_Downloads_yyyy_mm_dd-hh-mm-ss.csv
    -p  | --preserve        : Do not delete older files that start with 'SureDone_' in the download directory
        |                       - This funciton is limited to default download locations only.
        |                       - Defining custom output path will render this feature useless.
    -v  | --verbose         : Show outputs in terminal as well as log file
    -w  | --wait            : Custom timeout for requests invoked by the script (specified in seconds)
        |                       - Default: 15 seconds

Example:
    $ python3 suredone_download.py

    $ python3 suredone_download.py -f [config.yaml]
    $ python3 suredone_download.py -file [config.yaml]

    $ python3 suredone_download.py -f [config.yaml] -o [output.csv]
    $ python3 suredone_download.py -file [config.yaml] --output_file [output.csv]

    $ python3 suredone_download.py -f [config.yaml] -o [output.csv] -v -p
    $ python3 suredone_download.py -file [config.yaml] --output_file [output.csv] --verbose --preserve
'''
# Need python version 3.4 or higher for pathlib
from pathlib import Path
import sys
import inspect
import csv
import os
import pandas as pd
from sys import platform

# Record python version for checking when the script kicks in
PYTHON_VERSION = float(sys.version[:sys.version.index(' ')-2])

def main(argv):
    '''
    Main function that execute the whole functionality.

    Pipline:
    --------
        - Check Python Version
        - Check operating system
        - Parse arguments
        - Verify:
            - File exists
            - File is not opened by someone
        - Read excel file
        - Fill na values by empty string or Int32Dtype (nullable type)
            - Maybe fill all of them with just an empty string?


    :param argv:
    :return:
    '''
    localFrame = inspect.currentframe()


if __name__ == '__main__':
    # sys.stdout = LOGGER
    # sys.excepthook = LOGGER.exceptionLogger
    main(sys.argv[1:])


"""
if platform == "linux" or platform == "linux2":
    # Declare variables using linux enviromnent $DOWNLOADS
    inputfile = os.environ.get('DOWNLOADS') + '/InventoryList.xls'
    inputsheet = os.environ.get('DOWNLOADS') + '/Sheet1'
    outputfile = os.environ.get('DOWNLOADS') + '/gsp_inventory.tsv'

elif platform == "darwin":
    print(platform)

elif platform == "win32":
    # Declare windows variables using Path from patlib
    inputfile = Path.home() / 'Downloads' / 'InventoryList.xls'
    inputsheet = 'Sheet1'
    outputfile = Path.home() / 'Downloads' / 'gsp_inventory.tsv'

'''
# Declare variable with list of columns to specify formats from file to import
my_columns={
    'GSP PN':str,
    'Rank':str,
    'POP':str,
    'BCA':str,
    'VIO':str,
    'Description':str,
    'UPC':str,
    'LTL':float
    }
'''

# Read GSP Bearings, Sheet1 and convert some data types to string
# data = pd.read_excel(inputfile, inputsheet, converters=my_columns)
data = pd.read_excel(inputfile)

# Removes ALL Blank lines
# data.dropna(inplace=True)

'''
# Rename Some Columns
data.rename(columns={'GSP PN':'PartNumber',
                     'BCA':'Interchangepartnumber',
                     'LTL':'Cost'},                   
                      inplace=True)
'''

'''
# Replace NaN with blank
data['Interchangepartnumber'] = data['Interchangepartnumber'].fillna('')
'''

'''
# Replace / with ' '
data['Interchangepartnumber'] = data['Interchangepartnumber'].str.replace('/', ' ')
'''

'''
# Create (Alter) data frame to add column [VendorID]
data['VendorID'] = '82'
'''

'''
# Create (Alter) data frame to add column [LineMasterID]
data['LineMasterID'] = '1089'
'''

'''
# Convert PartNumber data type to string
data['PartNumber'] = data['PartNumber'].astype(str)
'''

'''
# Create (Alter) data frame to add column [Parts]
data['Part'] = data['PartNumber'].str.strip(' /-')           
'''

# List Columns to save in tsv file
'''
# In this case I am saving all columns
my_list=list(data.columns.values)
'''

my_list = [
    'Site',
    'ItemNumber',
    'QuantityOnHand'
]

# my_list=list(data.columns.values)

# Write data frame by selected columns to csv file
data.to_csv(outputfile, encoding='utf-8', escapechar='\\', float_format='%.2f', index=False, columns=my_list,
            line_terminator='\r\n', quoting=csv.QUOTE_NONE, sep='\t')  # Create csv file for SQL Server to import
'''
    columns = my_list      - Only save selected columns from my_list
    encoding='utf-8'       - Use utf encoding
    float_format='%.2f'    - Set to 2 decimal places
    index=False            - Turn off row number
    quoting=csv.QUOTE_NONE - Don't surround text columns with double quotes
    sep=','                - Use comma as column delimiter
'''
"""