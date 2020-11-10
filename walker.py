# -*- coding: utf-8 -*-
'''
    Clean up Walker Inventory feed
'''
import csv
import pandas as pd

inputfile='./walker.csv'
outputfile='walker.tsv'

# Declare variable with list of columns to specify formats from file to import
my_columns={
            'part description':str,
            'part number':str,
            'part MO inventory':str,
            'part GG inventory':str
            }

# Read suredone.csv
data = pd.read_csv(inputfile, converters=my_columns, skiprows=0)

data['part description'] = data['part description'].str.replace(',', '')

# List Columns to save in tsv file
# In this case I am saving all columns
my_list=list(data.columns.values)

'''
my_list = ['VendorID',
         'LineMasterID',
         'Part',
         'PartNumber',
         'Interchangepartnumber',
         'Description',
         'UPC',
         'Cost']
'''

# Write data frame by selected columns to csv file
data.to_csv(outputfile, encoding='utf-8', escapechar='\\', float_format='%.2f', index=False, columns = my_list, line_terminator='\r\n', quoting=csv.QUOTE_NONE, sep='\t') # Create csv file for SQL Server to import
'''
    columns = my_list      - Only save selected columns from my_list
    encoding='utf-8'       - Use utf encoding
    float_format='%.2f'    - Set to 2 decimal places
    index=False            - Turn off row number
    quoting=csv.QUOTE_NONE - Don't surround text columns with double quotes
    sep=','                - Use comma as column delimiter
'''
