#!/usr/bin/python3
# -*- coding: utf-8 -*-
"""
GSP Excel to CSV

@contributor: Hassan Ahmed
@contact: ahmed.hassan.112.ha@gmail.com
@owner: Patrick Mahoney
@version: 1.5

This module is created to Convert GSP Excel Inventory Feed file to TSV
    - To be run in the same directory the data file is located
    - Portable to run in linux or windows
"""

HELP_MESSAGE = '''Usage:
    The script is capable of running without any argument provided. All behavorial variables will be reset to default.

    $ python[3] gsp_inventory.py [options]

Parameters/Options:
    -h  | --help            : View usage help and examples
    -i  | --input           : Path to the input file
        |                       - Linux: Defaults to the file in $HOME/Downloads
        |                       - Windows: Defaults to the file in %USERPROFILE%\\downloads
    -o  | --output          : Path to the output file
        |                       - Linux: Defaults to the file in $HOME/Downloads
        |                       - Windows: Defaults to the file in %USERPROFILE%\\downloads
    -d  | --delimter        : Single character to be used as delimiter for the tsv (default='\\t' (tab space))
    -p  | --preserve        : Do not delete original file if declared
    -v  | --verbose         : Show outputs in terminal as well as log file

Example:
    $ python3 suredone_download.py

    $ python3 suredone_download.py -f [GSPInventoryFeed.xlsx]
    $ python3 suredone_download.py -file [GSPInventoryFeed.xlsx]

    $ python3 suredone_download.py -f [GSPInventoryFeed.xlsx] -o [gsp_inventory.tsv]
    $ python3 suredone_download.py -file [GSPInventoryFeed.xlsx] --output_file [gsp_inventory.tsv]

    $ python3 suredone_download.py -f [GSPInventoryFeed.xlsx] -o [gsp_inventory.tsv] -v -p
    $ python3 suredone_download.py -file [GSPInventoryFeed.xlsx] --output_file [gsp_inventory.tsv] --verbose --preserve
'''
# Need python version 3.4 or higher for pathlib
from pathlib import Path
import sys
import inspect
import time
import csv
import os
import pandas as pd
from datetime import datetime
import traceback
import getopt
import shutil

currentMilliTime = lambda: int(round(time.time() * 1000))

# Time tracking variables
RUN_TIME = currentMilliTime()
START_TIME = datetime.now()

# Record python version and platform for checking when the script kicks in
PYTHON_VERSION = float(sys.version[:sys.version.index(' ') - 2])
if sys.platform == 'win32' or sys.platform == 'win64':  # Windows
    PLATFORM = 'windows'
elif sys.platform == 'linux' or sys.platform == 'linux2':  # Linux
    PLATFORM = 'linux'


def main(argv):
    """
    Main function that execute the whole functionality.

    Pipline:
    --------
        - Check Python Version
        - Check operating system
        - Parse arguments
        - Find defaults if input and output not provided
        - Verify:
            - File exists
            - File is not opened by someone
        - Read excel file
        - Fill na values by empty string or Int32Dtype (nullable type)
            - Maybe fill all of them with just an empty string?
        - Save the file to .tsv

    :param argv: arguments coming from the commandline
    :return:
    """
    localFrame = inspect.currentframe()

    # Verify python version and platform type
    checkPlatformAndPythonVersion()
    LOGGER.writeLog("Platform type and python version verified.", localFrame.f_lineno)

    # Parse arguments
    inputFilePath, outputFilePath, delimiter, preserveOldFiles, verbose = parseArgs(argv)
    LOGGER.writeLog("Args parsed...", localFrame.f_lineno)
    LOGGER.writeLog("Input file path: {}".format(inputFilePath), localFrame.f_lineno)
    LOGGER.writeLog("Output file path: {}".format(outputFilePath), localFrame.f_lineno)
    LOGGER.writeLog("Using delimiter: {}".format("[TAB SPACE]" if delimiter == '\t' else delimiter),
                    localFrame.f_lineno)
    LOGGER.writeLog("Preserve input file: {}".format("NO" if not preserveOldFiles else "YES"), localFrame.f_lineno)
    LOGGER.writeLog("Verbose: {}".format("OFF" if not verbose else "ON"), localFrame.f_lineno)
    LOGGER.writeLog("===============================================", localFrame.f_lineno)

    # Read file
    data = pd.read_excel(inputFilePath)
    LOGGER.writeLog("File loaded...", localFrame.f_lineno)

    # Fill nas with null values so they can be interpreted as null in SQL Server
    data['Site'].fillna('', inplace=True)
    data['ItemNumber'].fillna('', inplace=True)
    data['QuantityOnHand'].fillna(0, inplace=True)

    # Convert quantity in hand to integer
    data['QuantityOnHand'] = data['QuantityOnHand'].astype(int)
    LOGGER.writeLog("Data processed.", localFrame.f_lineno)

    # Save file as tsv
    columnList = [
        'Site',
        'ItemNumber',
        'QuantityOnHand'
    ]
    # TODO: Add the logic where when the delimiter is tab, then extension is tsv and when the delimiter is comma, the
    #   extension is tsv, else txt.
    data.to_csv(outputFilePath, encoding='utf-8', escapechar='\\', float_format='%.2f', index=False, columns=columnList,
                line_terminator='\r\n', quoting=csv.QUOTE_NONE, sep=delimiter)

    LOGGER.writeLog("File saved as .tsv at path: {}".format(inputFilePath), localFrame.f_lineno)

    # Time to remove the original file (If preserve is declared as a command line flag)
    if not preserveOldFiles:
        os.remove(inputFilePath)
        LOGGER.writeLog("Removed input file.", localFrame.f_lineno)
    LOGGER.writeLog("Execution complete - exitting.", localFrame.f_lineno)


def parseArgs(argv):
    """
    Function that parses the arguments sent from the command line
    and returns the behavioral variables to the caller.

    :param argv: str: Arguments sent through the command line
    :return:
        inputPath: str: path to the directory where input file is present
            - Linux: Defaults to $HOME/Downloads in
            - Windows: Defaults to %USERPROFILE%\\downloads\\
<br>    delimiter: str: Single character to be used as delimiter for the tsv (default='\t')
<br>    outputPath: str: Output directory where output file is to be stored
            - Linux: Defaults to $HOME/Downloads in
            - Windows: Defaults to %USERPROFILE%\downloads\
        preserve: boolean: Determines whether to remove all occurrences of the input file (default=False)
        verbose: boolean: Show log outputs in the console
    """
    localFrame = inspect.currentframe()
    # Defining options in for command line arguments
    options = "hi:o:d:vp"
    long_options = ["help", "input=", "output=", 'delimiter=', 'verbose', 'preserve']
    inputFileExtension = '.xls'
    inputFileName = 'GSPInventoryFeed' + inputFileExtension

    # Note about validating if file is opened or not
    #   - One of the files is an Excel 1997-compatiblity Mode xls which opens regardless of whether it is being used by
    #   someone else
    #   - Second file a csv which also gives no problem in reading if it is opened by someone else
    # print(pd.read_csv(inputFilePath))

    # This is the same for now. Maybe need to change separately for both OSes
    if PLATFORM == 'windows':
        inputDefaultPath = os.path.join(os.path.expanduser('~'), 'Downloads', inputFileName)
    elif PLATFORM == 'linux':
        inputDefaultPath = os.path.join(os.path.expanduser('~'), 'Downloads', inputFileName)

    # This is the same for now. Maybe need to change separately for both OSes
    if PLATFORM == 'windows':
        outputDefaultPath = os.path.join(os.path.expanduser('~'), 'Downloads', 'gsp_inventory.tsv')
    elif PLATFORM == 'linux':
        outputDefaultPath = os.path.join(os.path.expanduser('~'), 'Downloads', 'gsp_inventory.tsv')

    # Arguments
    inputFilePath = inputDefaultPath
    outputFilePath = outputDefaultPath
    defaultDilimiter = '\t'
    delimiter = defaultDilimiter
    verbose = False
    preserveOldFiles = False

    # Extracting arguments
    try:
        opts, args = getopt.getopt(argv, options, long_options)
    except getopt.GetoptError:
        # Not logging here since this is a command-line feature and must be printed on console
        print("Error in arguments!")
        print(HELP_MESSAGE)
        exit()

    for option, value in opts:
        if option == '-h':
            # Turn on verbose, print help message, and exit
            LOGGER.verbose = True
            print(HELP_MESSAGE)
            sys.exit()
        elif option in ("-i", "--input"):
            inputFilePath = value
        elif option in ("-o", "--output"):
            outputFilePath = value
        elif option in ("-d", "--delimiter"):
            delimiter = value
            delimiter = validateDelimiter(delimiter, defaultDilimiter)
        elif option in ("-p", "--preserve"):
            preserveOldFiles = True
        elif option in ("-v", "--verbose"):
            verbose = True

    # Updating logger's behavior based on verbose
    LOGGER.verbose = verbose

    # Validate input file path
    if not os.path.exists(inputFilePath) or os.path.isdir(inputFilePath) or not inputFilePath.endswith(
            inputFileExtension):
        LOGGER.writeLog(
            """Invalid file path. Check if it exists, is not a directory and has {} extension. Exiting.""".format(
                inputFileExtension),
            localFrame.f_lineno, severity='code-breaker', data={'code': 1})
        exit()

    # Validate output file directory path
    if not os.path.exists(os.path.dirname(outputFilePath)) or os.path.isdir(outputFilePath):
        LOGGER.writeLog(
            """Invalid output file path. Check if it exists and is not a directory.
             \rReverting to defaults.""",
            localFrame.f_lineno, severity='warning', data={'code': 1})
        outputFilePath = outputDefaultPath

    return inputFilePath, outputFilePath, delimiter, preserveOldFiles, verbose


def validateDelimiter(delimiter, defaultDilimiter):
    """
    Function that validates the delimiter option input by the user.
    Main issues to check for is length and make sure that the chosen delimiter is within a list of acceptable options.
    :param delimiter: str: The user-specified delimiter option
    :param defaultDilimiter: str: The user-specified delimiter option
    :return:
<br>    delimiter : str: The same delimiter if validated and a ',' as a delimiter if not validated.
    """
    localFrame = inspect.currentframe()
    # Account for '\\t' and '\t'
    if delimiter == '\\t':
        delimiter = '\t'

    # Check for length
    if len(delimiter) > 1:
        LOGGER.writeLog("Length of the delimiter was greater than one character, switching to default ',' delimiter.",
                        localFrame.f_lineno, severity='warning')
        delimiter = ','
        return delimiter

    # Check that it's within acceptable options
    acceptableDelimiters = [',', '\t', ':', '|', ' ']

    if delimiter not in acceptableDelimiters:
        LOGGER.writeLog("Delimiter was not selected from acceptable options, switching to ',' default delimiter.",
                        localFrame.f_lineno, severity='warning')
        delimiter = defaultDilimiter
        return delimiter

    return delimiter


def checkPlatformAndPythonVersion():
    """
    Function that checks python version and platform
    Will exit the code with an error entry in the log if requirements not satisfied
    :return:
    """
    localFrame = inspect.currentframe()
    # Check if python version is 3.5 or higher
    """
    # NOTE:
    # More precision is not required since python is a very compatible and platform-free language (Windows Python 3.6 
    # and Linux Python 3.8 can easily run the same file without any errors.
    """
    if not PYTHON_VERSION >= 3.5:
        LOGGER.writeLog("Must use Python version 3.5 or higher!", localFrame.f_lineno, severity='code-breaker',
                        data={'code': 1})
        exit()

    # Check if the platform is either windows or linux
    if PLATFORM not in ('windows', 'linux'):
        LOGGER.writeLog("Please use Windows or Linux platform.", localFrame.f_lineno, severity='code-breaker',
                        data={'code': 1})


""" Custom Exceptions that will be caught by the script """


class Logger(object):
    """ The logger class that will handle all outputs, may it be console or log file. """

    def __init__(self, verbose=False):
        self.terminal = sys.stdout
        self.log = open(self.getLogPath(), "a")
        # Write the header row
        self.log.write(' Ind. |LineNo.| Time stamp  : Message')
        self.log.write('\n=====================================\n')
        self.verbose = verbose

    def getLogPath(self):
        """
        Function that will determine the default log file path based on the operating system being used.
        Will also create appropriate directories they aren't present.

        Returns
        -------
            - logFile : fileIO
                File IO for the whole script to log to.
        """
        # Define the file name for logging
        temp = datetime.now().strftime('%Y_%m_%d-%H-%M-%S')
        logFileName = "gsp_inventory_xlsx2tsv_" + temp + ".log"

        # If the platform is windows, set the log file path to the current user's Downloads/log folder
        if sys.platform == 'win32' or sys.platform == 'win64':  # Windows
            logFilePath = os.path.expandvars(r'%USERPROFILE%')
            logFilePath = os.path.join(logFilePath, 'Downloads')
            logFilePath = os.path.join(logFilePath, 'log')
            if os.path.exists(logFilePath):
                return os.path.join(logFilePath, logFileName)
            else:  # Create the log directory
                os.mkdir(logFilePath)
                return os.path.join(logFilePath, logFileName)

        # If Linux, set the download path to the $HOME/downloads folder
        elif sys.platform == 'linux' or sys.platform == 'linux2':  # Linux
            logFilePath = os.path.expanduser('~')
            logFilePath = os.path.join(logFilePath, 'log')
            if os.path.exists(logFilePath):
                return os.path.join(logFilePath, logFileName)
            else:  # Create the log directory
                os.mkdir(logFilePath)
                return os.path.join(logFilePath, logFileName)

    def write(self, message):
        if self.verbose:
            self.terminal.write(message)
            self.terminal.flush()
        self.log.write(message)

    def writeLog(self, message, lineNumber, severity='normal', data=None):
        """
        Function that writes out to the log file and console based on verbose.
        The function will change behavior slightly based on severity of the message.

        :param message: str: Message to write
        :param lineNumber: int: File line number that created this log entry.
        :param severity: str: Defines what the message is related to. Is the message:
                    - [N] : A 'normal' notification
                    - [W] : A 'warning'
                    - [E] : An 'error'
                    - [!] : A 'code-breaker error' (errors that are followed by the script exitting)
        :param data: dict: A dictionary that will contain additional information when a code-breaker error occurs
                Attributes:
                    - code : error code
                        1 : Generic error, only print the message.
                        2 : An API call was not successful. Response object attached.
                        3 : YAML loading error. Error object attached
                    - response : str
                        JSON-like str - the response recieved from the request in conern at the point of error.
                    - error : str
                        String produced by exception if an exception occured
        """
        # Get a timestamp
        timestamp = self.getCurrentTimestamp()

        # Format the message based on severity
        lineNumber = str(lineNumber)
        if severity == 'normal':
            indicator = '[N]'
            toWrite = ' ' + indicator + '  |  ' + lineNumber + '  | ' + timestamp + ': ' + message
        elif severity == 'warning':
            indicator = '[W]'
            toWrite = ' ' + indicator + '  |  ' + lineNumber + '  | ' + timestamp + ': ' + message
        elif severity == 'error':
            indicator = '[X]'
            toWrite = ' ' + indicator + '  |  ' + lineNumber + '  | ' + timestamp + ': ' + message
        elif severity == 'code-breaker':
            indicator = '[!]'
            toWrite = ' ' + indicator + '  |  ' + lineNumber + '  | ' + timestamp + ': ' + message

            if data['code'] == 2:  # Response recieved but unsuccessful
                details = '\n[ErrorDetailsStart]\n' + data['response'] + '\n[ErrorDetailsEnd]'
                toWrite = toWrite + details
            elif data['code'] == 3:  # YAML loading error
                details = '\n[ErrorDetailsStart]\n' + data['error'] + '\n[ErrorDetailsEnd]'
                toWrite = toWrite + details

        # Write out the message
        self.log.write(toWrite + '\n')
        if self.verbose:
            self.terminal.write(message + '\n')
            self.terminal.flush()

    def getCurrentTimestamp(self):
        """
        Simple function that calculates the current time stamp and simply formats it as a string and returns.
        Mainly aimed for logging.

        Returns
        -------
            - timestamp : str
                A formatted string of current time
        """
        return datetime.now().strftime("%H:%M:%S.%f")[:-3]

    def exceptionLogger(self, exctype, value, traceBack):
        """
        A simple printing function that will take place of the sys.excepthook function and print the results to the log instead of the console.

        Parameters
        ----------
            - exctype : object
                Exception type and details
            - Value : str
                The error passed while the exception was raised
            - traceBack : traceback object
                Contains information about the stack trace.
        """
        LOGGER.write('Exception Occured! Details follow below.\n')
        LOGGER.write('Type:{}\n'.format(exctype))
        LOGGER.write('Value:{}\n'.format(value))
        LOGGER.write('Traceback:\n')
        for i in traceback.format_list(traceback.extract_tb(traceBack)):
            LOGGER.write(i)

    def flush(self):
        # This flush method is needed for python 3 compatibility.
        # This handles the flush command by doing nothing.
        # You might want to specify some extra behavior here.
        pass


# Determine log file path
LOGGER = Logger(verbose=False)
if __name__ == '__main__':
    sys.stdout = LOGGER
    sys.excepthook = LOGGER.exceptionLogger
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

# Create (Alter) data frame to add column [LineMaster
'''ID]
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
