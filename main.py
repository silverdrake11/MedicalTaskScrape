import os.path
import sys

import extract


def get_extension(filepath):
    '''Return the lowercase file extension. Error if none exists'''
    extension = os.path.splitext(filepath)[1]
    extension = extension.lstrip('.').lower().strip()
    if not extension:
        print("ERROR: Filename must have an extension e.g. 'hello.xlsx' and not just 'hello' as a filename!\n")
        sys.exit()
    return extension


def check_html_extension(extension):
    '''Makes sure that it's an HTML type extension'''
    if extension == 'html' or extension == 'htm':
        return
    print("ERROR: Input file must be an {} file and not a {} file!\n".format('HTML', extension.upper()))
    sys.exit()


def check_csv_extension(extension):
    '''Makes sure it's a csv file'''
    if extension != 'csv':
        print("ERROR: Output file must be an {} and not a {} file!\n".format('CSV', extension.upper()))
        sys.exit()


def check_if_exists(filepath):
    if not os.path.isfile(filepath):
        print("ERROR: The file you entered (see below) does not exist!")
        print('  ' + filepath + "\n")
        sys.exit()


def check_write_lock(filepath):
    '''Errors out if another program such as Excel places a write lock on the output CSV file'''
    if os.path.isfile(filepath):
        try:
            csvfile = open(filepath, 'a')
        except Exception:
            print("ERROR: The output CSV below is locked! Please close Excel and try again!")
            print('  ' + filepath + '\n')
            sys.exit()


def check_directory(filepath):
    '''Creates output directories if they don't exist'''
    dirpath = os.path.dirname(filepath)
    if not os.path.isdir(dirpath):
        os.makedirs(dirpath)


def check_input(html_filepath):
    '''Error checking on the input filepath'''
    html_filepath = os.path.abspath(html_filepath)
    extension = get_extension(html_filepath)
    check_html_extension(extension)
    check_if_exists(html_filepath)


def check_output(csv_filepath):
    '''Error checking on the output filepath'''
    csv_filepath = os.path.abspath(csv_filepath)
    extension = get_extension(csv_filepath)
    check_csv_extension(extension)
    check_write_lock(csv_filepath)
    check_directory(csv_filepath)


def parse_user_input():

    print('\n')

    if len(sys.argv) == 1: # Interactive mode

        html_filepath = input("Please enter the path to the HTML file:\n")
        print('')
        check_input(html_filepath)

        csv_filepath = input("Please enter the path to the output CSV file:\n")
        print('')
        check_output(csv_filepath)

    elif len(sys.argv) == 3:

        print("USAGE ARGS: <input HTML filepath> <output CSV filepath>\n")

        html_filepath = sys.argv[1]
        csv_filepath = sys.argv[2]
        check_input(html_filepath)
        check_output(csv_filepath)

    else:

        print("USAGE ARGS: <input HTML filepath> <output CSV filepath>\n")
        print("ERROR: Incorrect number of arguments!\n")
        sys.exit()

    print("Processing data...\n")
    extract.write_csv(html_filepath, csv_filepath)
    print("SUCCESS!")
    print("File written to:")
    print(os.path.abspath(csv_filepath))
    print('')
    if len(sys.argv) == 1: # In interactive mode, to prevent cmd from exiting
        input("Press ENTER to exit..\n")


def user_input_helper(): 
    try:
        parse_user_input()
    except SystemExit:
        if len(sys.argv) == 1: # In interactive mode, to prevent cmd from exiting
            print('')
            input("Press ENTER to try again..\n")
            user_input_helper()


user_input_helper()