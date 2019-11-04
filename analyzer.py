import sys
import random
from os import listdir, path, makedirs
from os.path import isfile, join
from striprtf.striprtf import rtf_to_text
import pandas as pd
import xlsxwriter

__excelFileName = "data.xlsx"


def perform_commands():
    """ handle command line interface Analyzer tool"""
    args = sys.argv
    args = args[1:]  # Skip file name

    if len(args) == 0:
        print('You should pass command, please use --help for more info')
    else:
        command = args[0]
        if command == '--help':
            print('Analyzer command line interface')
            print('Commands:')
            print('   --start     ==> Start analyze and parse files')
            print('   --excel     ==> Show current standard excel file name in app')
            print('   --help      ==> show list of commands')

        elif command == '--start':
            if len(args)is not 1:
                print(
                    "Incorrect format, please use --help fot more info")
                exit(1)
            startAnalyzeProcess()
        elif command == '--excel':
            if len(args)is not 1:
                print(
                    "Incorrect format, please use --help fot more info")
                exit(1)
            print_green(
                "Excel file name should be: [{0}]".format(__excelFileName))
        else:
            print('Unrecognized argument.')


# --- Terminal Color
TC_RED = "\033[1;31m"
TC_BLUE = "\033[1;34m"
TC_CYAN = "\033[1;36m"
TC_GREEN = "\033[0;32m"
TC_RESET = "\033[0;0m"
TC_BOLD = "\033[;1m"
TC_REVERSE = "\033[;7m"
# ---


def print_blue(msg):
    sys.stdout.write(TC_BLUE)
    print(msg)
    sys.stdout.write(TC_RESET)


def print_green(msg):
    sys.stdout.write(TC_GREEN)
    print(msg)
    sys.stdout.write(TC_RESET)


def print_red(msg):
    sys.stdout.write(TC_RED)
    print(msg)
    sys.stdout.write(TC_RESET)


def progress_bar(count, total, suffix=''):
    bar_len = 60
    filled_len = int(round(bar_len * count / float(total)))
    percents = round(100.0 * count / float(total), 1)
    bar = '=' * filled_len + '-' * (bar_len - filled_len)
    sys.stdout.write('[%s] %s%s ...%s\r' % (bar, percents, '%', suffix))
    sys.stdout.flush()  # As suggested by Rom Ruben


def makeResultsDirectory():
    try:
        result_dir = "./results"
        if not path.exists(result_dir):
            makedirs(result_dir)
    except:
        print_red(
            'Failed to create [results] directory. please create it manually in the current directory next to this file.')


def convert_RTF_to_PlainTextList(filename):
    text = ""
    with open('./ECs/{0}'.format(filename)) as file:
        data = file.read()
        text = rtf_to_text(data)
    return text.splitlines()


def write_text_file(targetFileName, data_list):
    with open("./results/{0}".format(targetFileName), "w", encoding="utf-8") as f:
        for item in data_list:
            f.write("%s\n" % item)


def write_log_file(targetFileName, data_list):
    with open("./{0}".format(targetFileName), "w", encoding="utf-8") as f:
        for item in data_list:
            f.write("%s\n" % item)


def append_to_log_file(targetFileName, data_list):
    with open("./{0}".format(targetFileName), "a", encoding="utf-8") as f:
        for item in data_list:
            f.write("%s\n" % item)


def getAllFilesInDirectory():
    scan_path = "./ECs/"
    onlyRtfFiles = [f for f in listdir(scan_path) if isfile(
        join(scan_path, f)) and f.endswith(".rtf") and not f.startswith("~$")]
    return onlyRtfFiles  # as list


def splitFileNameWithDash(str_name):
    return str_name.split('-')


def removeFileExtention(str_name):
    return path.splitext(str_name)[0]


def readExcelFile():
    file = './'+__excelFileName
    df = pd.read_excel(file)
    return df


def iterateOverDataFrame(df):
    # iterate over rows with iterrows()
    for index, row in df.head().iterrows():
        # access data using column names
        print(index, row['gvkey'], row['conm'], row['CEO name'])


def findRowItemInExcelFileByKey(df, key):
    filter = df[df['gvkey'] == int(key)]
    if filter.shape[0] != 0:
        return filter
    else:
        return None


def get_CEO_from_dataFrameRow(df_one_row):
    if df_one_row is None:  # if row was None (row with key was not founded)
        return None

    ceo = df_one_row['CEO name']
    if ceo.isna().any().any():  # if cell was nan
        return None
    else:
        return ceo.values[0]  # cel has value


def write_excel_file(file_name, data_list):
    """ create excel file. Note: the file_name parameter should be without a file extension """
    file_name = './results/{0}.xlsx'.format(file_name)
    try:
        # Create an new Excel file and add a worksheet.
        workbook = xlsxwriter.Workbook(file_name)
        worksheet = workbook.add_worksheet()
        # set width to the first column to make the text clearer.
        worksheet.set_column('A:A', 50)
        # Write some numbers, with row/column notation.
        rowCounter = 0
        for line in data_list:
            if len(line.strip()) == 0:  # skip empty lines
                continue
            worksheet.write(rowCounter, 0, line)  # row,column,data
            rowCounter += 1
        # finalize file and close stream.
        workbook.close()
    except:
        pass


def parser_isTextStartWith_CEO_name(text, ceo_name):
    return text.startswith('{0}:'.format(ceo_name).upper()) or text.startswith('{0},'.format(ceo_name).upper())


def parser_isTextStartWithAnyCommander(text):
    return text[:4].strip().isupper()


def parser_isTextStartWithQuestionsAndAnswers(text):
    CONST_QA_PHRASE = 'Questions and Answers'
    return text.strip().startswith(CONST_QA_PHRASE)


def startAnalyzeProcess():
    makeResultsDirectory()
    # --- read excel file
    df = readExcelFile()  # DataFrame from whole excel file
    list_of_files = getAllFilesInDirectory()
    total_file_count = len(list_of_files)  # for report purpose
    # --- report
    _target_Error_log = []
    _target_Error_log_ExcelFiles = []
    print_blue('Total number of RTF files: [{0}]'.format(len(list_of_files)))
    progress_counter = 0
    for rtf_file in list_of_files:
        # progress bar
        progress_counter += 1
        progress_bar(progress_counter, total_file_count, "Analyzing")
        try:
            rtf_name_without_extention = removeFileExtention(rtf_file)
            file_name_parts = splitFileNameWithDash(rtf_name_without_extention)
            key = file_name_parts[0]  # find reference key from file name
            reference_row = findRowItemInExcelFileByKey(df, key)
            CEO_name = get_CEO_from_dataFrameRow(reference_row)
            texts_list = convert_RTF_to_PlainTextList(rtf_file)
            _targetOutputList_Before_QA = []  # result output to write
            _targetOutputList_After_QA = []  # result output to write
            is_QA_seen = False
            is_CEO_seen = False  # this check if CEO talked more that one paragraph
            for line in texts_list:
                # ------------------ try to parse text
                # check before and after of Q&A
                if is_QA_seen == False and parser_isTextStartWithQuestionsAndAnswers(line):
                    is_QA_seen = True
                    continue

                if parser_isTextStartWith_CEO_name(line, CEO_name):
                    is_CEO_seen = True
                    if is_QA_seen:
                        _targetOutputList_After_QA.append(line)
                    else:
                        _targetOutputList_Before_QA.append(line)
                else:
                    if parser_isTextStartWithAnyCommander(line):
                        is_CEO_seen = False
                    else:
                        # this paragraph is continue part of CEO speak.
                        if is_CEO_seen:
                            if is_QA_seen:
                                _targetOutputList_After_QA.append(line)
                            else:
                                _targetOutputList_Before_QA.append(line)
                # ------------------
        except:
            _target_Error_log.append(
                'Error at parsing -> File[{0}]'.format(rtf_file))
            continue

        # after parse text, write results to files
        # Text files
        write_text_file(
            '{0}-{1}.txt'.format(rtf_name_without_extention, "BQ&A"), _targetOutputList_Before_QA)
        write_text_file(
            '{0}-{1}.txt'.format(rtf_name_without_extention, "AQ&A"), _targetOutputList_After_QA)
        # Excel files
        write_excel_file(
            '{0}-{1}'.format(rtf_name_without_extention, "BQ&A"), _targetOutputList_Before_QA)
        write_excel_file(
            '{0}-{1}'.format(rtf_name_without_extention, "AQ&A"), _targetOutputList_After_QA)

    # after process completed - end of jobs
    # create Log file
    write_log_file("Log.txt", _target_Error_log)
    print("")  # new line
    print_blue("=========== Status ===========")
    print_green('The Total number of successful parses    :[{0}]'.format(
        total_file_count-len(_target_Error_log)))
    print_red('The Total number of parsing errors       :[{0}] --==> see Log.txt'.format(
        len(_target_Error_log)))
    print_blue("==============================")


# this makes perform_commands() runs automatically
if __name__ == '__main__':
    perform_commands()
