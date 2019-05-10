#
# updateCommon.py -- common methods and variables for updates
#
# some code from account_analysis.py by Mark Jenkins, ParIT Worker Co-operative <mark@parit.ca>
# some code from Google quickstart spreadsheets examples
#
# @author Mark Sattolo <epistemik@gmail.com>
# @version Python 3.6
# @created 2019-04-07
# @updated 2019-04-23

from sys import stdout, exit
from datetime import date, timedelta, datetime as dt
from bisect import bisect_right
from decimal import Decimal
from math import log10
import csv
import json
import pickle
import os.path as osp
from gnucash import GncNumeric
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import inspect

COLOR_FLAG = '\x1b['
BLACK   = COLOR_FLAG + '30m'
RED     = COLOR_FLAG + '91m'
GREEN   = COLOR_FLAG + '92m'
YELLOW  = COLOR_FLAG + '93m'
BLUE    = COLOR_FLAG + '94m'
MAGENTA = COLOR_FLAG + '95m'
CYAN    = COLOR_FLAG + '96m'
WHITE   = COLOR_FLAG + '97m'
COLOR_OFF = COLOR_FLAG + '0m'


def print_info(text, color='', inspector=True, newline=True):
    """
    Print information with choices of color, inspection info, newline
    """
    inspect_line = ''
    if text is None:
        text = '================================================================================================================='
        inspector = False
    if inspector:
        calling_frame = inspect.currentframe().f_back
        calling_file  = inspect.getfile(calling_frame).split('/')[-1]
        calling_line  = str(inspect.getlineno(calling_frame))
        inspect_line  = '[' + calling_file + '@' + calling_line + ']: '
    print(inspect_line + color + str(text) + COLOR_OFF, end=('\n' if newline else ''))


def print_error(text, newline=True):
    """
    Print Error information in RED with inspection info
    """
    calling_frame = inspect.currentframe().f_back
    parent_frame = calling_frame.f_back
    calling_file = inspect.getfile(calling_frame).split('/')[-1]
    calling_line = str(inspect.getlineno(calling_frame))
    parent_line = str(inspect.getlineno(parent_frame))
    inspect_line = '[' + calling_file + '@' + calling_line + '/' + parent_line + ']: '
    print(inspect_line + RED + str(text) + COLOR_OFF, end=('\n' if newline else ''))


# constant strings
QTR = 'Quarter'
YR  = 'Year'
MTH = 'Month'

# number of months in the period
PERIOD_QTR = 3

NUM_MONTHS = 12
ONE_DAY = timedelta(days=1)
ZERO = Decimal(0)

CREDENTIALS = 'secrets/credentials.json'

SHEETS_RW_SCOPE = ['https://www.googleapis.com/auth/spreadsheets']

SHEETS_EPISTEMIK_RW_TOKEN = {
    'P2' : 'secrets/token.sheets.epistemik.rw.pickle2' ,
    'P3' : 'secrets/token.sheets.epistemik.rw.pickle3' ,
    'P4' : 'secrets/token.sheets.epistemik.rw.pickle4'
}

# Spreadsheet ID
BUDGET_QTRLY_ID_FILE = 'secrets/Budget-qtrly.id'
# sheet names in Budget Quarterly
ALL_INC_SHEET    = 'All Inc 1'
ALL_INC_2_SHEET  = 'All Inc 2'
NEC_INC_SHEET    = 'Nec Inc 1'
NEC_INC_2_SHEET  = 'Nec Inc 2'
BAL_1_SHEET      = 'Balance 1'
BAL_2_SHEET      = 'Balance 2'
QTR_ASTS_SHEET   = 'Assets 1'
QTR_ASTS_2_SHEET = 'Assets 2'
ML_WORK_SHEET    = 'ML Work'
CALCULNS_SHEET   = 'Calculations'

TOKEN = SHEETS_EPISTEMIK_RW_TOKEN['P4']

# first data row in Budget-qtrly.gsht
BASE_ROW = 3

today = dt.now()
now = today.strftime("%Y-%m-%dT%H-%M-%S")


def year_span(year_diff, base_year_span, hdr_span):
    """
    calculate which row to update, factoring in the header row placed every so-many years
    :param      year_diff: int: difference in years to calculate for
    :param base_year_span: int: number of rows between equivalent positions in adjacent years, not including header rows
    :param       hdr_span: int: number of rows between header rows
    :return: int
    """
    return (year_diff * base_year_span) + (year_diff // hdr_span)


def get_year(str_year, base_year):
    """
    convert the string representation of a year to an int
    :param  str_year: string: to convert
    :param base_year:    int: earliest possible year
    :return: int
    """
    if not str_year.isnumeric():
        print_error("Input MUST be the String representation of a Year, e.g. '2013'!")
        exit(79)
    int_year = int(float(str_year))
    if int_year > today.year or int_year < base_year:
        print_error("Input MUST be the String representation of a Year between {} and {}!".format(today.year, base_year))
        exit(83)

    return int_year


def get_quarter(str_qtr):
    """
    convert the string representation of a quarter to an int
    :param  str_qtr: string: to convert
    :return: int
    """
    if not str_qtr.isnumeric():
        print_error("Input MUST be a String of 0..4!")
        exit(138)
    int_qtr = int(float(str_qtr))
    if int_qtr > 4 or int_qtr < 0:
        print_error("Input MUST be a String of 0..4!")
        exit(142)

    return int_qtr


def fill_cell(sheet, col, row, val, data_list=None):
    """
    Create a dictionary to contain Google Sheets update information for one cell and add to the submitted list
    :param     sheet:  string:  particular sheet in my Google spreadsheet to update
    :param       col:  string:  column
    :param       row:     int
    :param       val:  string | Decimal: value to send as string
    :param data_list:    list:  to append with created dict, if needed
    :return: data_list
    """
    if data_list is None:
        data_list = list()
    value = val.to_eng_string() if isinstance(val, Decimal) else val
    cell = {'range': sheet + '!' + col + str(row), 'values': [[value]]}
    print_info("cell = {}\n".format(cell))
    data_list.append(cell)
    return data_list


def get_budget_id():
    """
    get the budget id string from the file in the secrets folder
    :return: string
    """
    fp = open(BUDGET_QTRLY_ID_FILE, "r")
    fid = fp.readline().strip()
    print_info("Budget Id = '{}'\n".format(fid), CYAN)
    fp.close()
    return fid


def get_credentials():
    """
    get the proper credentials needed to write to the Google spreadsheet
    :return: pickle object
    """
    creds = None
    if osp.exists(TOKEN):
        with open(TOKEN, 'rb') as token:
            creds = pickle.load(token)

    # if there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS, SHEETS_RW_SCOPE)
            creds = flow.run_local_server()
        # save the credentials for the next run
        with open(TOKEN, 'wb') as token:
            pickle.dump(creds, token, pickle.HIGHEST_PROTOCOL)

    return creds


# noinspection PyUnresolvedReferences
def gnc_numeric_to_python_decimal(numeric):
    """
    convert a GncNumeric value to a python Decimal value
    :param numeric: GncNumeric: value to convert
    :return: Decimal
    """
    negative = numeric.negative_p()
    sign = 1 if negative else 0

    val = GncNumeric(numeric.num(), numeric.denom())
    result = val.to_decimal(None)
    if not result:
        raise Exception("GncNumeric value '{}' CANNOT be converted to decimal!".format(val.to_string()))

    digit_tuple = tuple(int(char) for char in str(val.num()) if char != '-')
    denominator = val.denom()
    exponent = int(log10(denominator))
    assert( (10 ** exponent) == denominator )
    return Decimal((sign, digit_tuple, -exponent))


def next_quarter_start(start_year, start_month):
    """
    get the year and month that starts the FOLLOWING quarter
    :param  start_year: int
    :param start_month: int
    :return: int, int
    """
    # add number of months for a Quarter
    next_month = start_month + PERIOD_QTR

    # use integer division to find out if the new end month is in a different year,
    # what year it is, and what the end month number should be changed to.
    next_year = start_year + ( (next_month - 1) // NUM_MONTHS )
    next_month = ( (next_month - 1) % NUM_MONTHS ) + 1

    return next_year, next_month


def current_quarter_end(start_year, start_month):
    """
    get the year and month that ends the CURRENT quarter
    :param  start_year: int
    :param start_month: int
    :return: date
    """
    end_year, end_month = next_quarter_start(start_year, start_month)

    # last step, the end date is one day back from the start of the next period
    # so we get a period end like 2010-03-31 instead of 2010-04-01
    return date(end_year, end_month, 1) - ONE_DAY


def generate_quarter_boundaries(start_year, start_month, num_qtrs):
    """
    get the start and end dates for the quarters in the submitted range
    :param  start_year: int
    :param start_month: int
    :param    num_qtrs: int: number of quarters to calculate
    :return: date, date
    """
    for i in range(num_qtrs):
        yield(date(start_year, start_month, 1), current_quarter_end(start_year, start_month))
        start_year, start_month = next_quarter_start(start_year, start_month)


def account_from_path(top_account, account_path, original_path=None):
    """
    recursive function to get a Gnucash Account: starting from the top account and following the path
    :param   top_account: Gnucash Account: base account
    :param  account_path:            list: path to follow
    :param original_path:            list: original call path
    :return: Gnucash Account
    """
    # print("top_account = %s, account_path = %s, original_path = %s" % (top_account, account_path, original_path))
    if original_path is None:
        original_path = account_path
    account, account_path = account_path[0], account_path[1:]
    # print("account = %s, account_path = %s" % (account, account_path))

    account = top_account.lookup_by_name(account)
    # print("account = " + str(account))
    if account is None:
        raise Exception("Path '" + str(original_path) + "' could NOT be found!")
    if len(account_path) > 0:
        return account_from_path(account, account_path, original_path)
    else:
        return account


def save_to_json(fname, t_str, json_data, indt=4):
    """
    print json data to a file -- add a time string to get a unique file name each run
    :param     fname: string with file path and name
    :param     t_str: string with timestamp to use
    :param json_data: json compatible struct
    :param      indt: int: indentation amount
    :return: string with file name
    """
    out_file = fname + '_' + t_str + ".json"
    print_info("json file is '{}'\n".format(out_file))
    fp = open(out_file, 'w')
    json.dump(json_data, fp, indent=indt)
    fp.close()
    return out_file


def get_splits(acct, period_starts, period_list):
    """
    get the splits for the account and each sub-account
    :param          acct: Gnucash Account
    :param period_starts:            list: start date for each period
    :param   period_list: list of structs: store the dates and amounts for each quarter
    :return: nil
    """
    # insert and add all splits in the periods of interest
    for split in acct.GetSplitList():
        trans = split.parent
        # GetDate() returns a datetime but need a date
        trans_date = trans.GetDate().date()

        # use binary search to find the period that starts before or on the transaction date
        period_index = bisect_right(period_starts, trans_date) - 1

        # ignore transactions with a date before the matching period start and after the last period_end
        if period_index >= 0 and trans_date <= period_list[len(period_list) - 1][1]:
            # get the period bucket appropriate for the split in question
            period = period_list[period_index]
            assert( period[1] >= trans_date >= period[0] )

            split_amount = gnc_numeric_to_python_decimal(split.GetAmount())

            # if the amount is negative this is a credit, else a debit
            debit_credit_offset = 1 if split_amount < ZERO else 0

            # add the debit or credit to the sum, using the offset to get in the right bucket
            period[2 + debit_credit_offset] += split_amount

            # add the debit or credit to the overall total
            period[4] += split_amount


def fill_splits(root_acct, target_path, period_starts, period_list):
    """
    fill the period list for each account
    :param     root_acct: Gnucash Account: from the Gnucash book
    :param   target_path:            list: account hierarchy from root account to target account
    :param period_starts:            list: start date for each period
    :param   period_list: list of structs: store the dates and amounts for each quarter
    :return: string with name of target_acct
    """
    account_of_interest = account_from_path(root_acct, target_path)
    acct_name = account_of_interest.GetName()
    print_info("\naccount_of_interest = {}".format(acct_name), BLUE)

    # get the split amounts for the parent account
    get_splits(account_of_interest, period_starts, period_list)
    descendants = account_of_interest.get_descendants()
    if len(descendants) > 0:
        # for EACH sub-account add to the overall total
        # print("Descendants of {}:".format(account_of_interest.GetName()))
        for subAcct in descendants:
            # print("{} balance = {}".format(subAcct.GetName(), gnc_numeric_to_python_decimal(subAcct.GetBalance())))
            get_splits(subAcct, period_starts, period_list)

    csv_write_period_list(period_list)
    return acct_name


def csv_write_period_list(period_list):
    """
    Write out the details of the submitted period list in csv format
    :param period_list: list of structs: store the dates and amounts for each quarter
    :return: nil
    """
    # write out the column headers
    csv_writer = csv.writer(stdout)
    # csv_writer.writerow('')
    csv_writer.writerow(('period start', 'period end', 'debits', 'credits', 'TOTAL'))

    # write out the overall totals for the account of interest
    for start_date, end_date, debit_sum, credit_sum, total in period_list:
        csv_writer.writerow((start_date, end_date, debit_sum, credit_sum, total))
