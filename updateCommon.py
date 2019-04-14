#
# updateCommon.py -- common methods and variables for updates
#
# some code from account_analysis.py by Mark Jenkins, ParIT Worker Co-operative <mark@parit.ca>
# some code from Google quickstart spreadsheets example
#
# @author Mark Sattolo <epistemik@gmail.com>
# @version Python 3.6
# @created 2019-04-07
# @updated 2019-04-07

from sys import stdout
from datetime import date, timedelta
from bisect import bisect_right
from decimal import Decimal
from math import log10
import csv
from gnucash import GncNumeric
import json
import pickle
import os.path as osp
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
    print(inspect_line + color + text + COLOR_OFF, end=('\n' if newline else ''))


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
    print(inspect_line + RED + text + COLOR_OFF, end=('\n' if newline else ''))


# constant strings
QTR = 'Quarter'

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
BALANCE_SHEET    = 'Balance 1'
QTR_ASTS_SHEET   = 'Assets 1'
QTR_ASTS_2_SHEET = 'Assets 2'
ML_WORK_SHEET    = 'ML Work'
CALCULNS_SHEET   = 'Calculations'

TOKEN = SHEETS_EPISTEMIK_RW_TOKEN['P4']

# base cell (Q1) locations in Budget-qtrly.gsht
BASE_ROW = 3


def get_budget_id():
    fp = open(BUDGET_QTRLY_ID_FILE, "r")
    fid = fp.readline().strip()
    print_info("\nBudget Id = '{}'".format(fid), CYAN)
    fp.close()
    return fid


def get_credentials():
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


def next_period_start(start_year, start_month):
    # add number of months for a Quarter
    end_month = start_month + PERIOD_QTR

    # use integer division to find out if the new end month is in a different year,
    # what year it is, and what the end month number should be changed to.
    end_year = start_year + ((end_month - 1) // NUM_MONTHS)
    end_month = ((end_month - 1) % NUM_MONTHS) + 1

    return end_year, end_month


def period_end(start_year, start_month):
    end_year, end_month = next_period_start(start_year, start_month)

    # last step, the end date is one day back from the start of the next period
    # so we get a period end like 2010-03-31 instead of 2010-04-01
    return date(end_year, end_month, 1) - ONE_DAY


def generate_period_boundaries(start_year, start_month, periods):
    for i in range(periods):
        yield( date(start_year, start_month, 1), period_end(start_year, start_month) )
        start_year, start_month = next_period_start(start_year, start_month)


def account_from_path(top_account, account_path, original_path=None):
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
    :param fname: string
    :param t_str: string
    :param json_data: json compatible struct
    :param indt: int: indentation amount
    :return: file name
    """
    out_file = fname + '_' + t_str + ".json"
    print_info("\njson file is '{}'".format(out_file))
    fp = open(out_file, 'w')
    json.dump(json_data, fp, indent=indt)
    fp.close()
    return out_file


def get_splits(acct, period_starts, period_list):
    """
    get the splits for the account and each sub-account
    :param acct: Gnucash account
    :param period_starts: list: start date for each period
    :param period_list: list of structs: store the dates and amounts for each quarter
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
    :param root_acct: Gnucash account: from the Gnucash book
    :param target_path: list: account hierarchy from root account to target account
    :param period_starts: list: start date for each period
    :param period_list: list of structs: store the dates and amounts for each quarter
    :return: acct_name: string: name of target_acct
    """
    account_of_interest = account_from_path(root_acct, target_path)
    acct_name = account_of_interest.GetName()
    print("\naccount_of_interest = {}".format(acct_name))

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
