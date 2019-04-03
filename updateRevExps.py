#
# updateRevExps.py -- use the Gnucash and Google APIs to update my BudgetQtrly document for a specified year
#                     in the 'All Inc' and 'Nec Inc' sheets
#
# some code from account_analysis.py by Mark Jenkins, ParIT Worker Co-operative <mark@parit.ca>
#
# @author Mark Sattolo <epistemik@gmail.com>
# @version Python 3.6
# @created 2019-03-30
# @updated 2019-03-31

from sys import argv, stdout
from datetime import date, timedelta, datetime
from bisect import bisect_right
from decimal import Decimal
from math import log10
import csv
from gnucash import Session, GncNumeric

import pickle
import os.path as osp
import datetime as dt
import json
import copy
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# constant strings
QTR   = 'Quarter'
YR    = 'Year'
REV   = 'Revenue'
INV   = 'Invest'
OTH   = 'Other'
SAL   = 'Salary'
BAL   = 'Balance'
CONT  = 'Contingent'
NEC   = 'Necessary'
DEDNS = 'Sal_Dedns'

# find the proper path to the accounts in the gnucash file
REV_ACCTS = {
    INV : ["REV_Invest"],
    OTH : ["REV_Other"],
    SAL : ["REV_Salary"]
}
EXP_ACCTS = {
    BAL   : ["EXP_Balance"],
    CONT  : ["EXP_CONTINGENT"],
    NEC   : ["EXP_NECESSARY"],
    DEDNS : ["EXP_Salary"]
}

# either for One Quarter or for Four Quarters if updating an entire Year
results = list()
# store the values needed to update the document
REV_EXP_RESULTS = {
    QTR   : '0',
    REV   : '0',
    BAL   : '0',
    CONT  : '0',
    NEC   : '0',
    DEDNS : '0'
}

# a dictionary with a period name as key, and number of months in that kind of period as the value
PERIODS = {
    QTR :  3 ,
    YR  : 12
}

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
BUDGET_QTRLY_SPRD_SHEET = '1YbHb7RjZUlA2gyaGDVgRoQYhjs9I8gndKJ0f1Cn-Zr0'
# sheet names in Budget Quarterly
ALL_INC_SHEET      = 'All Inc Quarterly'
ALL_INC_PRAC_SHEET = 'All Inc Practice'
NEC_INC_SHEET      = 'Nec Inc Quarterly'
NEC_INC_PRAC_SHEET = 'Nec Inc Practice'
BALANCE_SHEET      = 'Balance Sheet'
QTR_ASTS_SHEET     = 'Quarterly Assets'
ML_WORK_SHEET      = 'ML Work'
CALCULNS_SHEET     = 'Calculations'

TOKEN = SHEETS_EPISTEMIK_RW_TOKEN['P4']

now = dt.datetime.strftime(dt.datetime.now(), "%Y-%m-%dT%H-%M-%S")

# for One Quarter or for Four Quarters if updating an entire Year
data = list()
cell_data = {
    # sample data
    'range': 'Calculations!P47',
    'values': [ [ '$9033.66' ] ]
}

# base cell (2012-Q1) locations in Budget-qtrly.gsht
BASE_ROW = 3
REV_EXP_COLS = {
    REV   : 'D',
    BAL   : 'P',
    CONT  : 'O',
    NEC   : 'G',
    DEDNS : 'D'
}
BASE_YEAR = 2012
# number of rows between quarters in the same year
QTR_SPAN = 2
# number of rows between years
YEAR_SPAN = 11


# noinspection PyUnresolvedReferences
def gnc_numeric_to_python_decimal(numeric):
    negative = numeric.negative_p()
    sign = 1 if negative else 0

    copy = GncNumeric(numeric.num(), numeric.denom())
    result = copy.to_decimal(None)
    if not result:
        raise Exception("GncNumeric value '{}' CANNOT be converted to decimal!".format(copy.to_string()))

    digit_tuple = tuple(int(char) for char in str(copy.num()) if char != '-')
    denominator = copy.denom()
    exponent = int(log10(denominator))
    assert( (10 ** exponent) == denominator )
    return Decimal((sign, digit_tuple, -exponent))


def next_period_start(start_year, start_month, period_type):
    # add numbers of months for the period length
    end_month = start_month + PERIODS[period_type]

    # use integer division to find out if the new end month is in a different year,
    # what year it is, and what the end month number should be changed to.
    end_year = start_year + ((end_month - 1) // NUM_MONTHS)
    end_month = ((end_month - 1) % NUM_MONTHS) + 1

    return end_year, end_month


def period_end(start_year, start_month, period_type):
    if period_type not in PERIODS:
        raise Exception("'{}' is NOT a valid period >> MUST be one of '{}'!".format(period_type, str(PERIODS.keys())))

    end_year, end_month = next_period_start(start_year, start_month, period_type)

    # last step, the end date is one day back from the start of the next period
    # so we get a period end like 2010-03-31 instead of 2010-04-01
    return date(end_year, end_month, 1) - ONE_DAY


def generate_period_boundaries(start_year, start_month, period_type, periods):
    for i in range(periods):
        yield( date(start_year, start_month, 1), period_end(start_year, start_month, period_type) )
        start_year, start_month = next_period_start(start_year, start_month, period_type)


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


def get_splits(acct, period_starts, period_list):
    # insert and add all splits in the periods of interest
    for split in acct.GetSplitList():
        trans = split.parent
        # GetDate() returns a datetime
        tx_datetm = trans.GetDate()
        # convert to a date
        trans_date = tx_datetm.date()

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


def csv_write_period_list(period_list):
    # write out the column headers
    csv_writer = csv.writer(stdout)
    # csv_writer.writerow('')
    csv_writer.writerow(('period start', 'period end', 'debits', 'credits', 'TOTAL'))

    # write out the overall totals for the account of interest
    for start_date, end_date, debit_sum, credit_sum, total in period_list:
        csv_writer.writerow((start_date, end_date, debit_sum, credit_sum, total))


def get_revenue(root_account, period_starts, period_list, re_year, qtr):
    """
    Get REVENUE data for the specified Quarter
    :param root_account: string: root account of the gnucash book
    :param period_starts: struct: store the dates and amounts for each quarter
    :param period_list: list: start date for each period
    :param re_year: int: year to read
    :param qtr: int: quarter to read: 1-4
    """
    str_rev = '= '
    for item in REV_ACCTS:
        # reset the debit and credit totals for each individual account
        period_list[0][2] = 0
        period_list[0][3] = 0

        acct_base = REV_ACCTS[item]
        # print("acct = {}".format(acct_base))

        account_of_interest = account_from_path(root_account, acct_base)
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

        sum_revenue = (period_list[0][2] + period_list[0][3]) * (-1)
        str_rev += sum_revenue.to_eng_string() + (' + ' if item != SAL else '')
        print("{} Revenue for {}-Q{} = ${}".format(acct_name, re_year, qtr, sum_revenue))

    REV_EXP_RESULTS[REV] = str_rev


def get_expenses(root_account, period_starts, period_list, re_year, qtr):
    """
    Get EXPENSE data for the specified Quarter
    :param root_account: string: root account of the gnucash book
    :param period_starts: struct: store the dates and amounts for each quarter
    :param period_list: list: start date for each period
    :param re_year: int: year to read
    :param qtr: int: quarter to read: 1-4
    """
    for item in EXP_ACCTS:
        # reset the debit and credit totals for each individual account
        period_list[0][2] = 0
        period_list[0][3] = 0

        acct_base = EXP_ACCTS[item]
        # print("acct = {}".format(acct_base))

        account_of_interest = account_from_path(root_account, acct_base)
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

        sum_expenses = (period_list[0][2] + period_list[0][3])
        REV_EXP_RESULTS[item] = sum_expenses.to_eng_string()
        print("{} Expenses for {}-Q{} = ${}".format(acct_name.split('_')[-1], re_year, qtr, sum_expenses))


# noinspection PyUnboundLocalVariable,PyUnresolvedReferences
def get_rev_exps(gnucash_file, re_year, re_quarter):
    """
    Get revenue and expense data for ONE specified Quarter or ALL four Quarters for the specified Year
    :param gnucash_file: string: name of file used to read the values
    :param re_year: int: year to update
    :param re_quarter: int: 1-4 for quarter to update or 0 if updating entire year
    """
    num_quarters = 1 if re_quarter else 4
    print("find Revenue & Expenses in {} for {}{}".format(gnucash_file, re_year, ('-Q' + str(re_quarter)) if re_quarter else ''))

    try:
        gnucash_session = Session(gnucash_file, is_new=False)
        root_account = gnucash_session.book.get_root_account()

        for i in range(num_quarters):
            qtr = re_quarter if re_quarter else i + 1
            REV_EXP_RESULTS[QTR] = str(qtr)

            start_month = (qtr * 3) - 2

            # for each period keep the start date, end date, debits and credits sums and overall total
            period_list = [
                [
                    start_date, end_date,
                    ZERO, # debits sum
                    ZERO, # credits sum
                    ZERO  # TOTAL
                ]
                for start_date, end_date in generate_period_boundaries(re_year, start_month, QTR, 1)
            ]
            # print(period_list)
            # a copy of the above list with just the period start dates
            period_starts = [e[0] for e in period_list]
            # print(period_starts)

            get_revenue(root_account, period_starts, period_list, re_year, qtr)
            tot_revenue = period_list[0][4] * (-1)
            print("\n{} Revenue for {}-Q{} = ${}".format("TOTAL", re_year, qtr, tot_revenue))

            period_list[0][4] = 0
            get_expenses(root_account, period_starts, period_list, re_year, qtr)
            tot_expenses = period_list[0][4]
            print("\n{} Expenses for {}-Q{} = ${}\n".format("TOTAL", re_year, qtr, tot_expenses))

            results.append(copy.deepcopy(REV_EXP_RESULTS))
            period_list[0][4] = 0
            print(json.dumps(REV_EXP_RESULTS, indent=4))

        # no save needed, we're just reading...
        gnucash_session.end()

    except Exception as qe:
        print("Exception: {}!".format(qe))
        if "gnucash_session" in locals() and gnucash_session is not None:
            gnucash_session.end()


def fill_rev_exps_data(all_inc_dest, nec_inc_dest, re_year):
    """
    for each item in results, either 1 for one quarter or 4 for four quarters:
    create 5 cell_data's, one each for REV, BAL, CONT, NEC, DEDNS:
    fill in the range based on the year and quarter
    range = SHEET_NAME + '!' + calculated cell
    fill in the values based on the sheet being updated and the type of cell_data
    REV string is '= ${INV} + ${OTH} + ${SAL}'
    others are just the string from the item
    :param all_inc_dest: string: sheet location for REV, CONT & BAL
    :param nec_inc_dest: string: sheet location for DEDNS & NEC
    :param re_year: int: year to update
    """
    print("\nfill_rev_exps_data({}, {}, {})\n".format(all_inc_dest, nec_inc_dest, re_year))

    year_location = (re_year - BASE_YEAR) * YEAR_SPAN
    # get exact row from qtr value in each item
    for item in results:
        int_qtr = int(item[QTR])
        qtr_location = year_location + ((int_qtr - 1) * QTR_SPAN)
        print("{} = {}\n".format(QTR, item[QTR]))


def send_rev_exps(all_inc_dest, nec_inc_dest, re_year):
    """
    Send the data to the document
    :param all_inc_dest: string: sheet location for REV, CONT & BAL
    :param nec_inc_dest: string: sheet location for DEDNS & NEC
    :param re_year: int: year to update
    """
    print("\nsend_rev_exps({}, {}, {})".format(all_inc_dest, nec_inc_dest, re_year))
    print("cell_data['range'] = {}".format(cell_data['range']))
    print("cell_data['values'][0][0] = {}\n".format(cell_data['values'][0][0]))

    fill_rev_exps_data(all_inc_dest, nec_inc_dest, re_year)

    return

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if osp.exists(TOKEN):
        with open(TOKEN, 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS, SHEETS_RW_SCOPE)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open(TOKEN, 'wb') as token:
            pickle.dump(creds, token, pickle.HIGHEST_PROTOCOL)

    # Call the Sheets API
    service = build('sheets', 'v4', credentials=creds)
    srv_sheets = service.spreadsheets()

    my_body = {
        'valueInputOption': 'USER_ENTERED',
        'data': data
    }
    vals = srv_sheets.values()
    response = vals.batchUpdate(spreadsheetId=BUDGET_QTRLY_SPRD_SHEET, body=my_body).execute()

    print('{} cells updated!'.format(response.get('totalUpdatedCells')))
    print(json.dumps(response, indent=4))


def update_rev_exps_main():
    exe = argv[0].split('/')[-1]
    if len(argv) < 4:
        print("NOT ENOUGH parameters!")
        print("usage: {} <book url> <mode=prod|test> <year> [quarter]".format(exe))
        print("PROGRAM EXIT!")
        return

    print("\nrunning {} at run-time: {}\n".format(exe, str(datetime.now())))

    gnucash_file = argv[1]
    
    all_inc_dest = ALL_INC_PRAC_SHEET
    nec_inc_dest = NEC_INC_PRAC_SHEET
    if argv[2].lower() == 'prod':
        all_inc_dest = ALL_INC_SHEET
        nec_inc_dest = NEC_INC_SHEET
        
    re_year = int(argv[3])
    re_quarter = int(argv[4]) if len(argv) > 4 else 0

    get_rev_exps(gnucash_file, re_year, re_quarter)
    print('\nresults:')
    print(json.dumps(results, indent=4))

    send_rev_exps(all_inc_dest, nec_inc_dest, re_year)

    print("\n >>> PROGRAM ENDED.")


if __name__ == "__main__":
    update_rev_exps_main()
